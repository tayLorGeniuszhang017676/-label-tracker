"""
📦 面单识别 → Tracking 回填工具
Upload shipping labels → AI extracts tracking # & recipient → updates Excel
"""

import streamlit as st
import pandas as pd
import base64
import json
import re
import io
import time
from difflib import SequenceMatcher
from pathlib import Path

try:
    import anthropic
except ImportError:
    anthropic = None

try:
    import openpyxl
except ImportError:
    openpyxl = None

# ── Page config ──────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="面单识别 → Tracking 回填",
    page_icon="📦",
    layout="wide",
)

# ── Custom CSS ───────────────────────────────────────────────────────────────

st.markdown("""
<style>
    .stApp { }
    div[data-testid="stMetric"] {
        background: #f8fafc;
        border: 1px solid #e2e8f0;
        border-radius: 8px;
        padding: 12px 16px;
    }
    .match-high { color: #16a34a; font-weight: 600; }
    .match-mid  { color: #ca8a04; font-weight: 600; }
    .match-low  { color: #dc2626; font-weight: 600; }
</style>
""", unsafe_allow_html=True)

# ── Session state init ───────────────────────────────────────────────────────

if "extracted_labels" not in st.session_state:
    st.session_state.extracted_labels = []
if "match_results" not in st.session_state:
    st.session_state.match_results = []
if "excel_df" not in st.session_state:
    st.session_state.excel_df = None
if "excel_bytes" not in st.session_state:
    st.session_state.excel_bytes = None

# ── Helper functions ─────────────────────────────────────────────────────────

EXTRACT_PROMPT = """You are a shipping label reader. Extract the following from this shipping label image:
1. Recipient name (the name after "SHIP TO:" or the main recipient name)
2. Tracking number (the number after "TRACKING #:" or similar tracking identifier, usually starts with 1Z for UPS)

Respond ONLY in this exact JSON format, no markdown, no backticks, no extra text:
{"recipient_name": "FULL NAME HERE", "tracking_number": "TRACKING_NUMBER_HERE"}

Rules:
- Remove all spaces from the tracking number (e.g. "1Z 245 1R8 29 1876 2788" becomes "1Z2451R8291876 2788" with no spaces)
- Keep the recipient name exactly as shown on the label
- If you cannot find a field, use null for that field"""


def extract_from_label(client: "anthropic.Anthropic", image_bytes: bytes, media_type: str) -> dict:
    """Use Claude Vision to extract tracking info from a label image."""
    b64 = base64.standard_b64encode(image_bytes).decode("utf-8")

    try:
        response = client.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=500,
            messages=[
                {
                    "role": "user",
                    "content": [
                        {
                            "type": "image",
                            "source": {
                                "type": "base64",
                                "media_type": media_type,
                                "data": b64,
                            },
                        },
                        {"type": "text", "text": EXTRACT_PROMPT},
                    ],
                }
            ],
        )
        text = "".join(block.text for block in response.content if hasattr(block, "text"))
        clean = re.sub(r"```json|```", "", text).strip()
        result = json.loads(clean)
        # Normalize tracking number: remove all spaces
        if result.get("tracking_number"):
            result["tracking_number"] = result["tracking_number"].replace(" ", "")
        return result
    except json.JSONDecodeError:
        return {"recipient_name": None, "tracking_number": None, "error": f"JSON parse failed: {text[:200]}"}
    except Exception as e:
        return {"recipient_name": None, "tracking_number": None, "error": str(e)}


def normalize_name(name: str) -> str:
    """Normalize a name for comparison."""
    if not name:
        return ""
    return re.sub(r"\s+", " ", str(name).strip().upper())


def name_similarity(a: str, b: str) -> float:
    """Fuzzy match two names using SequenceMatcher."""
    na, nb = normalize_name(a), normalize_name(b)
    if not na or not nb:
        return 0.0
    if na == nb:
        return 1.0
    return SequenceMatcher(None, na, nb).ratio()


def find_best_match(label_name: str, names_series: pd.Series, threshold: float = 0.55) -> tuple:
    """Find the best matching row index for a label name."""
    best_idx, best_score, best_name = -1, 0.0, ""
    for idx, name in names_series.items():
        if pd.isna(name) or not str(name).strip():
            continue
        score = name_similarity(label_name, str(name))
        if score > best_score:
            best_score = score
            best_idx = idx
            best_name = str(name)
    if best_score >= threshold:
        return best_idx, best_score, best_name
    return -1, best_score, best_name


def get_media_type(filename: str) -> str:
    """Get MIME type from filename."""
    ext = Path(filename).suffix.lower()
    return {
        ".png": "image/png",
        ".jpg": "image/jpeg",
        ".jpeg": "image/jpeg",
        ".gif": "image/gif",
        ".webp": "image/webp",
        ".bmp": "image/bmp",
    }.get(ext, "image/png")


# ── Sidebar: API Key ────────────────────────────────────────────────────────

with st.sidebar:
    st.markdown("### ⚙️ 设置")
    api_key = st.text_input(
        "Anthropic API Key",
        type="password",
        help="在 console.anthropic.com 获取 API Key",
        placeholder="sk-ant-...",
    )

    st.divider()
    st.markdown("### 📖 使用说明")
    st.markdown("""
    1. 输入 API Key
    2. 上传面单图片（支持批量）
    3. 点击「开始识别」
    4. 上传 ParcelOutbound Excel
    5. 确认匹配结果
    6. 下载更新后的 Excel
    """)

    st.divider()
    st.markdown(
        "<div style='color:#94a3b8;font-size:12px'>Built for 海势云帆 出海业务</div>",
        unsafe_allow_html=True,
    )

# ── Main UI ──────────────────────────────────────────────────────────────────

st.markdown("# 📦 面单识别 → Tracking 回填")
st.markdown("上传 UPS/FedEx 面单图片，AI 自动识别 Tracking Number，按收件人匹配回填到出库单 Excel")

# ── Step 1: Upload & extract labels ─────────────────────────────────────────

st.markdown("---")
st.markdown("### ① 上传面单图片")

uploaded_labels = st.file_uploader(
    "选择面单图片（支持 PNG / JPG，可多选）",
    type=["png", "jpg", "jpeg", "webp", "bmp"],
    accept_multiple_files=True,
    key="label_uploader",
)

if uploaded_labels:
    # Preview thumbnails
    cols = st.columns(min(len(uploaded_labels), 5))
    for i, f in enumerate(uploaded_labels[:5]):
        with cols[i]:
            st.image(f, caption=f.name, width=120)
    if len(uploaded_labels) > 5:
        st.caption(f"... 还有 {len(uploaded_labels) - 5} 张")

    # Extract button
    if not api_key:
        st.warning("⬅ 请先在左侧输入 API Key")
    else:
        if st.button("🔍 开始识别", type="primary", use_container_width=True):
            if anthropic is None:
                st.error("请安装 anthropic: `pip install anthropic`")
                st.stop()

            client = anthropic.Anthropic(api_key=api_key)
            results = []
            progress = st.progress(0, text="正在识别面单...")

            for i, label_file in enumerate(uploaded_labels):
                progress.progress(
                    (i + 1) / len(uploaded_labels),
                    text=f"正在识别 {label_file.name} ({i+1}/{len(uploaded_labels)})",
                )
                image_bytes = label_file.getvalue()
                media_type = get_media_type(label_file.name)
                extracted = extract_from_label(client, image_bytes, media_type)
                results.append(
                    {
                        "文件名": label_file.name,
                        "收件人": extracted.get("recipient_name", ""),
                        "Tracking #": extracted.get("tracking_number", ""),
                        "状态": "❌ " + extracted.get("error", "") if extracted.get("error") else "✅ 成功",
                    }
                )
                # Small delay to avoid rate limit
                if i < len(uploaded_labels) - 1:
                    time.sleep(0.3)

            progress.empty()
            st.session_state.extracted_labels = results
            st.success(f"识别完成！成功 {sum(1 for r in results if r['状态'].startswith('✅'))}/{len(results)} 张")

# Show extraction results
if st.session_state.extracted_labels:
    st.markdown("**识别结果：**")
    df_labels = pd.DataFrame(st.session_state.extracted_labels)
    st.dataframe(
        df_labels,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Tracking #": st.column_config.TextColumn(width="large"),
        },
    )

# ── Step 2: Upload Excel ────────────────────────────────────────────────────

if st.session_state.extracted_labels:
    st.markdown("---")
    st.markdown("### ② 上传 ParcelOutbound Excel")

    uploaded_excel = st.file_uploader(
        "选择 ParcelOutbound Excel 文件",
        type=["xlsx", "xls"],
        key="excel_uploader",
    )

    if uploaded_excel:
        excel_bytes = uploaded_excel.getvalue()
        st.session_state.excel_bytes = excel_bytes

        # Read with openpyxl data_only to get cached formula values
        wb = openpyxl.load_workbook(io.BytesIO(excel_bytes), data_only=True)
        main_sheet = wb.sheetnames[0]  # 小包出库单
        ws = wb[main_sheet]

        # Find headers
        headers = {}
        for col in range(1, ws.max_column + 1):
            val = ws.cell(row=1, column=col).value
            if val:
                headers[val] = col

        recipient_col = None
        tracking_col = None
        for key, col in headers.items():
            if "收件人" in str(key) or "RecipientName" in str(key):
                recipient_col = col
            if "Tracking" in str(key) or "跟踪号" in str(key):
                tracking_col = col

        if not recipient_col or not tracking_col:
            st.error("找不到收件人或 Tracking 列，请确认 Excel 格式")
            st.stop()

        # Build dataframe from cached values
        rows_data = []
        for row in range(2, ws.max_row + 1):
            name_val = ws.cell(row=row, column=recipient_col).value
            tracking_val = ws.cell(row=row, column=tracking_col).value
            if name_val:  # skip empty rows
                rows_data.append(
                    {
                        "excel_row": row,
                        "收件人": str(name_val).strip(),
                        "现有 Tracking": tracking_val or "",
                    }
                )

        df_excel = pd.DataFrame(rows_data)
        st.session_state.excel_df = df_excel
        st.session_state.recipient_col = recipient_col
        st.session_state.tracking_col = tracking_col
        st.session_state.main_sheet = main_sheet

        st.success(f"已加载 **{len(df_excel)}** 条订单（Sheet: {main_sheet}）")

        # ── Step 3: Matching ─────────────────────────────────────────────────

        st.markdown("---")
        st.markdown("### ③ 匹配结果")

        match_results = []
        for label in st.session_state.extracted_labels:
            if not label["状态"].startswith("✅") or not label["收件人"]:
                continue

            best_idx, best_score, best_name = find_best_match(
                label["收件人"], df_excel["收件人"]
            )

            match_results.append(
                {
                    "面单收件人": label["收件人"],
                    "匹配到 Excel": best_name if best_idx >= 0 else "—",
                    "相似度": best_score,
                    "Tracking": label["Tracking #"],
                    "excel_row": df_excel.iloc[best_idx]["excel_row"] if best_idx >= 0 else -1,
                    "现有 Tracking": df_excel.iloc[best_idx]["现有 Tracking"] if best_idx >= 0 else "",
                    "接受": best_score >= 0.7,
                }
            )

        if match_results:
            st.session_state.match_results = match_results

            # Summary metrics
            n_total = len(match_results)
            n_high = sum(1 for m in match_results if m["相似度"] >= 0.85)
            n_mid = sum(1 for m in match_results if 0.7 <= m["相似度"] < 0.85)
            n_low = sum(1 for m in match_results if m["相似度"] < 0.7)

            c1, c2, c3, c4 = st.columns(4)
            c1.metric("总识别", n_total)
            c2.metric("高匹配 ≥85%", n_high)
            c3.metric("中匹配 70-84%", n_mid)
            c4.metric("低匹配 <70%", n_low)

            # Editable table
            st.markdown("**点击「接受」列可切换是否回填该条 Tracking：**")

            df_match = pd.DataFrame(match_results)
            edited = st.data_editor(
                df_match[["接受", "面单收件人", "匹配到 Excel", "相似度", "Tracking", "现有 Tracking"]],
                use_container_width=True,
                hide_index=True,
                disabled=["面单收件人", "匹配到 Excel", "相似度", "Tracking", "现有 Tracking"],
                column_config={
                    "接受": st.column_config.CheckboxColumn("✓ 接受", default=False),
                    "相似度": st.column_config.ProgressColumn("相似度", min_value=0, max_value=1, format="%.0f%%"),
                    "Tracking": st.column_config.TextColumn(width="large"),
                },
            )

            # Update accepted states from the editor
            accepted_mask = edited["接受"].tolist()
            for i, acc in enumerate(accepted_mask):
                match_results[i]["接受"] = acc

            # ── Step 4: Download ─────────────────────────────────────────────

            st.markdown("---")
            st.markdown("### ④ 下载更新后的 Excel")

            n_accepted = sum(1 for m in match_results if m["接受"])

            if n_accepted == 0:
                st.info("没有接受的匹配项，请在上方勾选要回填的条目")
            else:
                st.markdown(f"将回填 **{n_accepted}** 条 Tracking Number 到 Excel")

                if st.button(f"⬇️ 生成并下载更新后的 Excel（{n_accepted} 条）", type="primary", use_container_width=True):
                    # Load the FORMULA workbook (not data_only) so we preserve formulas
                    wb_write = openpyxl.load_workbook(io.BytesIO(st.session_state.excel_bytes))
                    ws_write = wb_write[st.session_state.main_sheet]

                    filled_count = 0
                    for m in match_results:
                        if m["接受"] and m["excel_row"] > 0 and m["Tracking"]:
                            ws_write.cell(
                                row=m["excel_row"],
                                column=st.session_state.tracking_col,
                            ).value = m["Tracking"]
                            filled_count += 1

                    # Save to bytes
                    output = io.BytesIO()
                    wb_write.save(output)
                    output.seek(0)

                    st.download_button(
                        label=f"📥 下载 ParcelOutbound_Updated.xlsx（已填 {filled_count} 条 Tracking）",
                        data=output.getvalue(),
                        file_name="ParcelOutbound_Updated.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )
                    st.success(f"✅ 已生成！共填入 {filled_count} 条 Tracking Number")
        else:
            st.warning("没有可匹配的识别结果")
