"""
📦 面单识别 → Tracking 回填工具
Upload shipping labels → OCR extracts tracking # & recipient → updates Excel
Free: no API key needed, uses Tesseract OCR
"""

import streamlit as st
import pandas as pd
import re
import io
from difflib import SequenceMatcher
from pathlib import Path
from PIL import Image, ImageFilter, ImageEnhance
import pytesseract

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

st.markdown("""
<style>
    div[data-testid="stMetric"] {
        background: #f8fafc;
        border: 1px solid #e2e8f0;
        border-radius: 8px;
        padding: 12px 16px;
    }
</style>
""", unsafe_allow_html=True)

# ── Session state ────────────────────────────────────────────────────────────

for key in ["extracted_labels", "match_results"]:
    if key not in st.session_state:
        st.session_state[key] = []
for key in ["excel_bytes"]:
    if key not in st.session_state:
        st.session_state[key] = None

# ── OCR + Extraction ─────────────────────────────────────────────────────────

# Tracking number patterns
TRACKING_PATTERNS = [
    # UPS: 1Z + 6 alphanumeric + 2 digits + 8 digits (18 chars total)
    r'1Z\s*[A-Z0-9]{3}\s*[A-Z0-9]{3}\s*[A-Z0-9]{2}\s*[A-Z0-9]{4}\s*[A-Z0-9]{4}',
    # UPS alternative
    r'1Z[A-Z0-9\s]{16,22}',
    # FedEx: 12 or 15 digits
    r'\b\d{12,15}\b',
    # USPS: 20-22 digits
    r'\b\d{20,22}\b',
]

def preprocess_image(img: Image.Image) -> Image.Image:
    """Enhance image for better OCR accuracy."""
    # Convert to grayscale
    img = img.convert("L")
    # Increase contrast
    enhancer = ImageEnhance.Contrast(img)
    img = enhancer.enhance(2.0)
    # Sharpen
    img = img.filter(ImageFilter.SHARPEN)
    # Scale up small images
    w, h = img.size
    if w < 1000:
        scale = 2
        img = img.resize((w * scale, h * scale), Image.LANCZOS)
    return img


def extract_tracking(text: str) -> str:
    """Extract tracking number from OCR text."""
    # Look for explicit "TRACKING" label first
    tracking_match = re.search(
        r'TRACKING\s*#?\s*:?\s*(1Z[\sA-Z0-9]+)',
        text, re.IGNORECASE
    )
    if tracking_match:
        clean = re.sub(r'\s+', '', tracking_match.group(1))
        # UPS tracking numbers are exactly 18 characters
        if clean.startswith('1Z') and len(clean) >= 18:
            return clean[:18]
        return clean

    # Try each pattern
    for pattern in TRACKING_PATTERNS:
        matches = re.findall(pattern, text, re.IGNORECASE)
        for m in matches:
            clean = re.sub(r'\s+', '', m)
            if clean.startswith('1Z') and len(clean) >= 16:
                return clean

    return ""


def extract_recipient(text: str) -> str:
    """Extract recipient name from OCR text (after SHIP TO)."""
    # Pattern: SHIP TO: followed by name on next line(s)
    ship_match = re.search(
        r'SHIP\s*TO\s*:?\s*\n\s*(.+?)(?:\n|$)',
        text, re.IGNORECASE
    )
    if ship_match:
        name = ship_match.group(1).strip()
        # Clean up common OCR artifacts
        name = re.sub(r'[^\w\s\'-]', '', name).strip()
        if len(name) >= 2 and not name.isdigit():
            return name

    # Alternative: look for name pattern after SHIP TO block
    ship_block = re.search(
        r'SHIP\s*TO\s*:?(.*?)(?:(?:\d{1,5}\s+\w)|(?:APARTMENT|APT|UNIT|SUITE))',
        text, re.IGNORECASE | re.DOTALL
    )
    if ship_block:
        lines = [l.strip() for l in ship_block.group(1).strip().split('\n') if l.strip()]
        for line in lines:
            clean = re.sub(r'[^\w\s\'-]', '', line).strip()
            # Name is typically all letters, at least 2 chars
            if clean and not clean.isdigit() and len(clean) >= 2:
                # Skip if it looks like an address (starts with number)
                if not re.match(r'^\d', clean):
                    return clean

    return ""


def process_label(image_bytes: bytes) -> dict:
    """Process a single shipping label image."""
    try:
        img = Image.open(io.BytesIO(image_bytes))
        processed = preprocess_image(img)

        # Run OCR
        text = pytesseract.image_to_string(processed, config='--psm 6')

        tracking = extract_tracking(text)
        recipient = extract_recipient(text)

        return {
            "recipient_name": recipient or None,
            "tracking_number": tracking or None,
            "ocr_text": text,  # Keep for debugging
            "error": None if (tracking or recipient) else "无法从图片中提取信息，可能图片质量不够清晰",
        }
    except Exception as e:
        return {
            "recipient_name": None,
            "tracking_number": None,
            "ocr_text": "",
            "error": str(e),
        }


# ── Excel helpers ────────────────────────────────────────────────────────────

def normalize_name(name: str) -> str:
    if not name:
        return ""
    return re.sub(r"\s+", " ", str(name).strip().upper())


def name_similarity(a: str, b: str) -> float:
    na, nb = normalize_name(a), normalize_name(b)
    if not na or not nb:
        return 0.0
    if na == nb:
        return 1.0
    return SequenceMatcher(None, na, nb).ratio()


def find_best_match(label_name: str, names_list: list, threshold: float = 0.55):
    best_idx, best_score, best_name = -1, 0.0, ""
    for idx, name in enumerate(names_list):
        if not name or not str(name).strip():
            continue
        score = name_similarity(label_name, str(name))
        if score > best_score:
            best_score = score
            best_idx = idx
            best_name = str(name)
    if best_score >= threshold:
        return best_idx, best_score, best_name
    return -1, best_score, best_name


def read_excel_recipients(excel_bytes: bytes):
    """Read recipient names from Excel, handling VLOOKUP formulas."""
    wb = openpyxl.load_workbook(io.BytesIO(excel_bytes), data_only=True)
    main_sheet = wb.sheetnames[0]
    ws = wb[main_sheet]

    recipient_col = tracking_col = None
    for col in range(1, ws.max_column + 1):
        val = ws.cell(row=1, column=col).value
        if val:
            val_str = str(val)
            if "收件人" in val_str or "RecipientName" in val_str:
                recipient_col = col
            if "Tracking" in val_str or "跟踪号" in val_str:
                tracking_col = col

    if not recipient_col or not tracking_col:
        return None, None, None, None

    # Read cached values
    records = []
    for row in range(2, ws.max_row + 1):
        name_val = ws.cell(row=row, column=recipient_col).value
        tracking_val = ws.cell(row=row, column=tracking_col).value
        has_data = any(ws.cell(row=row, column=c).value for c in range(1, min(10, ws.max_column + 1)))
        if has_data:
            records.append({
                "excel_row": row,
                "收件人": str(name_val).strip() if name_val else "",
                "现有 Tracking": str(tracking_val).strip() if tracking_val else "",
            })

    # If no names found, try source sheet (for VLOOKUP cases)
    if all(r["收件人"] in ("", "None") for r in records):
        for sn in wb.sheetnames:
            if sn != main_sheet:
                ws2 = wb[sn]
                src_recipient_col = src_order_col = None
                for col in range(1, min(30, ws2.max_column + 1)):
                    val = ws2.cell(row=1, column=col).value
                    if val:
                        if "收件人" in str(val):
                            src_recipient_col = col
                        if "订单号" in str(val):
                            src_order_col = col

                if src_recipient_col and src_order_col:
                    order_to_name = {}
                    for row in range(2, ws2.max_row + 1):
                        order = ws2.cell(row=row, column=src_order_col).value
                        name = ws2.cell(row=row, column=src_recipient_col).value
                        if order and name:
                            order_to_name[str(order).strip()] = str(name).strip()

                    ws_main = wb[main_sheet]
                    platform_col = None
                    for col in range(1, ws_main.max_column + 1):
                        val = ws_main.cell(row=1, column=col).value
                        if val and ("Platform Number" in str(val) or "平台单号" in str(val)):
                            platform_col = col
                            break

                    if platform_col and order_to_name:
                        for rec in records:
                            order_num = ws_main.cell(row=rec["excel_row"], column=platform_col).value
                            if order_num and str(order_num).strip() in order_to_name:
                                rec["收件人"] = order_to_name[str(order_num).strip()]
                    break

    records = [r for r in records if r["收件人"] and r["收件人"] not in ("", "None")]
    return records, recipient_col, tracking_col, main_sheet


# ── UI ───────────────────────────────────────────────────────────────────────

st.markdown("# 📦 面单识别 → Tracking 回填")
st.markdown("上传面单图片，自动识别 Tracking Number，按收件人匹配回填到出库单")
st.caption("✅ 完全免费，无需 API Key")

# ── Step 1 ───────────────────────────────────────────────────────────────────

st.markdown("---")
st.markdown("### ① 上传面单图片")

uploaded_labels = st.file_uploader(
    "选择面单图片（PNG / JPG，可多选）",
    type=["png", "jpg", "jpeg", "webp", "bmp"],
    accept_multiple_files=True,
    key="label_uploader",
)

if uploaded_labels:
    cols = st.columns(min(len(uploaded_labels), 5))
    for i, f in enumerate(uploaded_labels[:5]):
        with cols[i]:
            st.image(f, caption=f.name, width=120)
    if len(uploaded_labels) > 5:
        st.caption(f"... 还有 {len(uploaded_labels) - 5} 张")

    if st.button("🔍 开始识别", type="primary", use_container_width=True):
        results = []
        progress = st.progress(0, text="正在识别面单...")

        for i, label_file in enumerate(uploaded_labels):
            progress.progress(
                (i + 1) / len(uploaded_labels),
                text=f"正在识别 {label_file.name} ({i+1}/{len(uploaded_labels)})",
            )
            extracted = process_label(label_file.getvalue())
            results.append({
                "文件名": label_file.name,
                "收件人": extracted.get("recipient_name") or "",
                "Tracking #": extracted.get("tracking_number") or "",
                "状态": f"❌ {extracted['error']}" if extracted.get("error") else "✅ 成功",
                "_ocr_text": extracted.get("ocr_text", ""),
            })

        progress.empty()
        st.session_state.extracted_labels = results
        n_ok = sum(1 for r in results if r["状态"].startswith("✅"))
        if n_ok > 0:
            st.success(f"识别完成！成功 {n_ok}/{len(results)} 张")
        else:
            st.error(f"识别失败，请检查图片清晰度")

# Show results
if st.session_state.extracted_labels:
    st.markdown("**识别结果：**")
    df_labels = pd.DataFrame(st.session_state.extracted_labels)
    display_cols = [c for c in df_labels.columns if not c.startswith("_")]
    st.dataframe(df_labels[display_cols], use_container_width=True, hide_index=True)

    # Debug: show OCR text
    with st.expander("🔍 调试：查看 OCR 原始文本"):
        for r in st.session_state.extracted_labels:
            st.markdown(f"**{r['文件名']}:**")
            st.code(r.get("_ocr_text", ""), language=None)

# ── Step 2 ───────────────────────────────────────────────────────────────────

successful_labels = [l for l in st.session_state.extracted_labels if l["状态"].startswith("✅")]
if successful_labels:
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

        records, recipient_col, tracking_col, main_sheet = read_excel_recipients(excel_bytes)

        if records is None:
            st.error("找不到收件人或 Tracking 列，请确认 Excel 格式")
            st.stop()

        if len(records) == 0:
            st.error("没有找到收件人数据。请先用 Excel 打开文件保存一次后重新上传。")
            st.stop()

        st.session_state.excel_records = records
        st.session_state.recipient_col = recipient_col
        st.session_state.tracking_col = tracking_col
        st.session_state.main_sheet = main_sheet

        st.success(f"已加载 **{len(records)}** 条订单（Sheet: {main_sheet}）")

        with st.expander("预览 Excel 收件人"):
            st.dataframe(
                pd.DataFrame(records)[["excel_row", "收件人", "现有 Tracking"]],
                use_container_width=True, hide_index=True,
            )

        # ── Step 3 ──────────────────────────────────────────────────────────
        st.markdown("---")
        st.markdown("### ③ 匹配结果")

        names_list = [r["收件人"] for r in records]
        match_results = []
        for label in successful_labels:
            if not label["收件人"]:
                continue
            best_idx, best_score, best_name = find_best_match(label["收件人"], names_list)
            match_results.append({
                "面单收件人": label["收件人"],
                "匹配到 Excel": best_name if best_idx >= 0 else "—",
                "相似度": best_score,
                "Tracking": label["Tracking #"],
                "excel_row": records[best_idx]["excel_row"] if best_idx >= 0 else -1,
                "现有 Tracking": records[best_idx]["现有 Tracking"] if best_idx >= 0 else "",
                "接受": best_score >= 0.7,
            })

        if match_results:
            st.session_state.match_results = match_results

            c1, c2, c3 = st.columns(3)
            c1.metric("总识别", len(match_results))
            c2.metric("匹配成功 ≥70%", sum(1 for m in match_results if m["相似度"] >= 0.7))
            c3.metric("匹配失败 <70%", sum(1 for m in match_results if m["相似度"] < 0.7))

            st.markdown("**勾选「接受」确认要回填的条目：**")
            df_match = pd.DataFrame(match_results)
            edited = st.data_editor(
                df_match[["接受", "面单收件人", "匹配到 Excel", "相似度", "Tracking", "现有 Tracking"]],
                use_container_width=True, hide_index=True,
                disabled=["面单收件人", "匹配到 Excel", "相似度", "Tracking", "现有 Tracking"],
                column_config={
                    "接受": st.column_config.CheckboxColumn("✓ 接受", default=False),
                    "相似度": st.column_config.ProgressColumn("相似度", min_value=0, max_value=1, format="%.0f%%"),
                    "Tracking": st.column_config.TextColumn(width="large"),
                },
            )

            accepted_mask = edited["接受"].tolist()
            for i, acc in enumerate(accepted_mask):
                match_results[i]["接受"] = acc

            # ── Step 4 ──────────────────────────────────────────────────────
            st.markdown("---")
            st.markdown("### ④ 下载更新后的 Excel")

            n_accepted = sum(1 for m in match_results if m["接受"])
            if n_accepted == 0:
                st.info("请在上方勾选要回填的条目")
            else:
                st.markdown(f"将回填 **{n_accepted}** 条 Tracking Number")
                if st.button(f"⬇️ 生成更新后的 Excel（{n_accepted} 条）", type="primary", use_container_width=True):
                    wb_write = openpyxl.load_workbook(io.BytesIO(st.session_state.excel_bytes))
                    ws_write = wb_write[st.session_state.main_sheet]

                    filled = 0
                    for m in match_results:
                        if m["接受"] and m["excel_row"] > 0 and m["Tracking"]:
                            ws_write.cell(row=m["excel_row"], column=st.session_state.tracking_col).value = m["Tracking"]
                            filled += 1

                    output = io.BytesIO()
                    wb_write.save(output)
                    output.seek(0)

                    st.download_button(
                        label=f"📥 下载 ParcelOutbound_Updated.xlsx（{filled} 条已填）",
                        data=output.getvalue(),
                        file_name="ParcelOutbound_Updated.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )
                    st.success(f"✅ 共填入 {filled} 条 Tracking Number")
        else:
            st.warning("没有可匹配的识别结果")

# ── Sidebar ──────────────────────────────────────────────────────────────────

with st.sidebar:
    st.markdown("### 📖 使用说明")
    st.markdown("""
    1. 上传面单图片（支持批量）
    2. 点击「开始识别」
    3. 上传 ParcelOutbound Excel
    4. 确认匹配结果
    5. 下载更新后的 Excel
    """)
    st.divider()
    st.markdown("### ℹ️ 支持的面单类型")
    st.markdown("""
    - ✅ UPS（1Z 开头）
    - ✅ FedEx（12-15 位数字）
    - ✅ USPS（20-22 位数字）
    """)
    st.divider()
    st.caption("✅ 完全免费，无需 API Key")
    st.caption("Built for 跨境电商出海业务")
