"""
📦 面单文件名 → Tracking 回填工具
PDF 文件名 = 订单号_Tracking → 直接解析文件名（无需OCR）→ 按订单号匹配回填 Excel
"""

import streamlit as st
import pandas as pd
import re
import io
from pathlib import Path

try:
    import openpyxl
except ImportError:
    openpyxl = None

# ── Page config ──────────────────────────────────────────────────────────────

st.set_page_config(page_title="面单识别 → Tracking 回填", page_icon="📦", layout="wide")

st.markdown("""
<style>
    div[data-testid="stMetric"] {
        background: #f8fafc; border: 1px solid #e2e8f0;
        border-radius: 8px; padding: 12px 16px;
    }
</style>
""", unsafe_allow_html=True)

# ── Session state ────────────────────────────────────────────────────────────

for key in ["extracted_labels", "match_results"]:
    if key not in st.session_state:
        st.session_state[key] = []
if "excel_bytes" not in st.session_state:
    st.session_state.excel_bytes = None

# ── 文件名解析（替代 OCR）─────────────────────────────────────────────────────

def parse_filename(filename: str):
    """从文件名解析 订单号 + Tracking。
    文件名格式：订单号_Tracking.pdf
    e.g. '114-8302232-3163464_1Z2W4A130342999947.pdf'
         → 订单号 '114-8302232-3163464', Tracking '1Z2W4A130342999947'
    只按第一个下划线切分，Tracking 内部如再有下划线会原样保留。
    """
    stem = Path(filename).stem                      # 去掉 .pdf / .png 等后缀
    stem = re.sub(r'\s*\(\d+\)\s*$', '', stem)      # 去掉重复文件的 " (1)" 等尾巴
    stem = stem.strip()

    if "_" not in stem:
        return stem, ""                             # 没有下划线 → 只有订单号，无 Tracking

    order, tracking = stem.split("_", 1)
    # 如需让 Tracking 保留下划线本身（即 "_XXXX"），把上面这行换成：
    # order, tracking = stem.split("_", 1); tracking = "_" + tracking
    return order.strip(), tracking.strip()


def process_file(filename: str) -> dict:
    """处理一个上传文件：只解析文件名，不读取文件内容。"""
    order, tracking = parse_filename(filename)
    return {
        "filename": filename,
        "order_number": order,
        "tracking": tracking,
        "error": None if tracking else "文件名中没有下划线，无法提取 Tracking",
    }


# ── Excel ────────────────────────────────────────────────────────────────────

def read_excel(excel_bytes: bytes):
    """Read Excel and return records with order numbers."""
    wb = openpyxl.load_workbook(io.BytesIO(excel_bytes), data_only=True)
    main_sheet = wb.sheetnames[0]
    ws = wb[main_sheet]

    # Find columns
    cols = {}
    for col in range(1, ws.max_column + 1):
        val = ws.cell(row=1, column=col).value
        if not val:
            continue
        s = str(val)
        if "Platform Number" in s or "平台单号" in s:
            cols["order"] = col
        if "Tracking" in s or "跟踪号" in s:
            cols["tracking"] = col
        if "收件人" in s or "RecipientName" in s:
            cols["recipient"] = col

    if "order" not in cols or "tracking" not in cols:
        return None, None

    # Read recipient names (handle VLOOKUP)
    # First try cached values
    records = []
    for row in range(2, ws.max_row + 1):
        order_val = ws.cell(row=row, column=cols["order"]).value
        tracking_val = ws.cell(row=row, column=cols["tracking"]).value
        recipient_val = ws.cell(row=row, column=cols.get("recipient", 1)).value if "recipient" in cols else ""

        if order_val:
            records.append({
                "excel_row": row,
                "订单号": str(order_val).strip(),
                "收件人": str(recipient_val).strip() if recipient_val else "",
                "现有 Tracking": str(tracking_val).strip() if tracking_val else "",
            })

    # If recipient is VLOOKUP with no cached value, try source sheet
    if "recipient" in cols and all(r["收件人"] in ("", "None") for r in records):
        for sn in wb.sheetnames:
            if sn != main_sheet:
                ws2 = wb[sn]
                src_rcpt = src_order = None
                for col in range(1, min(30, ws2.max_column + 1)):
                    val = ws2.cell(row=1, column=col).value
                    if val:
                        if "收件人" in str(val): src_rcpt = col
                        if "订单号" in str(val): src_order = col
                if src_rcpt and src_order:
                    o2n = {}
                    for row in range(2, ws2.max_row + 1):
                        o = ws2.cell(row=row, column=src_order).value
                        n = ws2.cell(row=row, column=src_rcpt).value
                        if o and n:
                            o2n[str(o).strip()] = str(n).strip()
                    for rec in records:
                        if rec["订单号"] in o2n:
                            rec["收件人"] = o2n[rec["订单号"]]
                    break

    return records, {
        "tracking_col": cols["tracking"],
        "main_sheet": main_sheet,
    }


# ── UI ───────────────────────────────────────────────────────────────────────

st.markdown("# 📦 面单识别 → Tracking 回填")
st.markdown("上传面单文件（**文件名 = 订单号_Tracking**），直接从文件名提取 Tracking，按订单号精准匹配回填到 Excel")
st.caption("✅ 完全免费 | 无需 OCR，秒出结果 | 按订单号精准匹配")

# ── Step 1 ───────────────────────────────────────────────────────────────────

st.markdown("---")
st.markdown("### ① 上传面单")
st.caption("⚠️ 文件名格式：`订单号_Tracking`，如 `114-8302232-3163464_1Z2W4A130342999947.pdf`（第一个下划线前 = 订单号，下划线后 = Tracking）")

uploaded_files = st.file_uploader(
    "选择面单文件（PDF / PNG / JPG，可多选）",
    type=["pdf", "png", "jpg", "jpeg", "webp", "bmp"],
    accept_multiple_files=True,
    key="label_uploader",
)

if uploaded_files:
    # Preview file list with extracted order numbers + tracking
    preview_data = []
    for f in uploaded_files:
        order, tracking = parse_filename(f.name)
        preview_data.append({
            "文件": f.name,
            "提取的订单号": order,
            "提取的 Tracking": tracking if tracking else "⚠️ 无下划线",
        })
    st.dataframe(pd.DataFrame(preview_data), use_container_width=True, hide_index=True)

    if st.button("🔍 开始识别 Tracking", type="primary", use_container_width=True):
        all_results = []
        for f in uploaded_files:
            r = process_file(f.name)
            all_results.append({
                "文件名": r["filename"],
                "订单号": r["order_number"],
                "Tracking #": r["tracking"],
                "状态": f"❌ {r['error']}" if r["error"] else "✅ 成功",
            })

        st.session_state.extracted_labels = all_results
        n_ok = sum(1 for r in all_results if r["状态"].startswith("✅"))
        if n_ok > 0:
            st.success(f"识别完成！成功提取 {n_ok}/{len(all_results)} 条 Tracking")
        else:
            st.error("未提取到 Tracking，请检查文件名中是否包含下划线")

# Show results
if st.session_state.extracted_labels:
    st.markdown("**识别结果：**")
    df = pd.DataFrame(st.session_state.extracted_labels)
    st.dataframe(df, use_container_width=True, hide_index=True)

# ── Step 2 ───────────────────────────────────────────────────────────────────

successful = [l for l in st.session_state.extracted_labels if l["状态"].startswith("✅")]
if successful:
    st.markdown("---")
    st.markdown("### ② 上传 ParcelOutbound Excel")

    uploaded_excel = st.file_uploader("选择 Excel 文件", type=["xlsx", "xls"], key="excel_uploader")

    if uploaded_excel:
        excel_bytes = uploaded_excel.getvalue()
        st.session_state.excel_bytes = excel_bytes
        records, meta = read_excel(excel_bytes)

        if records is None:
            st.error("找不到「平台单号」或「Tracking」列，请确认 Excel 格式")
            st.stop()
        if len(records) == 0:
            st.error("Excel 中没有订单数据")
            st.stop()

        st.session_state.excel_records = records
        st.session_state.meta = meta
        st.success(f"已加载 **{len(records)}** 条订单")

        # Build order number lookup
        order_lookup = {}  # order_number -> index in records
        for idx, rec in enumerate(records):
            order_lookup[rec["订单号"]] = idx

        # ── Step 3: Match ────────────────────────────────────────────────
        st.markdown("---")
        st.markdown("### ③ 匹配结果（按订单号精准匹配）")

        match_results = []
        for label in successful:
            order = label["订单号"]
            if order in order_lookup:
                idx = order_lookup[order]
                rec = records[idx]
                match_results.append({
                    "订单号": order,
                    "Tracking": label["Tracking #"],
                    "Excel 收件人": rec["收件人"],
                    "现有 Tracking": rec["现有 Tracking"],
                    "excel_row": rec["excel_row"],
                    "匹配": "✅ 精准匹配",
                    "接受": True,
                })
            else:
                match_results.append({
                    "订单号": order,
                    "Tracking": label["Tracking #"],
                    "Excel 收件人": "",
                    "现有 Tracking": "",
                    "excel_row": -1,
                    "匹配": "❌ Excel 中无此订单号",
                    "接受": False,
                })

        st.session_state.match_results = match_results

        n_matched = sum(1 for m in match_results if m["excel_row"] > 0)
        n_miss = len(match_results) - n_matched

        c1, c2, c3 = st.columns(3)
        c1.metric("总面单", len(match_results))
        c2.metric("匹配成功", n_matched)
        c3.metric("未匹配", n_miss)

        st.markdown("**确认要回填的条目：**")
        df_match = pd.DataFrame(match_results)
        edited = st.data_editor(
            df_match[["接受", "订单号", "匹配", "Tracking", "Excel 收件人", "现有 Tracking"]],
            use_container_width=True, hide_index=True,
            disabled=["订单号", "匹配", "Tracking", "Excel 收件人", "现有 Tracking"],
            column_config={
                "接受": st.column_config.CheckboxColumn("✓ 接受", default=False),
                "Tracking": st.column_config.TextColumn(width="large"),
            },
        )
        for i, acc in enumerate(edited["接受"].tolist()):
            match_results[i]["接受"] = acc

        # ── Step 4: Download ─────────────────────────────────────────────
        st.markdown("---")
        st.markdown("### ④ 下载更新后的 Excel")
        n_acc = sum(1 for m in match_results if m["接受"] and m["excel_row"] > 0)
        if n_acc == 0:
            st.info("没有可回填的条目")
        else:
            st.markdown(f"将回填 **{n_acc}** 条 Tracking Number")
            if st.button(f"⬇️ 生成更新后的 Excel（{n_acc} 条）", type="primary", use_container_width=True):
                wb_w = openpyxl.load_workbook(io.BytesIO(st.session_state.excel_bytes))
                ws_w = wb_w[st.session_state.meta["main_sheet"]]
                filled = 0
                for m in match_results:
                    if m["接受"] and m["excel_row"] > 0 and m["Tracking"]:
                        ws_w.cell(row=m["excel_row"], column=st.session_state.meta["tracking_col"]).value = m["Tracking"]
                        filled += 1
                out = io.BytesIO()
                wb_w.save(out)
                out.seek(0)
                st.download_button(
                    f"📥 下载 ParcelOutbound_Updated.xlsx（{filled} 条已填）",
                    data=out.getvalue(),
                    file_name="ParcelOutbound_Updated.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
                st.success(f"✅ 共填入 {filled} 条 Tracking Number")

# ── Sidebar ──────────────────────────────────────────────────────────────────

with st.sidebar:
    st.markdown("### 📖 使用说明")
    st.markdown("""
    1. 上传面单文件（**文件名 = 订单号_Tracking**）
    2. 点击「开始识别」提取 Tracking（直接取文件名下划线后的部分）
    3. 上传 ParcelOutbound Excel
    4. 按订单号自动精准匹配
    5. 确认后下载更新的 Excel
    """)
    st.divider()
    st.markdown("### ⚠️ 重要")
    st.markdown("""
    文件名必须是 `订单号_Tracking`！
    例如：
    `114-8302232-3163464_1Z2W4A130342999947.pdf`
    第一个下划线之前用于匹配 Excel 中的
    「Platform Number/平台单号」列，
    下划线之后的部分作为 Tracking 回填
    """)
    st.divider()
    st.markdown("### ℹ️ 支持格式")
    st.markdown("**面单：** PDF / PNG / JPG")
    st.markdown("**提取方式：** 文件名解析（无OCR）")
    st.divider()
    st.caption("✅ 完全免费，无需 API Key")


