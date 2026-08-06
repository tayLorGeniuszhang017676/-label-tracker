"""
📦 面单识别 → Tracking 回填工具
Upload shipping labels (PDF/PNG/JPG) → OCR extracts tracking # & recipient → updates Excel
Free: no API key needed
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
    from pdf2image import convert_from_bytes
    PDF_SUPPORT = True
except ImportError:
    PDF_SUPPORT = False

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
for key in ["excel_bytes"]:
    if key not in st.session_state:
        st.session_state[key] = None

# ── OCR + Extraction ─────────────────────────────────────────────────────────

def preprocess_image(img: Image.Image) -> Image.Image:
    img = img.convert("L")
    img = ImageEnhance.Contrast(img).enhance(2.0)
    img = img.filter(ImageFilter.SHARPEN)
    w, h = img.size
    if w < 1000:
        img = img.resize((w * 2, h * 2), Image.LANCZOS)
    return img


def extract_tracking(text: str) -> str:
    """Extract tracking number from OCR text."""
    # UPS: look for TRACKING label
    match = re.search(r'TRACKING\s*#?\s*:?\s*(1Z[\sA-Z0-9]+)', text, re.IGNORECASE)
    if match:
        clean = re.sub(r'\s+', '', match.group(1))
        if clean.startswith('1Z') and len(clean) >= 18:
            return clean[:18]
        return clean

    # UPS: find 1Z pattern anywhere
    match = re.search(r'(1Z\s*[A-Z0-9]{3}\s*[A-Z0-9]{3}\s*[A-Z0-9]{2}\s*[A-Z0-9]{4}\s*[A-Z0-9]{4})', text, re.IGNORECASE)
    if match:
        clean = re.sub(r'\s+', '', match.group(1))
        if len(clean) >= 18:
            return clean[:18]
        return clean

    # FedEx: 12 or 15 digits near "tracking" or standalone
    match = re.search(r'(?:TRACKING|TRACK)\s*#?\s*:?\s*(\d{12,15})', text, re.IGNORECASE)
    if match:
        return match.group(1)

    # USPS: 20-22 digits
    match = re.search(r'\b(\d{20,22})\b', text)
    if match:
        return match.group(1)

    return ""


def extract_recipient(text: str) -> str:
    """Extract recipient name from OCR text."""
    # Pattern 1: SHIP TO: followed by name
    match = re.search(r'SHIP\s*TO\s*:?\s*\n\s*(.+?)(?:\n|$)', text, re.IGNORECASE)
    if match:
        name = re.sub(r'[^\w\s\'-]', '', match.group(1).strip()).strip()
        if len(name) >= 2 and not name.isdigit() and not re.match(r'^\d', name):
            return name

    # Pattern 2: DELIVER TO / TO:
    match = re.search(r'(?:DELIVER|RECIPIENT)\s*(?:TO)?\s*:?\s*\n\s*(.+?)(?:\n|$)', text, re.IGNORECASE)
    if match:
        name = re.sub(r'[^\w\s\'-]', '', match.group(1).strip()).strip()
        if len(name) >= 2 and not name.isdigit() and not re.match(r'^\d', name):
            return name

    # Pattern 3: look for name block after SHIP TO
    block = re.search(
        r'SHIP\s*TO\s*:?(.*?)(?:\d{1,5}\s+\w)',
        text, re.IGNORECASE | re.DOTALL
    )
    if block:
        lines = [l.strip() for l in block.group(1).strip().split('\n') if l.strip()]
        for line in lines:
            clean = re.sub(r'[^\w\s\'-]', '', line).strip()
            if clean and not clean.isdigit() and len(clean) >= 2 and not re.match(r'^\d', clean):
                return clean

    return ""


def pdf_to_images(pdf_bytes: bytes) -> list:
    """Convert PDF pages to PIL Images."""
    if not PDF_SUPPORT:
        return []
    try:
        images = convert_from_bytes(pdf_bytes, dpi=300)
        return images
    except Exception as e:
        st.error(f"PDF 转换失败: {e}")
        return []


def process_single_image(img: Image.Image) -> dict:
    """Process one image through OCR."""
    try:
        processed = preprocess_image(img)
        text = pytesseract.image_to_string(processed, config='--psm 6')
        tracking = extract_tracking(text)
        recipient = extract_recipient(text)
        return {
            "recipient_name": recipient or None,
            "tracking_number": tracking or None,
            "ocr_text": text,
            "error": None if (tracking or recipient) else "无法提取信息，可能图片不够清晰",
        }
    except Exception as e:
        return {"recipient_name": None, "tracking_number": None, "ocr_text": "", "error": str(e)}


def process_file(file_bytes: bytes, filename: str) -> list:
    """
    Process an uploaded file. Returns a list of results
    (one per page for PDFs, one for images).
    """
    ext = Path(filename).suffix.lower()
    results = []

    if ext == ".pdf":
        images = pdf_to_images(file_bytes)
        if not images:
            return [{"filename": filename, "recipient_name": None,
                     "tracking_number": None, "ocr_text": "",
                     "error": "PDF 转图片失败"}]
        for i, img in enumerate(images):
            extracted = process_single_image(img)
            page_label = f"{filename} (第{i+1}页)" if len(images) > 1 else filename
            # Only include pages that have tracking or recipient info
            if extracted.get("tracking_number") or extracted.get("recipient_name"):
                results.append({"filename": page_label, **extracted})
            elif i == 0 and len(images) == 1:
                # Single page PDF that failed
                results.append({"filename": page_label, **extracted})
        # If multi-page PDF had no results, report it
        if not results:
            results.append({"filename": filename, "recipient_name": None,
                           "tracking_number": None, "ocr_text": "",
                           "error": f"PDF 共 {len(images)} 页，未识别到面单信息"})
    else:
        img = Image.open(io.BytesIO(file_bytes))
        extracted = process_single_image(img)
        results.append({"filename": filename, **extracted})

    return results


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
    wb = openpyxl.load_workbook(io.BytesIO(excel_bytes), data_only=True)
    main_sheet = wb.sheetnames[0]
    ws = wb[main_sheet]

    recipient_col = tracking_col = None
    for col in range(1, ws.max_column + 1):
        val = ws.cell(row=1, column=col).value
        if val:
            s = str(val)
            if "收件人" in s or "RecipientName" in s:
                recipient_col = col
            if "Tracking" in s or "跟踪号" in s:
                tracking_col = col

    if not recipient_col or not tracking_col:
        return None, None, None, None

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

    # If VLOOKUP cached values missing, try source sheet
    if all(r["收件人"] in ("", "None") for r in records):
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
                    ws_m = wb[main_sheet]
                    pcol = None
                    for col in range(1, ws_m.max_column + 1):
                        val = ws_m.cell(row=1, column=col).value
                        if val and ("Platform Number" in str(val) or "平台单号" in str(val)):
                            pcol = col
                            break
                    if pcol and o2n:
                        for rec in records:
                            on = ws_m.cell(row=rec["excel_row"], column=pcol).value
                            if on and str(on).strip() in o2n:
                                rec["收件人"] = o2n[str(on).strip()]
                    break

    records = [r for r in records if r["收件人"] and r["收件人"] not in ("", "None")]
    return records, recipient_col, tracking_col, main_sheet


# ── UI ───────────────────────────────────────────────────────────────────────

st.markdown("# 📦 面单识别 → Tracking 回填")
st.markdown("上传面单（PDF 或图片），自动识别 Tracking Number，按收件人匹配回填到出库单")
st.caption("✅ 完全免费，无需 API Key | 支持 PDF + PNG/JPG 批量上传")

# ── Step 1 ───────────────────────────────────────────────────────────────────

st.markdown("---")
st.markdown("### ① 上传面单")

uploaded_files = st.file_uploader(
    "选择面单文件（支持 PDF / PNG / JPG，可多选）",
    type=["pdf", "png", "jpg", "jpeg", "webp", "bmp"],
    accept_multiple_files=True,
    key="label_uploader",
)

if uploaded_files:
    # Show file list
    file_info = []
    for f in uploaded_files:
        ext = Path(f.name).suffix.lower()
        icon = "📄" if ext == ".pdf" else "🖼️"
        file_info.append(f"{icon} {f.name} ({f.size / 1024:.0f} KB)")
    st.markdown("**已选文件：** " + " · ".join(file_info))

    if st.button("🔍 开始识别", type="primary", use_container_width=True):
        all_results = []
        progress = st.progress(0, text="正在识别面单...")

        for i, uploaded_file in enumerate(uploaded_files):
            progress.progress(
                (i + 1) / len(uploaded_files),
                text=f"正在处理 {uploaded_file.name} ({i+1}/{len(uploaded_files)})",
            )
            file_results = process_file(uploaded_file.getvalue(), uploaded_file.name)
            for r in file_results:
                all_results.append({
                    "文件名": r["filename"],
                    "收件人": r.get("recipient_name") or "",
                    "Tracking #": r.get("tracking_number") or "",
                    "状态": f"❌ {r['error']}" if r.get("error") else "✅ 成功",
                    "_ocr_text": r.get("ocr_text", ""),
                })

        progress.empty()
        st.session_state.extracted_labels = all_results
        n_ok = sum(1 for r in all_results if r["状态"].startswith("✅"))
        if n_ok > 0:
            st.success(f"识别完成！成功 {n_ok}/{len(all_results)} 条")
        else:
            st.error("识别失败，请检查文件清晰度。可展开下方调试信息查看 OCR 原文。")

# Show results
if st.session_state.extracted_labels:
    st.markdown("**识别结果：**")
    df_labels = pd.DataFrame(st.session_state.extracted_labels)
    display_cols = [c for c in df_labels.columns if not c.startswith("_")]
    st.dataframe(df_labels[display_cols], use_container_width=True, hide_index=True)

    with st.expander("🔍 调试：查看 OCR 原始文本"):
        for r in st.session_state.extracted_labels:
            st.markdown(f"**{r['文件名']}:**")
            st.code(r.get("_ocr_text", ""), language=None)

# ── Step 2 ───────────────────────────────────────────────────────────────────

successful = [l for l in st.session_state.extracted_labels if l["状态"].startswith("✅")]
if successful:
    st.markdown("---")
    st.markdown("### ② 上传 ParcelOutbound Excel")

    uploaded_excel = st.file_uploader("选择 Excel 文件", type=["xlsx", "xls"], key="excel_uploader")

    if uploaded_excel:
        excel_bytes = uploaded_excel.getvalue()
        st.session_state.excel_bytes = excel_bytes
        records, rcol, tcol, sheet = read_excel_recipients(excel_bytes)

        if records is None:
            st.error("找不到收件人或 Tracking 列")
            st.stop()
        if len(records) == 0:
            st.error("没有找到收件人数据。请先用 Excel 打开文件保存一次后重新上传。")
            st.stop()

        st.session_state.excel_records = records
        st.session_state.recipient_col = rcol
        st.session_state.tracking_col = tcol
        st.session_state.main_sheet = sheet
        st.success(f"已加载 **{len(records)}** 条订单（Sheet: {sheet}）")

        with st.expander("预览 Excel 收件人"):
            st.dataframe(pd.DataFrame(records)[["excel_row", "收件人", "现有 Tracking"]],
                        use_container_width=True, hide_index=True)

        # ── Step 3 ──────────────────────────────────────────────────────────
        st.markdown("---")
        st.markdown("### ③ 匹配结果")

        names_list = [r["收件人"] for r in records]
        match_results = []
        for label in successful:
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
            c3.metric("低匹配 <70%", sum(1 for m in match_results if m["相似度"] < 0.7))

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
            for i, acc in enumerate(edited["接受"].tolist()):
                match_results[i]["接受"] = acc

            # ── Step 4 ──────────────────────────────────────────────────────
            st.markdown("---")
            st.markdown("### ④ 下载更新后的 Excel")
            n_acc = sum(1 for m in match_results if m["接受"])
            if n_acc == 0:
                st.info("请在上方勾选要回填的条目")
            else:
                st.markdown(f"将回填 **{n_acc}** 条 Tracking Number")
                if st.button(f"⬇️ 生成更新后的 Excel（{n_acc} 条）", type="primary", use_container_width=True):
                    wb_w = openpyxl.load_workbook(io.BytesIO(st.session_state.excel_bytes))
                    ws_w = wb_w[st.session_state.main_sheet]
                    filled = 0
                    for m in match_results:
                        if m["接受"] and m["excel_row"] > 0 and m["Tracking"]:
                            ws_w.cell(row=m["excel_row"], column=st.session_state.tracking_col).value = m["Tracking"]
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
    1. 上传面单文件（PDF 或图片）
    2. 点击「开始识别」
    3. 上传 ParcelOutbound Excel
    4. 确认匹配结果
    5. 下载更新后的 Excel
    """)
    st.divider()
    st.markdown("### ℹ️ 支持格式")
    st.markdown("""
    **面单：** PDF / PNG / JPG
    **快递：** UPS (1Z) / FedEx / USPS
    **Excel：** .xlsx
    """)
    st.divider()
    st.caption("✅ 完全免费，无需 API Key")
