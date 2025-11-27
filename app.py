# app.py
"""
Streamlit app: Collate Failure Report + Failure Summary into core buckets,
auto-loads mapping.xlsx from the repo. The downloaded Excel will have
'Reason' columns grouped & collapsed so the + expand control is visible in Excel.

This version is defensive: it tries st.data_editor, falls back to
st.experimental_data_editor, and if neither exists provides a CSV-based
mapping edit fallback so the app won't crash on older Streamlit runtimes.
"""
import streamlit as st
import pandas as pd
import re
from collections import defaultdict
from io import BytesIO
from datetime import datetime
import os
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows

st.set_page_config(page_title="Myntra QC Collator", layout="wide")

CORE_ORDER = [
    "Compliance", "Content", "Formatting", "Image Formatting",
    "Image", "Size", "Image Sequence"
]

def split_bullets(text):
    if pd.isna(text) or text is None:
        return []
    s = str(text)
    parts = re.split(r'\n|•|◦|·|-{2,}|—|–', s)
    return [p.strip(" \t:.-") for p in parts if p.strip()]

def parse_reason_and_msg(bullet):
    if not bullet:
        return ("", "")
    parts = re.split(r'\s*[:\-—–]\s*', bullet, maxsplit=1)
    if len(parts) == 2:
        title = parts[0].strip()
        msg = parts[1].strip()
    else:
        title = parts[0].strip()
        msg = ""
    return (title, msg)

def normalize_title(t):
    return re.sub(r'\s+', ' ', str(t).strip().lower())

def excel_bytes_from_wb(wb):
    bio = BytesIO()
    wb.save(bio)
    return bio.getvalue()

def build_grouped_workbook(visible_df, full_df, reason_columns, mapping_used_df):
    wb = Workbook()
    ws = wb.active
    ws.title = "cleaned_output"
    # Write header
    all_cols = list(visible_df.columns) + reason_columns
    ws.append(all_cols)
    # Write rows
    for _, row in full_df.iterrows():
        row_vals = [row.get(c, "") for c in visible_df.columns] + [row.get(c, "") for c in reason_columns]
        ws.append(row_vals)
    # Set outline/hide reason columns (they will be collapsed in Excel)
    start_idx = len(visible_df.columns) + 1
    end_idx = start_idx + len(reason_columns) - 1
    for col_idx in range(start_idx, end_idx + 1):
        col_letter = get_column_letter(col_idx)
        ws.column_dimensions[col_letter].outlineLevel = 1
        ws.column_dimensions[col_letter].hidden = True
    # mapping_used sheet
    if mapping_used_df is None:
        mapping_used_df = pd.DataFrame(columns=["Header","CoreBucket"])
    ws_map = wb.create_sheet("mapping_used")
    for r in dataframe_to_rows(mapping_used_df, index=False, header=True):
        ws_map.append(r)
    # dropped columns sheet
    ws_dropped = wb.create_sheet("dropped_reason_columns")
    for c in reason_columns:
        ws_dropped.append([c])
    return wb

# --- UI ---
st.title("Myntra QC — Collate & Clean (robust editor + grouped export)")

st.markdown("""
Upload your Output Excel (XLSX). This app will:
- Collate Failure Summary + Failure Report into core buckets,
- Remove columns containing 'Reason' in their headers from the visible sheet,
- Export an Excel with Reason columns grouped & collapsed (so Excel shows + to expand),
- Load bundled `mapping.xlsx` (editable). App is compatible with multiple Streamlit versions.
""")

uploaded_xlsx = st.file_uploader("Upload Output Excel (XLSX)", type=["xlsx"])
col1, col2, col3 = st.columns(3)
with col1:
    style_col = st.text_input("styleid column name", value="styleid")
with col2:
    summary_col = st.text_input("Failure Summary column name", value="Failure Summary")
with col3:
    report_col = st.text_input("Failure Report column name", value="Failure Report")

# Load mapping.xlsx bundled with repo
mapping_path_local = os.path.join(os.path.dirname(__file__), "mapping.xlsx")
if not os.path.exists(mapping_path_local):
    st.error("mapping.xlsx not found in the app folder. Please ensure mapping.xlsx exists.")
    st.stop()

map_df = pd.read_excel(mapping_path_local).fillna("")
# Normalize mapping columns
cols = [c.strip() for c in map_df.columns]
if "Header" in cols and "CoreBucket" in cols:
    map_df = map_df[["Header", "CoreBucket"]]
else:
    # fallback: rename first two columns
    if len(map_df.columns) >= 2:
        map_df.columns = ["Header", "CoreBucket"]
    else:
        # create empty scaffold
        map_df = pd.DataFrame(columns=["Header", "CoreBucket"])

st.subheader("Bundled mapping (edit if needed)")

# Robust editor selection:
editor_df = None
editor_used_method = None
if hasattr(st, "data_editor"):
    try:
        editor_df = st.data_editor(map_df, num_rows="dynamic")
        editor_used_method = "data_editor"
    except Exception:
        editor_df = None

if editor_df is None and hasattr(st, "experimental_data_editor"):
    try:
        editor_df = st.experimental_data_editor(map_df, num_rows="dynamic")
        editor_used_method = "experimental_data_editor"
    except Exception:
        editor_df = None

if editor_df is None:
    st.warning("Interactive mapping editor not available in this Streamlit runtime. Use the CSV fallback below to edit mapping.")
    st.markdown("Download the current mapping, edit it locally as CSV, then upload below to use the edited mapping.")
    csv_bytes = map_df.to_csv(index=False).encode("utf-8")
    st.download_button("Download current mapping (CSV)", data=csv_bytes, file_name="mapping_current.csv", mime="text/csv")
    uploaded_map_csv = st.file_uploader("Upload edited mapping CSV (optional)", type=["csv"])
    if uploaded_map_csv:
        try:
            editor_df = pd.read_csv(uploaded_map_csv).fillna("")
            # ensure at least two columns exist
            if len(editor_df.columns) >= 2:
                editor_df = editor_df.iloc[:, :2]
                editor_df.columns = ["Header", "CoreBucket"]
            else:
                st.error("Uploaded CSV must have at least two columns (Header, CoreBucket). Using bundled mapping.")
                editor_df = map_df.copy()
        except Exception as e:
            st.error("Failed to read uploaded CSV. Using bundled mapping.")
            editor_df = map_df.copy()
    else:
        editor_df = map_df.copy()
    editor_used_method = "csv_fallback"

st.caption(f"Mapping editor method used: {editor_used_method}")

if uploaded_xlsx is None:
    st.info("Upload the Output Excel to collate and clean. The app will use the bundled mapping by default.")
    st.stop()

# Read input workbook
try:
    input_book = pd.read_excel(uploaded_xlsx, sheet_name=None, dtype=str)
    sheet_name = list(input_book.keys())[0]
    df = input_book[sheet_name].fillna("")
except Exception as e:
    st.error("Failed to read uploaded Excel: " + str(e))
    st.stop()

# Build mapping dict from edited mapping
mapping = {}
for _, r in editor_df.iterrows():
    hdr = str(r.get("Header","")).strip()
    bucket = str(r.get("CoreBucket","")).strip() or "Unmapped"
    if hdr:
        mapping[hdr] = bucket

# find reason columns
all_cols = list(df.columns)
reason_cols = [c for c in all_cols if re.search(r'\breason\b', str(c), flags=re.IGNORECASE)]

# Ensure core columns exist
for core in CORE_ORDER:
    if core not in df.columns:
        df[core] = ""

unmapped_reasons = set()

# Process each row: collation logic (same conservative matching)
for idx, row in df.iterrows():
    summary_text = row.get(summary_col, "") if summary_col in df.columns else ""
    report_text = row.get(report_col, "") if report_col in df.columns else ""

    summary_bullets = split_bullets(summary_text)
    report_bullets = split_bullets(report_text)

    summary_parsed = [parse_reason_and_msg(b) for b in summary_bullets]
    report_parsed = [parse_reason_and_msg(b) for b in report_bullets]

    report_map = {}
    for title, msg in report_parsed:
        key = normalize_title(title) if title else normalize_title(msg[:30])
        report_map.setdefault(key, []).append(msg or "")

    core_bucket_msgs = defaultdict(list)
    used_report_indices = set()

    for i, (title, s_msg) in enumerate(summary_parsed):
        key = normalize_title(title)
        matched = False
        if key and key in report_map:
            for rep_msg in report_map[key]:
                full_msg = rep_msg if rep_msg else s_msg
                bucket = mapping.get(title, None) or mapping.get(title.strip(), None)
                if not bucket:
                    unmapped_reasons.add(title)
                    bucket = "Unmapped"
                core_bucket_msgs[bucket].append(f"{title}: {full_msg}".strip())
                matched = True
        if not matched:
            found = False
            for j, (r_title, r_msg) in enumerate(report_parsed):
                if j in used_report_indices:
                    continue
                hay = (r_title or "") + " " + (r_msg or "")
                if title and title.lower() in hay.lower():
                    bucket = mapping.get(title, None) or "Unmapped"
                    if bucket == "Unmapped":
                        unmapped_reasons.add(title)
                    core_bucket_msgs[bucket].append(f"{title}: {r_msg or s_msg}".strip())
                    used_report_indices.add(j)
                    found = True
                    break
            if not found:
                bucket = mapping.get(title, None) or "Unmapped"
                if bucket == "Unmapped":
                    unmapped_reasons.add(title)
                core_bucket_msgs[bucket].append(f"{title}: {s_msg}".strip())

    for j, (r_title, r_msg) in enumerate(report_parsed):
        if j in used_report_indices:
            continue
        title = r_title or (r_msg[:30] if r_msg else "Unknown")
        bucket = mapping.get(title, None) or mapping.get(title.strip(), None) or "Unmapped"
        if bucket == "Unmapped":
            unmapped_reasons.add(title)
        core_bucket_msgs[bucket].append(f"{title}: {r_msg}".strip())

    for core in CORE_ORDER:
        if core in core_bucket_msgs and core_bucket_msgs[core]:
            df.at[idx, core] = "\n".join(core_bucket_msgs[core])

# visible_df drops reason columns
visible_df = df.drop(columns=reason_cols, errors='ignore')

st.success("Collation complete. Preview below.")
st.dataframe(visible_df.head(10), use_container_width=True)

# Build grouped workbook and offer download
wb = build_grouped_workbook(visible_df, df, reason_cols, editor_df)
out_bytes = excel_bytes_from_wb(wb)

st.download_button(
    "Download cleaned & grouped Excel (Reason cols collapsed)",
    data=out_bytes,
    file_name=f"myntra_collated_grouped_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

st.info("Tip: When you open the downloaded Excel in MS Excel, you'll see the + control to expand the grouped 'Reason' columns.")
