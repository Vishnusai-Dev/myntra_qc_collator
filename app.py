# app.py
"""
Streamlit app — minimal, robust implementation:
- Adds 7 core columns after styleid
- Concatenates Failure Summary + Failure Report into the correct core column using mapping.xlsx
- Removes all columns containing the word 'Reason' in the header (case-insensitive)
- Reshuffles original columns so mapped columns are grouped by core bucket order
- Keeps original column values unchanged (only adds core columns)
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

st.set_page_config(page_title="Myntra QC — Collate to core buckets", layout="wide")

CORE_ORDER = ["Compliance", "Content", "Formatting", "Image Formatting", "Image", "Size", "Image Sequence"]

def normalize(s):
    if s is None:
        return ""
    s = str(s).strip()
    s = re.sub(r'[\t\r\n]+', ' ', s)
    s = re.sub(r'\s+', ' ', s)
    s = s.strip(" .:-–—")
    return s.lower()

def split_bullets(text):
    if pd.isna(text) or text is None:
        return []
    s = str(text)
    parts = re.split(r'\n|•|◦|·|-{2,}|—|–', s)
    return [p.strip() for p in parts if p.strip()]

def parse_title_msg(bullet):
    if not bullet:
        return ("","")
    parts = re.split(r'\s*[:\-—–]\s*', bullet, maxsplit=1)
    if len(parts) == 2:
        return parts[0].strip(), parts[1].strip()
    return parts[0].strip(), ""

def build_excel_bytes(df, filename="collated_output.xlsx"):
    # simple writer using openpyxl via pandas
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="cleaned_output", index=False)
        writer.save()
    bio.seek(0)
    return bio.getvalue()

st.title("Myntra QC — Collate failures into 7 core buckets")

st.markdown("""
Upload your Output Excel (XLSX).  
This tool will:
- Remove columns whose header contains the word **Reason**,
- Add 7 core columns (Compliance, Content, Formatting, Image Formatting, Image, Size, Image Sequence) after `styleid`,
- Concatenate mapped Failure Summary + Failure Report messages into the correct core column (one block per bucket),
- Keep all original columns (values unchanged) but reshuffle them grouped by bucket order from mapping.xlsx.
""")

uploaded = st.file_uploader("Upload Output Excel (.xlsx)", type=["xlsx"])
style_col = st.text_input("styleid column name", value="styleid")
summary_col = st.text_input("Failure Summary column name", value="Failure Summary")
report_col = st.text_input("Failure Report column name", value="Failure Report")

# Load mapping.xlsx from repo (bundled)
mapping_path = os.path.join(os.path.dirname(__file__), "mapping.xlsx")
if not os.path.exists(mapping_path):
    st.error("mapping.xlsx not found next to app.py. Please add mapping.xlsx (Header, CoreBucket).")
    st.stop()

map_df = pd.read_excel(mapping_path).fillna("")
# Use first two columns as Header and CoreBucket
if len(map_df.columns) >= 2:
    map_df = map_df.iloc[:, :2]
    map_df.columns = ["Header", "CoreBucket"]
else:
    map_df = pd.DataFrame(columns=["Header","CoreBucket"])

# Build normalized mapping dict: normalized header -> core bucket
mapping_norm_to_bucket = {}
mapping_header_order = []  # preserve mapping order if needed
for _, r in map_df.iterrows():
    hdr = str(r["Header"]).strip()
    bucket = str(r["CoreBucket"]).strip()
    if hdr:
        mapping_norm_to_bucket[normalize(hdr)] = bucket
        mapping_header_order.append(hdr)

# upload check
if uploaded is None:
    st.info("Upload the Output Excel to process.")
    st.stop()

# read input (first sheet)
try:
    xls = pd.read_excel(uploaded, sheet_name=None, dtype=str)
    sheet_name = list(xls.keys())[0]
    df_in = xls[sheet_name].fillna("")
except Exception as e:
    st.error("Failed to read uploaded Excel: " + str(e))
    st.stop()

# 1) Remove all "Reason" columns (case-insensitive)
all_cols = list(df_in.columns)
reason_cols = [c for c in all_cols if re.search(r'\breason\b', str(c), flags=re.IGNORECASE)]
df_no_reasons = df_in.drop(columns=reason_cols, errors='ignore')

# 2) Prepare output dataframe: keep original columns (without reason cols) but we will insert core cols after styleid
orig_cols_no_reason = list(df_no_reasons.columns)

# If style_col not present, warn and try to proceed with first column as styleid
if style_col not in orig_cols_no_reason:
    st.warning(f"styleid column '{style_col}' not found in uploaded sheet. Using first column as styleid.")
    style_col = orig_cols_no_reason[0]

# Build list of original non-style columns (preserve original order)
orig_after_style = [c for c in orig_cols_no_reason if c != style_col]

# 3) Reshuffle original columns based on mapping:
#    Group original columns by their mapped bucket according to mapping.xlsx (by normalized header). 
#    The order of groups follows CORE_ORDER. Within each group keep the original relative order from the file.
grouped_cols = []
unmapped_cols = []

# build normalized lookup for original column names to preserve matching
orig_norm_to_col = {normalize(c): c for c in orig_after_style}

# For each core bucket in CORE_ORDER, collect original columns that map to it (preserve orig order)
for core in CORE_ORDER:
    # scan orig_after_style in original order and pick those whose mapping_norm_to_bucket maps to this core
    picked = []
    for col in orig_after_style:
        n = normalize(col)
        mapped_bucket = mapping_norm_to_bucket.get(n, "")
        if mapped_bucket == core:
            picked.append(col)
    grouped_cols.extend(picked)

# Any original columns that were not mapped (or mapped to other values) — keep them at the end preserving order
for col in orig_after_style:
    if col not in grouped_cols:
        unmapped_cols.append(col)

# Final original column order after reshuffle: grouped_cols (by CORE_ORDER) + unmapped_cols
reshuffled_original_cols = grouped_cols + unmapped_cols

# 4) Build output DataFrame rows: for each row, compute combined core messages
output_rows = []
for idx, row in df_no_reasons.iterrows():
    out_row = {}
    # copy original values (we'll assemble df later)
    for col in orig_cols_no_reason:
        out_row[col] = row.get(col, "")
    # initialize combined bucket messages
    for core in CORE_ORDER:
        out_row[core] = ""
    # parse bullets from summary and report
    summary_bullets = split_bullets(row.get(summary_col, ""))
    report_bullets = split_bullets(row.get(report_col, ""))
    # parse into (title,msg)
    summary_parsed = [parse_title_msg(b) for b in summary_bullets]
    report_parsed = [parse_title_msg(b) for b in report_bullets]
    # build report map keyed by normalized title -> list of messages
    report_map = {}
    for t, m in report_parsed:
        key = normalize(t) if t else normalize(m[:40])
        report_map.setdefault(key, []).append(m or "")
    # for each summary bullet, map to bucket using mapping_norm_to_bucket (match title against mapping)
    bucket_acc = defaultdict(list)
    for title, s_msg in summary_parsed:
        if not title:
            continue
        key = normalize(title)
        # prefer report msg if available
        rep_msg = ""
        if key in report_map and report_map[key]:
            rep_msg = report_map[key][0]
        else:
            # attempt to find in report_parsed via substring
            for rt, rm in report_parsed:
                hay = (rt or "") + " " + (rm or "")
                if title and title.lower() in hay.lower():
                    rep_msg = rm or ""
                    break
        final_msg = rep_msg if rep_msg else s_msg
        # map title to bucket via mapping (normalized match)
        bucket = mapping_norm_to_bucket.get(key, "")
        if not bucket:
            # try simple substring mapping: if any mapping header substring found in title
            matched = ""
            for hdr in mapping_norm_to_bucket.keys():
                if hdr and hdr in key:
                    matched = mapping_norm_to_bucket.get(hdr, "")
                    break
            bucket = matched or ""
        if not bucket:
            bucket = "Unmapped"
        bucket_acc[bucket].append(f"{title}: {final_msg}".strip())
    # also process report bullets that were not covered in summary
    for rt, rm in report_parsed:
        key = normalize(rt) if rt else normalize(rm[:40])
        # skip if already in bucket_acc (heuristic)
        already = False
        for msgs in bucket_acc.values():
            for m in msgs:
                if (rm and rm in m) or (rt and rt in m):
                    already = True
                    break
            if already:
                break
        if already:
            continue
        final_msg = rm or ""
        bucket = mapping_norm_to_bucket.get(key, "")
        if not bucket:
            # substring attempt
            matched = ""
            for hdr in mapping_norm_to_bucket.keys():
                if hdr and hdr in key:
                    matched = mapping_norm_to_bucket.get(hdr, "")
                    break
            bucket = matched or ""
        if not bucket:
            bucket = "Unmapped"
        bucket_acc[bucket].append(f"{rt if rt else key}: {final_msg}".strip())
    # set combined messages into out_row for each core
    for core in CORE_ORDER:
        msgs = bucket_acc.get(core, [])
        if msgs:
            out_row[core] = "\n".join(msgs)
        else:
            out_row[core] = ""
    # append out_row to list
    output_rows.append(out_row)

# Build final DataFrame with desired column order:
# styleid -> CORE_ORDER columns -> reshuffled_original_cols (which are original non-reason cols in mapped group order)
final_cols = [style_col] + CORE_ORDER + reshuffled_original_cols

# Ensure all final_cols are present in the out rows; if any original column missing (edge case), add empty
processed_df = pd.DataFrame(output_rows, columns=final_cols)

# show preview
st.success("Processing done — preview below (first 10 rows)")
st.dataframe(processed_df.head(10), use_container_width=True)

# provide download
out_bytes = build_excel_bytes(processed_df)
st.download_button(
    "Download collated Excel",
    data=out_bytes,
    file_name=f"myntra_collated_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

st.caption("Notes: 'Reason' columns were removed. Combined core columns are inserted after styleid and contain concatenated messages mapped to that bucket. Original columns remain (reshuffled grouped by mapping).")
