# app.py
"""
Streamlit app with improved fuzzy matching (RapidFuzz) for mapping reason titles to core buckets.
Exports grouped Excel where Reason columns are collapsed.
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

# fuzzy matcher
from rapidfuzz import process, fuzz

st.set_page_config(page_title="Myntra QC Collator (fuzzy)", layout="wide")

CORE_ORDER = [
    "Compliance", "Content", "Formatting", "Image Formatting",
    "Image", "Size", "Image Sequence"
]

FUZZ_THRESHOLD = st.sidebar.slider("Fuzzy-match threshold", 70, 100, 82)

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
    if t is None:
        return ""
    s = str(t).strip()
    # Remove trailing punctuation, collapse whitespace, lowercase
    s = re.sub(r'[\t\r\n]+', ' ', s)
    s = re.sub(r'\s+', ' ', s)
    s = s.strip(" .:-–—")
    return s.lower()

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
    ws_map = wb.create_sheet("mapping_used")
    if mapping_used_df is None:
        mapping_used_df = pd.DataFrame(columns=["Header","CoreBucket","NormalizedHeader","Notes"])
    for r in dataframe_to_rows(mapping_used_df, index=False, header=True):
        ws_map.append(r)
    # dropped columns sheet
    ws_dropped = wb.create_sheet("dropped_reason_columns")
    for c in reason_columns:
        ws_dropped.append([c])
    return wb

st.title("Myntra QC — Collate & Clean (Fuzzy matching)")

st.markdown("""
This app uses fuzzy matching (RapidFuzz) to match reason titles from Failure Summary/Report to the mapping.
Adjust the fuzzy-match threshold in the left panel if you want stricter/looser matching.
""")

uploaded_xlsx = st.file_uploader("Upload Output Excel (XLSX)", type=["xlsx"])
col1, col2, col3 = st.columns(3)
with col1:
    style_col = st.text_input("styleid column name", value="styleid")
with col2:
    summary_col = st.text_input("Failure Summary column name", value="Failure Summary")
with col3:
    report_col = st.text_input("Failure Report column name", value="Failure Report")

# Load bundled mapping.xlsx
mapping_path_local = os.path.join(os.path.dirname(__file__), "mapping.xlsx")
if not os.path.exists(mapping_path_local):
    st.error("mapping.xlsx not found")
    st.stop()

map_df = pd.read_excel(mapping_path_local).fillna("")
cols = [c.strip() for c in map_df.columns]
if "Header" in cols and "CoreBucket" in cols:
    map_df = map_df[["Header", "CoreBucket"]]
else:
    if len(map_df.columns) >= 2:
        map_df.columns = ["Header", "CoreBucket"]
    else:
        map_df = pd.DataFrame(columns=["Header","CoreBucket"])

# simple inline editor fallback
try:
    editor_df = st.data_editor(map_df, num_rows="dynamic")
except Exception:
    try:
        editor_df = st.experimental_data_editor(map_df, num_rows="dynamic")
    except Exception:
        st.write("Mapping editor not available; using bundled mapping")
        editor_df = map_df.copy()

# Build normalized mapping dictionary
mapping = {}
normalized_map_keys = []
original_to_bucket = {}
for _, r in editor_df.iterrows():
    hdr = str(r["Header"]).strip()
    bucket = str(r["CoreBucket"]).strip() or "Unmapped"
    if hdr:
        n = normalize_title(hdr)
        mapping[n] = bucket
        normalized_map_keys.append(n)
        original_to_bucket[hdr] = bucket

if uploaded_xlsx is None:
    st.info("Upload the Output Excel to collate and clean.")
    st.stop()

# Read workbook
input_book = pd.read_excel(uploaded_xlsx, sheet_name=None, dtype=str)
sheet_name = list(input_book.keys())[0]
df = input_book[sheet_name].fillna("")

# find reason columns to preserve (but hide) and to drop from visible
all_cols = list(df.columns)
reason_cols = [c for c in all_cols if re.search(r'\breason\b', str(c), flags=re.IGNORECASE)]

# Ensure core columns exist
for core in CORE_ORDER:
    if core not in df.columns:
        df[core] = ""

unmapped_reasons = set()
mapped_log = []  # to record mapping used

# Precompute normalized_map_keys for fuzzy process
norm_keys = list(mapping.keys())

for idx, row in df.iterrows():
    summary_text = row.get(summary_col, "") if summary_col in df.columns else ""
    report_text = row.get(report_col, "") if report_col in df.columns else ""

    summary_bullets = split_bullets(summary_text)
    report_bullets = split_bullets(report_text)

    summary_parsed = [parse_reason_and_msg(b) for b in summary_bullets]
    report_parsed = [parse_reason_and_msg(b) for b in report_bullets]

    # build report map with normalized titles
    report_map = {}
    for title, msg in report_parsed:
        key = normalize_title(title) if title else normalize_title(msg[:30])
        report_map.setdefault(key, []).append(msg or "")

    core_bucket_msgs = defaultdict(list)
    used_report_indices = set()

    # Helper to attempt to map a title -> bucket using exact then fuzzy
    def map_title_to_bucket(orig_title):
        norm = normalize_title(orig_title)
        # exact normalized
        if norm in mapping:
            return mapping[norm], "exact_norm"
        # fuzzy match using RapidFuzz
        if norm_keys:
            best = process.extractOne(norm, norm_keys, scorer=fuzz.token_sort_ratio)
            if best:
                match_key, score, _ = best
                if score >= FUZZ_THRESHOLD:
                    return mapping[match_key], f"fuzzy({score})"
        return None, None

    # Iterate summary bullets and map
    for i, (title, s_msg) in enumerate(summary_parsed):
        if not title:
            continue
        bucket, how = map_title_to_bucket(title)
        if bucket:
            # try to fetch best detail message from report_map if exists
            rep_msgs = report_map.get(normalize_title(title), [])
            msg = rep_msgs[0] if rep_msgs else s_msg
            core_bucket_msgs[bucket].append(f"{title}: {msg}".strip())
            mapped_log.append({"Title": title, "Bucket": bucket, "Method": how})
        else:
            # fallback: try substring in report_parsed
            found = False
            for j, (r_title, r_msg) in enumerate(report_parsed):
                hay = (r_title or "") + " " + (r_msg or "")
                if title and title.lower() in hay.lower():
                    bucket2, how2 = map_title_to_bucket(r_title or title)
                    if bucket2:
                        core_bucket_msgs[bucket2].append(f"{title}: {r_msg or s_msg}".strip())
                        mapped_log.append({"Title": title, "Bucket": bucket2, "Method": f"substr->{how2}"})
                        found = True
                        break
            if not found:
                unmapped_reasons.add(title)
                core_bucket_msgs["Unmapped"].append(f"{title}: {s_msg}".strip())

    # Also handle report bullets not covered
    for j, (r_title, r_msg) in enumerate(report_parsed):
        norm = normalize_title(r_title or r_msg[:30])
        # skip if already added
        # naive check: if this exact message already exists among core_bucket_msgs skip
        present = False
        for msgs in core_bucket_msgs.values():
            if any((r_msg or "").strip() in m for m in msgs):
                present = True
                break
        if present:
            continue
        bucket, how = map_title_to_bucket(r_title or r_msg[:30])
        if bucket:
            core_bucket_msgs[bucket].append(f"{r_title or r_msg[:30]}: {r_msg}".strip())
            mapped_log.append({"Title": r_title or r_msg[:30], "Bucket": bucket, "Method": how})
        else:
            # fuzzy on message snippet
            if norm_keys:
                best = process.extractOne(norm, norm_keys, scorer=fuzz.token_sort_ratio)
                if best and best[1] >= FUZZ_THRESHOLD:
                    match_key = best[0]
                    core_bucket_msgs[mapping[match_key]].append(f"{r_title or r_msg[:30]}: {r_msg}".strip())
                    mapped_log.append({"Title": r_title or r_msg[:30], "Bucket": mapping[match_key], "Method": f"fuzzy_msg({best[1]})"})
                    continue
            unmapped_reasons.add(r_title or r_msg[:30])
            core_bucket_msgs["Unmapped"].append(f"{r_title or r_msg[:30]}: {r_msg}".strip())

    # write aggregated
    for core in CORE_ORDER + ["Unmapped"]:
        if core in core_bucket_msgs and core_bucket_msgs[core]:
            df.at[idx, core] = "\n".join(core_bucket_msgs[core])

# prepare visible df without reason columns
visible_df = df.drop(columns=reason_cols, errors='ignore')

st.success("Collation complete. Preview (first 10 rows)")
st.dataframe(visible_df.head(10), use_container_width=True)

# Prepare mapping_used log data
mapping_used_df = pd.DataFrame(mapped_log)
# If mapping_used_df is empty, fill with mapping table
if mapping_used_df.empty:
    mapping_used_df = pd.DataFrame([{"Header": k, "CoreBucket": mapping[k], "NormalizedHeader": k} for k in mapping])

# Build workbook and return bytes
wb = build_grouped_workbook(visible_df, df, reason_cols, mapping_used_df)
bio = BytesIO()
wb.save(bio)
out_bytes = bio.getvalue()

st.download_button(
    "Download cleaned & grouped Excel (Reason cols collapsed)",
    data=out_bytes,
    file_name=f"myntra_collated_grouped_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

if unmapped_reasons:
    st.warning(f"{len(unmapped_reasons)} unmapped reason titles were found. Download the exported workbook and inspect sheet 'mapping_used' or 'dropped_reason_columns' to update mapping.")
