# app.py
"""
Streamlit app: Collate Failure Summary + Failure Report into core buckets,
with robust concatenation of report messages and explicit reshuffle of columns.
Exports Excel with Reason columns grouped & collapsed.
"""
import streamlit as st
import pandas as pd
import re
from collections import defaultdict, OrderedDict
from io import BytesIO
from datetime import datetime
import os
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows

# try fuzzy lib
try:
    from rapidfuzz import process, fuzz
    RAPIDFUZZ = True
except Exception:
    RAPIDFUZZ = False

st.set_page_config(page_title="Myntra QC Collator", layout="wide")

# canonical core order (exact)
CORE_ORDER = [
    "Compliance", "Content", "Formatting", "Image Formatting",
    "Image", "Size", "Image Sequence"
]

# fuzzy threshold if rapidfuzz present
DEFAULT_FUZZ = 82

def normalize_title(t):
    if t is None:
        return ""
    s = str(t)
    s = s.strip()
    s = re.sub(r'[\t\r\n]+', ' ', s)
    s = re.sub(r'\s+', ' ', s)
    s = s.strip(" .:-–—")
    return s.lower()

def split_bullets(text):
    if pd.isna(text) or text is None:
        return []
    s = str(text)
    parts = re.split(r'\n|•|◦|·|•|-{2,}|—|–', s)
    return [p.strip() for p in parts if p.strip()]

def parse_reason_and_msg(bullet):
    if not bullet:
        return ("", "")
    parts = re.split(r'\s*[:\-—–]\s*', bullet, maxsplit=1)
    if len(parts) == 2:
        return parts[0].strip(), parts[1].strip()
    return parts[0].strip(), ""

def build_grouped_workbook(visible_df, full_df, reason_columns, mapping_used_df):
    wb = Workbook()
    ws = wb.active
    ws.title = "cleaned_output"
    # header = visible columns + reason columns (reason columns positioned after visible to be grouped)
    all_cols = list(visible_df.columns) + reason_columns
    ws.append(all_cols)
    for _, row in full_df.iterrows():
        row_vals = [row.get(c, "") for c in visible_df.columns] + [row.get(c, "") for c in reason_columns]
        ws.append(row_vals)
    # hide reason columns (grouped)
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
    # dropped columns
    ws_dropped = wb.create_sheet("dropped_reason_columns")
    for c in reason_columns:
        ws_dropped.append([c])
    return wb

st.title("Myntra QC — Collate & Clean (fixed reshuffle & concat)")

st.markdown("""
Upload your Output Excel. This version:
- Ensures core-bucket columns appear first (exact order),
- Concatenates Failure Report messages (preferred) with Failure Summary titles,
- Exports grouped/collapsed Reason columns.
""")

uploaded_xlsx = st.file_uploader("Upload Output Excel (XLSX)", type=["xlsx"])
c1, c2, c3 = st.columns(3)
with c1:
    style_col = st.text_input("styleid column name", value="styleid")
with c2:
    summary_col = st.text_input("Failure Summary column name", value="Failure Summary")
with c3:
    report_col = st.text_input("Failure Report column name", value="Failure Report")

# Load mapping.xlsx bundled
mapping_path_local = os.path.join(os.path.dirname(__file__), "mapping.xlsx")
if not os.path.exists(mapping_path_local):
    st.error("mapping.xlsx missing in repo. Put mapping.xlsx next to app.py")
    st.stop()

map_df = pd.read_excel(mapping_path_local).fillna("")
# ensure columns
if len(map_df.columns) >= 2:
    map_df = map_df.iloc[:, :2]
    map_df.columns = ["Header", "CoreBucket"]
else:
    map_df = pd.DataFrame(columns=["Header","CoreBucket"])

# editor (try modern first, fallback)
editor_df = None
try:
    editor_df = st.data_editor(map_df, num_rows="dynamic")
except Exception:
    try:
        editor_df = st.experimental_data_editor(map_df, num_rows="dynamic")
    except Exception:
        # fallback: show static table
        st.write("Mapping editor not supported in this runtime. Using bundled mapping.")
        editor_df = map_df.copy()

# build mapping dict, preserving mapping order (OrderedDict of core buckets -> list of headers)
mapping_hdr_to_bucket = {}
ordered_headers = []
for _, r in editor_df.iterrows():
    hdr = str(r["Header"]).strip()
    bucket = str(r["CoreBucket"]).strip() or "Unmapped"
    if hdr:
        mapping_hdr_to_bucket[hdr] = bucket
        ordered_headers.append(hdr)

# create normalized mapping lookup for exact & fuzzy
norm_map = {normalize_title(h): mapping_hdr_to_bucket[h] for h in mapping_hdr_to_bucket}
norm_keys = list(norm_map.keys())

# fuzzy threshold control
if RAPIDFUZZ:
    fuzz_threshold = st.sidebar.slider("Fuzzy threshold", 60, 95, DEFAULT_FUZZ)
else:
    fuzz_threshold = None
    st.sidebar.info("RapidFuzz not installed; fuzzy matching disabled. Add rapidfuzz to requirements to enable.")

if uploaded_xlsx is None:
    st.info("Upload the Output Excel to collate and clean.")
    st.stop()

# Read workbook (first sheet)
try:
    input_book = pd.read_excel(uploaded_xlsx, sheet_name=None, dtype=str)
    sheet_name = list(input_book.keys())[0]
    df = input_book[sheet_name].fillna("")
except Exception as e:
    st.error("Failed to read uploaded Excel: " + str(e))
    st.stop()

all_cols = list(df.columns)
reason_cols = [c for c in all_cols if re.search(r'\breason\b', str(c), flags=re.IGNORECASE)]

# ensure core columns exist
for core in CORE_ORDER + ["Unmapped"]:
    if core not in df.columns:
        df[core] = ""

# helper to find report message for a title
def find_best_report_message(title, report_parsed, report_map):
    # try normalized exact
    n = normalize_title(title)
    if n in report_map and report_map[n]:
        return report_map[n][0]
    # try substring in report_parsed messages/titles
    for rt, rm in report_parsed:
        hay = (rt or "") + " " + (rm or "")
        if title and title.lower() in hay.lower():
            return rm or ""
    # try fuzzy match (title -> report titles) if available
    if RAPIDFUZZ:
        report_titles_norm = [normalize_title(rt or rm[:30]) for rt,rm in report_parsed]
        if report_titles_norm:
            best = process.extractOne(n, report_titles_norm, scorer=fuzz.token_sort_ratio)
            if best and best[1] >= (fuzz_threshold or DEFAULT_FUZZ):
                idx = report_titles_norm.index(best[0])
                return report_parsed[idx][1] or ""
    return ""

mapped_log = []
unmapped_set = set()

# iterate rows
for idx, row in df.iterrows():
    summary_text = row.get(summary_col, "") if summary_col in df.columns else ""
    report_text = row.get(report_col, "") if report_col in df.columns else ""

    summary_bullets = split_bullets(summary_text)
    report_bullets = split_bullets(report_text)

    summary_parsed = [parse_reason_and_msg(b) for b in summary_bullets]
    report_parsed = [parse_reason_and_msg(b) for b in report_bullets]

    # build report_map normalized_title -> [messages]
    report_map = {}
    for rt, rm in report_parsed:
        key = normalize_title(rt) if rt else normalize_title(rm[:40])
        report_map.setdefault(key, []).append(rm or "")

    core_msgs = defaultdict(list)

    # process summary bullets (primary driver)
    for title, s_msg in summary_parsed:
        if not title:
            continue
        # pick the best message from report if available
        rep_msg = find_best_report_message(title, report_parsed, report_map)
        final_msg = rep_msg if rep_msg else s_msg
        # map title -> bucket
        n = normalize_title(title)
        bucket = None
        if n in norm_map:
            bucket = norm_map[n]
            mapped_log.append({"Title": title, "Bucket": bucket, "Method": "exact_norm"})
        else:
            # try fuzzy on mapping keys (if available)
            if RAPIDFUZZ and norm_keys:
                best = process.extractOne(n, norm_keys, scorer=fuzz.token_sort_ratio)
                if best and best[1] >= (fuzz_threshold or DEFAULT_FUZZ):
                    bucket = norm_map[best[0]]
                    mapped_log.append({"Title": title, "Bucket": bucket, "Method": f"fuzzy({best[1]})"})
        if not bucket:
            # attempt substring match with mapping headers
            found = False
            for hdr in mapping_hdr_to_bucket.keys():
                if hdr and hdr.lower() in title.lower():
                    bucket = mapping_hdr_to_bucket[hdr]
                    mapped_log.append({"Title": title, "Bucket": bucket, "Method": "substr_maphdr"})
                    found = True
                    break
            if not found:
                unmapped_set.add(title)
                bucket = "Unmapped"
                mapped_log.append({"Title": title, "Bucket": bucket, "Method": "unmapped"})
        core_msgs[bucket].append(f"{title}: {final_msg}".strip())

    # handle report bullets that didn't appear in summary (extra)
    for rt, rm in report_parsed:
        # check if message already captured
        rep_key = normalize_title(rt) if rt else normalize_title(rm[:40])
        already = False
        for msgs in core_msgs.values():
            for m in msgs:
                if (rm and rm in m) or (rt and rt in m):
                    already = True
                    break
            if already:
                break
        if already:
            continue
        # map rt to bucket similar to above
        n = normalize_title(rt or rm[:30])
        bucket = None
        if n in norm_map:
            bucket = norm_map[n]
            mapped_log.append({"Title": rt or rm[:30], "Bucket": bucket, "Method": "exact_norm_report"})
        else:
            if RAPIDFUZZ and norm_keys:
                best = process.extractOne(n, norm_keys, scorer=fuzz.token_sort_ratio)
                if best and best[1] >= (fuzz_threshold or DEFAULT_FUZZ):
                    bucket = norm_map[best[0]]
                    mapped_log.append({"Title": rt or rm[:30], "Bucket": bucket, "Method": f"fuzzy_report({best[1]})"})
        if not bucket:
            # substring mapping
            found = False
            for hdr in mapping_hdr_to_bucket.keys():
                if hdr and hdr.lower() in (rt or rm or "").lower():
                    bucket = mapping_hdr_to_bucket[hdr]
                    mapped_log.append({"Title": rt or rm[:30], "Bucket": bucket, "Method": "substr_maphdr_report"})
                    found = True
                    break
            if not found:
                bucket = "Unmapped"
                unmapped_set.add(rt or rm[:30])
                mapped_log.append({"Title": rt or rm[:30], "Bucket": bucket, "Method": "unmapped_report"})
        core_msgs[bucket].append(f"{rt or rm[:30]}: {rm}".strip())

    # write combined strings into df core buckets
    for core in CORE_ORDER + ["Unmapped"]:
        if core in core_msgs and core_msgs[core]:
            df.at[idx, core] = "\n".join(core_msgs[core])

# final visible df: reorder columns so CORE_ORDER first (only those present), then remaining columns excluding reason columns
present_cores = [c for c in CORE_ORDER if c in df.columns]
other_cols = [c for c in df.columns if c not in present_cores and c not in reason_cols]
visible_cols_ordered = present_cores + other_cols
visible_df = df[visible_cols_ordered].copy()

st.success("Collation complete. Preview (first 10 rows)")
st.dataframe(visible_df.head(10), use_container_width=True)

# mapping_used DataFrame for audit
mapping_used_df = pd.DataFrame(mapped_log)
if mapping_used_df.empty:
    # produce mapping table fallback
    mapping_used_df = pd.DataFrame([
        {"Header": h, "CoreBucket": mapping_hdr_to_bucket[h], "NormalizedHeader": normalize_title(h)}
        for h in mapping_hdr_to_bucket
    ])

# Build grouped workbook and provide download
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

if unmapped_set:
    st.warning(f"{len(unmapped_set)} unmapped reason titles found. See 'mapping_used' sheet in downloaded workbook for details.")
