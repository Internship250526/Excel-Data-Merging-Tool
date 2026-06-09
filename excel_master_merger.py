import streamlit as st
import pandas as pd
import difflib
from io import BytesIO

st.set_page_config(
    page_title="EPC Data Merger",
    page_icon="🔀",
    layout="wide"
)

# ─── KNOWN COLUMN NAME VARIANTS → STANDARD NAME
NORMALIZE = {
    'apex id': 'APEX ID', 'apexid': 'APEX ID', 'apex_id': 'APEX ID',
    'apex id2': 'APEX ID', 'apex id 2': 'APEX ID',
    'sl. no.': 'Sl. No.', 'sl no': 'Sl. No.', 'sl.no.': 'Sl. No.',
    'serial no': 'Sl. No.', 'serial number': 'Sl. No.',
    'sno': 'Sl. No.', 'sr no': 'Sl. No.', 'sr. no.': 'Sl. No.',
    'pd': 'PD',
    'business format': 'Business Format', 'biz format': 'Business Format',
    'format': 'Format',
    'sub format': 'Sub Format', 'subformat': 'Sub Format',
    'site type': 'Site Type',
    'zone': 'Zone',
    'zh': 'ZH', 'zonal head': 'ZH',
    'state': 'State',
    'site name': 'Site Name',
    'city': 'City',
    'area (sqft)': 'Area (Sqft)', 'area': 'Area (Sqft)', 'sqft': 'Area (Sqft)',
    'rfc date': 'RFC Date',
    'actual start date': 'Actual Start Date', 'start date': 'Actual Start Date',
    'standard duration': 'Standard Duration', 'duration': 'Standard Duration',
    'planned finish date': 'Planned Finish Date',
    'forecasted finish date': 'Forecasted Finish Date',
    'forecast date': 'Forecasted Finish Date',
    'actual finish date': 'Actual Finish Date',
    'hoto date': 'HOTO Date',
    'hoto status': 'HOTO Status',
    'launch date': 'Launch Date',
    'launched / ytl': 'LAUNCHED / YTL', 'launched/ytl': 'LAUNCHED / YTL',
    'status': 'LAUNCHED / YTL', 'epc status': 'LAUNCHED / YTL',
    'fy': 'FY', 'financial year': 'FY',
    'aop / non aop': 'AOP / NON AOP', 'aop/non aop': 'AOP / NON AOP',
    'planned (month) bucket': 'Planned (Month) Bucket',
    'actual (month) bucket': 'Actual (Month) Bucket',
    'store code': 'Store Code', 'storecode': 'Store Code',
    '9008 store code': '9008 Store Code', '9008 store': '9008 Store Code',
    'np01 site code': 'NP01 Site Code', 'np01 code': 'NP01 Site Code',
    'np01': 'NP01 Site Code',
    'eic': 'EIC', 'cluster': 'Cluster',
    'pm head': 'PM Head', 'pmhead': 'PM Head',
    'pm': 'PM', 'pm planner': 'PM Planner',
    'land id': 'Land ID',
    'change note / site specific duration': 'Change Note / Site Specific Duration',
    'rfc weekwise bucket': 'RFC Weekwise Bucket',
    'planned weekwise bucket': 'Planned Weekwise Bucket',
    'actual weekwise bucket': 'Actual Weekwise Bucket',
    'march target (date) epcho': 'March Target Date (EPCHO)',
}

def normalize_col(col):
    """Normalize a column name to standard. Keep as-is if unknown."""
    key = str(col).lower().strip()
    return NORMALIZE.get(key, str(col).strip())

def detect_header_row(df_raw):
    best_row, best_score = 0, 0
    for i in range(min(10, len(df_raw))):
        score = sum(1 for v in df_raw.iloc[i]
                    if isinstance(v, str) and len(v.strip()) > 1)
        if score > best_score:
            best_score, best_row = score, i
    return best_row

def load_sheet(file_obj, sheet_name):
    file_obj.seek(0)
    df_raw = pd.read_excel(file_obj, sheet_name=sheet_name, header=None)
    header_row = detect_header_row(df_raw)
    file_obj.seek(0)
    df = pd.read_excel(file_obj, sheet_name=sheet_name, header=header_row)
    # Normalize column names
    df.columns = [normalize_col(c) for c in df.columns]
    # Drop fully unnamed/empty columns
    df = df.loc[:, ~df.columns.str.lower().str.startswith('unnamed')]
    df = df.loc[:, ~df.columns.str.lower().str.match(r'^column\d*$')]
    df = df.dropna(how='all')
    # Handle duplicate column names after normalization — keep first occurrence
    seen = {}
    new_cols = []
    for col in df.columns:
        if col not in seen:
            seen[col] = 0
            new_cols.append(col)
        else:
            seen[col] += 1
            new_cols.append(f"{col}._dup{seen[col]}")
    df.columns = new_cols
    # Drop the duplicate suffix columns
    df = df.loc[:, ~df.columns.str.contains('._dup')]
    return df

# ─── HEADER
st.markdown("""
<div style="background: linear-gradient(135deg, #1a1a2e, #0f3460);
     padding: 1.8rem 2rem; border-radius: 12px; margin-bottom: 1.5rem;">
    <h1 style="color: white; margin: 0; font-size: 1.9rem;">🔀 EPC Data Merger</h1>
    <p style="color: #a0aec0; margin: 0.4rem 0 0 0;">
        Upload multiple Excel files → Select sheets → Auto-merge on APEX ID → Download unified master
    </p>
</div>
""", unsafe_allow_html=True)

with st.expander("ℹ️ How this works", expanded=False):
    st.markdown("""
    1. **Upload** up to 5 Excel files from different teams
    2. **Select the sheet** in each file that has the site data
    3. **Click Merge** — app auto-normalizes column names and takes the union of all columns
    4. **Resolve conflicts** — same APEX ID from two files with different data, you pick which to keep
    5. **Download** the clean merged Excel — ready to upload into the dashboard

    Every column from every file is kept. No column is dropped. If a site doesn't have data for a column, that cell is blank.
    """)

# ════════════════════════════════
# STEP 1 — UPLOAD
# ════════════════════════════════
st.markdown("## Step 1 — Upload Files")
uploaded_files = st.file_uploader(
    "Upload up to 5 Excel files",
    type=["xlsx", "xls"],
    accept_multiple_files=True
)

if not uploaded_files:
    st.info("Upload one or more Excel files to begin.")
    st.stop()

if len(uploaded_files) > 5:
    st.error("Maximum 5 files. Please remove extras.")
    st.stop()

st.success(f"✅ {len(uploaded_files)} file(s) uploaded")

# ════════════════════════════════
# STEP 2 — SHEET SELECTION
# ════════════════════════════════
st.markdown("## Step 2 — Select Data Sheet")
st.caption("Pick the sheet with site data — usually 'Backup' or 'Main Backup'")

PREFERRED = ['Main Backup', 'Backup', 'Sheet1']
file_configs = []

for i, f in enumerate(uploaded_files):
    with st.expander(f"📄 {f.name}", expanded=True):
        try:
            f.seek(0)
            sheets = pd.ExcelFile(f).sheet_names
            default = next((p for p in PREFERRED if p in sheets), sheets[0])

            selected = st.selectbox(
                "Sheet containing site data",
                sheets,
                index=sheets.index(default),
                key=f"sheet_{i}"
            )

            f.seek(0)
            df = load_sheet(f, selected)

            c1, c2, c3 = st.columns(3)
            c1.metric("Rows", len(df))
            c2.metric("Columns", len(df.columns))
            c3.metric("Sheet", selected)

            st.dataframe(df.head(3), use_container_width=True)

            # Show any column name normalizations applied
            f.seek(0)
            raw_cols = pd.read_excel(f, sheet_name=selected, nrows=0).columns.tolist()
            changes = [(str(r).strip(), normalize_col(r))
                       for r in raw_cols
                       if normalize_col(r) != str(r).strip()
                       and not str(r).lower().startswith('unnamed')]
            if changes:
                with st.expander(f"🔄 {len(changes)} column name(s) auto-normalized"):
                    for orig, norm in changes:
                        st.markdown(f"- `{orig}` → **{norm}**")

            file_configs.append({"filename": f.name, "sheet": selected, "df": df})

        except Exception as e:
            st.error(f"Could not read {f.name}: {e}")

if not file_configs:
    st.stop()

# ════════════════════════════════
# STEP 3 — MERGE
# ════════════════════════════════
st.markdown("## Step 3 — Merge")

# Show what the union will look like
all_cols = []
for config in file_configs:
    all_cols.extend(config["df"].columns.tolist())
unique_cols = list(dict.fromkeys(all_cols))  # preserve order, remove duplicates

with st.expander(f"📋 Preview — {len(unique_cols)} unique columns in final merged file"):
    st.write(unique_cols)

if st.button("🔀 Merge All Files Now", type="primary"):
    st.session_state.pop("final_df", None)
    st.session_state.pop("combined", None)
    st.session_state.pop("conflict_ids", None)

    dfs = []
    for config in file_configs:
        df = config["df"].copy()
        df["_source"] = config["filename"]
        df = df.dropna(how="all")
        dfs.append(df)

    combined = pd.concat(dfs, ignore_index=True, sort=False)

    if "APEX ID" not in combined.columns:
        st.error("❌ APEX ID column not found in any file. Cannot merge without a primary key.")
        st.stop()

    combined = combined.dropna(subset=["APEX ID"])
    combined["APEX ID"] = combined["APEX ID"].astype(str).str.strip()
    combined = combined[~combined["APEX ID"].str.lower().isin(["nan", "none", ""])]

    # Remove exact duplicate rows
    combined = combined.drop_duplicates(
        subset=[c for c in combined.columns if c != "_source"]
    )

    # ── DATA CLEANING ──────────────────────────────────────────
    # 1. Fix date columns: strip time, remove 1900 epoch artifacts
    for col in combined.columns:
        try:
            if pd.api.types.is_datetime64_any_dtype(combined[col]):
                combined[col] = combined[col].where(combined[col].dt.year > 1900, other=pd.NaT)
                combined[col] = combined[col].dt.strftime('%d-%m-%Y').where(combined[col].notna(), other='')
            elif combined[col].dtype == object:
                converted = pd.to_datetime(combined[col], errors='coerce', dayfirst=True)
                if converted.notna().sum() > len(combined) * 0.1:
                    converted = converted.where(converted.dt.year > 1900, other=pd.NaT)
                    combined[col] = converted.dt.strftime('%d-%m-%Y').where(converted.notna(), other='')
        except Exception:
            pass

    # 2. Drop columns entirely empty or all blank/nan strings
    combined = combined.dropna(axis=1, how='all')
    mask = (combined.astype(str).isin(['', 'nan', 'None', 'NaT', 'NaN'])).all()
    combined = combined.loc[:, ~mask]

    # 3. Drop entirely empty rows
    data_cols = [c for c in combined.columns if c != "_source"]
    combined = combined.dropna(subset=data_cols, how='all')
    # ───────────────────────────────────────────────────────────

    # Find APEX ID conflicts
    conflict_ids = combined[
        combined.duplicated(subset=["APEX ID"], keep=False)
    ]["APEX ID"].unique()

    st.session_state["combined"] = combined
    st.session_state["conflict_ids"] = conflict_ids

    if len(conflict_ids) == 0:
        final_df = combined.drop(columns=["_source"], errors="ignore")
        st.session_state["final_df"] = final_df
        st.success(f"✅ Merged — **{len(final_df)} unique sites**, **{len(final_df.columns)} columns**, zero conflicts.")
    else:
        st.warning(f"⚠️ **{len(conflict_ids)} APEX ID(s)** appear in multiple files with conflicting data. Resolve below.")

# ─── CONFLICT RESOLUTION
if ("conflict_ids" in st.session_state
        and len(st.session_state["conflict_ids"]) > 0
        and "final_df" not in st.session_state):

    combined = st.session_state["combined"]
    conflict_ids = st.session_state["conflict_ids"]

    st.markdown("### ⚠️ Resolve Conflicts")
    st.caption("Same APEX ID found in multiple files with different data. Pick which row to keep.")

    choices = {}
    for apex_id in conflict_ids:
        rows = combined[combined["APEX ID"] == apex_id].reset_index(drop=True)
        with st.expander(f"🔴 APEX ID: {apex_id}", expanded=False):
            st.dataframe(
                rows.drop(columns=["_source"], errors="ignore"),
                use_container_width=True
            )
            sources = rows["_source"].tolist()
            choice = st.radio(f"Keep row from:", sources, key=f"res_{apex_id}")
            choices[apex_id] = choice

    if st.button("✅ Confirm & Generate Final File", type="primary"):
        final_rows = []
        for apex_id in combined["APEX ID"].unique():
            rows = combined[combined["APEX ID"] == apex_id]
            if apex_id in choices:
                row = rows[rows["_source"] == choices[apex_id]].iloc[0]
            else:
                row = rows.iloc[0]
            final_rows.append(row)

        final_df = pd.DataFrame(final_rows).drop(columns=["_source"], errors="ignore")
        final_df = final_df.reset_index(drop=True)
        st.session_state["final_df"] = final_df
        st.success(f"✅ Done — {len(final_df)} unique sites.")
        st.rerun()

# ════════════════════════════════
# STEP 4 — DOWNLOAD
# ════════════════════════════════
if "final_df" in st.session_state:
    st.markdown("## Step 4 — Download")

    final_df = st.session_state["final_df"]

    m1, m2, m3 = st.columns(3)
    m1.metric("Total Sites", len(final_df))
    m2.metric("Total Columns", len(final_df.columns))
    m3.metric("Files Merged", len(uploaded_files))

    st.dataframe(final_df, use_container_width=True)

    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        final_df.to_excel(writer, index=False, sheet_name="Main Backup")
    buffer.seek(0)

    st.download_button(
        "📥 Download Merged Master Excel",
        data=buffer,
        file_name="merged_master.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.caption("Upload this file to the EPC Site Launch Tracker dashboard.")

    if st.button("🔄 Start Over"):
        for key in ["final_df", "combined", "conflict_ids"]:
            st.session_state.pop(key, None)
        st.rerun()
