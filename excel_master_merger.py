import streamlit as st
import pandas as pd
import difflib
from io import BytesIO

st.set_page_config(
    page_title="EPC Data Merger",
    page_icon="🔀",
    layout="wide"
)

# ─── STANDARD TARGET COLUMNS
STANDARD_COLS = [
    "APEX ID", "Sl. No.", "PD", "Business Format", "Format", "Sub Format",
    "Site Type", "Zone", "ZH", "State", "Site Name", "City", "Area (Sqft)",
    "RFC Date", "Actual Start Date", "Standard Duration",
    "Planned Finish Date", "Forecasted Finish Date", "Actual Finish Date",
    "HOTO Date", "Launch Date", "LAUNCHED / YTL", "FY",
    "AOP / NON AOP", "Planned (Month) Bucket", "Actual (Month) Bucket",
    "Store code", "EIC", "Cluster", "PM Head", "PM", "PM Planner",
    "Land ID", "9008 Store Code", "NP01 Site Code",
    "Change Note / Site Specific Duration"
]

# ─── KNOWN COLUMN NAME VARIANTS
KNOWN_MAPPINGS = {
    'apex id': 'APEX ID', 'apexid': 'APEX ID', 'apex_id': 'APEX ID',
    'sl. no.': 'Sl. No.', 'sl no': 'Sl. No.', 'sl.no.': 'Sl. No.',
    'serial no': 'Sl. No.', 'serial number': 'Sl. No.', 'sno': 'Sl. No.',
    'slno': 'Sl. No.', 'sr no': 'Sl. No.', 'sr. no.': 'Sl. No.',
    'pd': 'PD',
    'business format': 'Business Format', 'biz format': 'Business Format',
    'format': 'Format',
    'sub format': 'Sub Format', 'subformat': 'Sub Format', 'sub-format': 'Sub Format',
    'site type': 'Site Type', 'sitetype': 'Site Type', 'type': 'Site Type',
    'zone': 'Zone',
    'zh': 'ZH', 'zonal head': 'ZH', 'zonal_head': 'ZH',
    'state': 'State',
    'site name': 'Site Name', 'sitename': 'Site Name',
    'city': 'City',
    'area (sqft)': 'Area (Sqft)', 'area': 'Area (Sqft)',
    'sqft': 'Area (Sqft)', 'area sqft': 'Area (Sqft)',
    'rfc date': 'RFC Date', 'rfc': 'RFC Date',
    'actual start date': 'Actual Start Date', 'start date': 'Actual Start Date',
    'standard duration': 'Standard Duration', 'duration': 'Standard Duration',
    'planned finish date': 'Planned Finish Date', 'planned finish': 'Planned Finish Date',
    'forecasted finish date': 'Forecasted Finish Date',
    'forecasted finish': 'Forecasted Finish Date', 'forecast date': 'Forecasted Finish Date',
    'actual finish date': 'Actual Finish Date', 'actual finish': 'Actual Finish Date',
    'hoto date': 'HOTO Date', 'hoto': 'HOTO Date',
    'launch date': 'Launch Date',
    'launched / ytl': 'LAUNCHED / YTL', 'launched/ytl': 'LAUNCHED / YTL',
    'status': 'LAUNCHED / YTL', 'epc status': 'LAUNCHED / YTL',
    'fy': 'FY', 'financial year': 'FY',
    'aop / non aop': 'AOP / NON AOP', 'aop/non aop': 'AOP / NON AOP', 'aop': 'AOP / NON AOP',
    'planned (month) bucket': 'Planned (Month) Bucket', 'planned bucket': 'Planned (Month) Bucket',
    'actual (month) bucket': 'Actual (Month) Bucket', 'actual bucket': 'Actual (Month) Bucket',
    'store code': 'Store code', 'storecode': 'Store code',
    'eic': 'EIC', 'cluster': 'Cluster',
    'pm head': 'PM Head', 'pmhead': 'PM Head',
    'pm': 'PM', 'pm planner': 'PM Planner', 'pmplanner': 'PM Planner',
    'land id': 'Land ID', 'landid': 'Land ID',
    '9008 store code': '9008 Store Code',
    'np01 site code': 'NP01 Site Code',
    'change note / site specific duration': 'Change Note / Site Specific Duration',
}


def fuzzy_match_col(col_name):
    col_lower = col_name.lower().strip()
    if col_lower in KNOWN_MAPPINGS:
        return KNOWN_MAPPINGS[col_lower]
    std_lowers = [s.lower() for s in STANDARD_COLS]
    matches = difflib.get_close_matches(col_lower, std_lowers, n=1, cutoff=0.72)
    if matches:
        return STANDARD_COLS[std_lowers.index(matches[0])]
    return None


def detect_header_row(df_raw):
    """Find the row with the most string values — that's the header."""
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
    df.columns = [str(c).strip() for c in df.columns]
    df = df.loc[:, ~df.columns.str.startswith('Unnamed')]
    df = df.dropna(how='all')
    return df


# ─── HEADER
st.markdown("""
<div style="background: linear-gradient(135deg, #1a1a2e, #0f3460);
     padding: 1.8rem 2rem; border-radius: 12px; margin-bottom: 1.5rem;">
    <h1 style="color: white; margin: 0; font-size: 1.9rem;">🔀 EPC Data Merger</h1>
    <p style="color: #a0aec0; margin: 0.4rem 0 0;">
        Upload multiple Excel files → Select sheets → Auto-map columns →
        Merge on APEX ID → Download clean master
    </p>
</div>
""", unsafe_allow_html=True)

with st.expander("ℹ️ How this works", expanded=False):
    st.markdown("""
    1. **Upload** up to 5 Excel files received from different teams
    2. **Select the sheet** in each file that contains site data (e.g. "Backup", "Main Backup")
    3. **Review column mapping** — tool auto-detects matching columns, you can adjust
    4. Click **Merge** — data combined on APEX ID, exact duplicates removed
    5. **Resolve conflicts** — same APEX ID from two files with different data, you pick which to keep
    6. **Download** the clean merged Excel file → upload to the dashboard
    """)

# ════════════════════════════════════════
# STEP 1 — UPLOAD
# ════════════════════════════════════════
st.markdown("## Step 1 — Upload Files")
uploaded_files = st.file_uploader(
    "Upload up to 5 Excel files",
    type=["xlsx", "xls"],
    accept_multiple_files=True,
    help="These are the raw files received from different PM teams before merging"
)

if not uploaded_files:
    st.info("Upload one or more Excel files to begin.")
    st.stop()

if len(uploaded_files) > 5:
    st.error("Maximum 5 files. Please remove extras.")
    st.stop()

st.success(f"✅ {len(uploaded_files)} file(s) uploaded")

# ════════════════════════════════════════
# STEP 2 — SHEET SELECTION
# ════════════════════════════════════════
st.markdown("## Step 2 — Select Data Sheet")
st.caption("Pick the sheet that has the actual site data — usually 'Backup' or 'Main Backup'")

PREFERRED_SHEETS = ['Main Backup', 'Backup', 'Sheet1', 'Sheet5']

file_configs = []

for i, f in enumerate(uploaded_files):
    with st.expander(f"📄 {f.name}", expanded=True):
        try:
            f.seek(0)
            xl = pd.ExcelFile(f)
            sheets = xl.sheet_names

            default = sheets[0]
            for p in PREFERRED_SHEETS:
                if p in sheets:
                    default = p
                    break

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

            file_configs.append({
                "filename": f.name,
                "sheet": selected,
                "df": df,
                "file": f
            })

        except Exception as e:
            st.error(f"Could not read {f.name}: {e}")

if not file_configs:
    st.stop()

# ════════════════════════════════════════
# STEP 3 — COLUMN MAPPING
# ════════════════════════════════════════
st.markdown("## Step 3 — Column Mapping")
st.caption("Auto-detected mappings shown below. Change any that are wrong.")

all_mappings = []

for config in file_configs:
    df = config["df"]
    fname = config["filename"]

    with st.expander(f"🗂️ {fname} — {config['sheet']}", expanded=True):
        col_map = {}

        header = st.columns([2, 2, 1])
        header[0].markdown("**Original Column**")
        header[1].markdown("**Maps To (Standard)**")
        header[2].markdown("**Confidence**")

        options = ["— skip —"] + STANDARD_COLS

        for col in df.columns:
            matched = fuzzy_match_col(col)
            default_idx = 0
            confidence = ""
            if matched and matched in STANDARD_COLS:
                default_idx = STANDARD_COLS.index(matched) + 1
                confidence = "🟢 Auto" if col.lower().strip() in KNOWN_MAPPINGS else "🟡 Fuzzy"

            row = st.columns([2, 2, 1])
            row[0].write(col)
            chosen = row[1].selectbox(
                "",
                options,
                index=default_idx,
                key=f"map_{fname}_{col}",
                label_visibility="collapsed"
            )
            row[2].write(confidence or "⚪ Manual")

            if chosen != "— skip —":
                col_map[col] = chosen

        all_mappings.append({
            "filename": fname,
            "col_map": col_map,
            "df": df
        })

# ════════════════════════════════════════
# STEP 4 — MERGE
# ════════════════════════════════════════
st.markdown("## Step 4 — Merge")

if "final_df" in st.session_state:
    st.info("A merged file already exists below. Click 'Start Over' to re-merge.")

if st.button("🔀 Merge All Files Now", type="primary"):
    st.session_state.pop("final_df", None)
    st.session_state.pop("conflict_choices", None)

    dfs = []
    for m in all_mappings:
        df = m["df"].copy()
        col_map = m["col_map"]
        rename = {orig: std for orig, std in col_map.items() if orig in df.columns}
        df = df.rename(columns=rename)
        keep = [std for std in col_map.values() if std in df.columns]
        df = df[keep].copy()
        df = df.dropna(how="all")
        df["_source"] = m["filename"]
        dfs.append(df)

    combined = pd.concat(dfs, ignore_index=True, sort=False)

    if "APEX ID" not in combined.columns:
        st.error("❌ APEX ID not found. Make sure at least one file maps the APEX ID column.")
        st.stop()

    combined = combined.dropna(subset=["APEX ID"])
    combined["APEX ID"] = combined["APEX ID"].astype(str).str.strip()
    combined = combined[combined["APEX ID"].str.lower() != "nan"]

    # Remove exact duplicates
    combined = combined.drop_duplicates()

    # Find APEX ID conflicts
    conflict_ids = combined[combined.duplicated(subset=["APEX ID"], keep=False)]["APEX ID"].unique()

    st.session_state["combined"] = combined
    st.session_state["conflict_ids"] = conflict_ids

    if len(conflict_ids) == 0:
        final_df = combined.drop(columns=["_source"], errors="ignore")
        st.session_state["final_df"] = final_df
        st.success(f"✅ Merged — {len(final_df)} unique sites, zero conflicts.")
    else:
        st.warning(f"⚠️ {len(conflict_ids)} APEX ID(s) appear in multiple files with conflicting data. Resolve below.")

# ─── CONFLICT RESOLUTION
if "conflict_ids" in st.session_state and len(st.session_state["conflict_ids"]) > 0 and "final_df" not in st.session_state:
    combined = st.session_state["combined"]
    conflict_ids = st.session_state["conflict_ids"]

    st.markdown("### Resolve Conflicts")
    st.caption("Each card shows the same APEX ID from two different files. Pick which row to keep.")

    choices = {}
    for apex_id in conflict_ids:
        rows = combined[combined["APEX ID"] == apex_id].reset_index(drop=True)
        with st.expander(f"🔴 APEX ID: {apex_id}", expanded=False):
            st.dataframe(rows.drop(columns=["_source"], errors="ignore"), use_container_width=True)
            sources = rows["_source"].tolist()
            choice = st.radio(
                "Keep row from:",
                sources,
                key=f"res_{apex_id}"
            )
            choices[apex_id] = choice

    if st.button("✅ Confirm Choices & Generate Final File", type="primary"):
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
        st.success(f"✅ Final file ready — {len(final_df)} unique sites.")
        st.rerun()

# ════════════════════════════════════════
# STEP 5 — DOWNLOAD
# ════════════════════════════════════════
if "final_df" in st.session_state:
    st.markdown("## Step 5 — Download Merged Master File")

    final_df = st.session_state["final_df"]

    st.dataframe(final_df, use_container_width=True)

    m1, m2, m3 = st.columns(3)
    m1.metric("Total Sites", len(final_df))
    m2.metric("Columns", len(final_df.columns))
    m3.metric("Files Merged", len(uploaded_files))

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
