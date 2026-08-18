import streamlit as st
import pandas as pd
import html
import os

st.set_page_config(page_title="Universal SAP LTMC Mapper", layout="wide")

st.title("🚀 Dynamic Universal SAP LTMC Data Mapper")
st.write("Upload any Excel workbook, select your source sheets, map fields dynamically via UI dropdowns, and compile to standard LTMC XML.")

# Fixed per spec — always these values in the target file
FISCAL_YEAR_FIXED = "2026"
CURRENCY_FIXED = "INR"

# Exact target structure, in order, matching the real LTMC "Cumulative Values" tab
TARGET_COLUMNS = [
    "Company Code",
    "External/Legacy Asset Number",
    "Asset Subnumber",
    "Depreciation Area",
    "Current Fiscal Year",
    "Cumulated Acquisition Value",
    "Cum. Revaluation on the Replacem. Value",
    "Cumulative Investment Grants",
    "Accumulated Ordinary Depreciation",
    "Cumulative Special Depreciation",
    "Cumulative Unplanned Depreciation",
    "Cumulative Transfer of Reserves",
    "Cumulative Interest",
    "Cumulative Revaluation of Ordinary Depr.",
    "Cumulative Down Payments",
    "Currency Key",
]

# --- New tabs: Master Details, Inventory, Origin — all sourced from I GAAP ONLY,
# mapped by fixed column letter per the spec (not by header-text matching, since
# these letters were given explicitly and verified against the real file). ---
def letter_to_idx(letter):
    idx = 0
    for c in letter:
        idx = idx * 26 + (ord(c.upper()) - 64)
    return idx - 1


# (source_letter, target_field_name) — source letters refer to columns in the
# I GAAP tab only, per the "Get the data from I GAAP tab only" instruction.
MASTER_DETAILS_FIELDS = [
    ("A", "Company code"),
    ("B", "External/Legacy Asset Number*"),
    ("E", "Asset Subnumber*"),
    ("F", "Asset Class*"),
    ("L", "Asset Description"),
    ("M", "Asset Description 2"),
    ("O", "Quantity"),
    ("P", "Base Unit of Measure"),
    # "Asset Is Managed Historically" is a fixed constant ("X"), not sourced from a column
    ("L", "Asset Main Number Text"),
]

INVENTORY_FIELDS = [
    ("A", "Company code"),
    ("B", "External/Legacy Asset Number*"),
    ("E", "Asset Subnumber*"),
]

ORIGIN_FIELDS = [
    ("A", "Company code"),
    ("B", "External/Legacy Asset Number*"),
    ("E", "Asset Subnumber*"),
]

ASSET_MANAGED_HISTORICALLY_FIXED = "X"


def write_letter_mapped_worksheet(f, worksheet_name, df_source, field_defs, clean_xml_fn, asset_col_letter="B"):
    """
    Writes one additional <Worksheet> block sourced purely from I GAAP (df_source),
    mapping fixed source column letters to target field names. Skips rows with a
    blank Asset Number (same validity rule used for the main Cumulative Values tab).
    Returns (rows_written, rows_skipped).
    """
    asset_idx = letter_to_idx(asset_col_letter)
    rows_written = 0
    rows_skipped = 0

    f.write(f' <Worksheet ss:Name="{clean_xml_fn(worksheet_name)}">\n'.encode('utf-8'))
    f.write(b'  <Table>\n')

    f.write(b'   <Row>\n')
    for _, target_name in field_defs:
        f.write(f'    <Cell><Data ss:Type="String">{clean_xml_fn(target_name)}</Data></Cell>\n'.encode('utf-8'))
    if worksheet_name == "Master Details":
        f.write(f'    <Cell><Data ss:Type="String">{clean_xml_fn("Asset Is Managed Historically")}</Data></Cell>\n'.encode('utf-8'))
    f.write(b'   </Row>\n')

    n_cols = df_source.shape[1]
    for _, row in df_source.iterrows():
        if asset_idx >= n_cols:
            break
        asset_key = str(row.iloc[asset_idx]).strip().split('.')[0]
        if not asset_key or asset_key == "nan" or "Asset No" in asset_key or asset_key.startswith("Asset Master"):
            rows_skipped += 1
            continue

        f.write(b'   <Row>\n')
        for source_letter, _ in field_defs:
            src_idx = letter_to_idx(source_letter)
            val = str(row.iloc[src_idx]).strip() if src_idx < n_cols else ""
            if val == "nan":
                val = ""
            f.write(f'    <Cell><Data ss:Type="String">{clean_xml_fn(val)}</Data></Cell>\n'.encode('utf-8'))
        if worksheet_name == "Master Details":
            f.write(f'    <Cell><Data ss:Type="String">{clean_xml_fn(ASSET_MANAGED_HISTORICALLY_FIXED)}</Data></Cell>\n'.encode('utf-8'))
        f.write(b'   </Row>\n')
        rows_written += 1

    f.write(b'  </Table>\n')
    f.write(b' </Worksheet>\n')
    return rows_written, rows_skipped


# --- STEP 1: DYNAMIC WORKBOOK INSPECTION ---
st.subheader("📁 Step 1: Upload Source Workbook & Select Tabs")
uploaded_file = st.file_uploader("Upload Source Excel File (.xlsx)", type=["xlsx"])


def detect_header_row(uploaded_file, sheet_name, max_scan=15):
    """
    Auto-detect which row is the real header row by scanning the first
    `max_scan` rows and picking the first row where most cells are
    non-blank text. Returns the number of rows to skip (0-indexed count
    of rows ABOVE the header).
    """
    uploaded_file.seek(0)
    preview = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=None, dtype=str, nrows=max_scan)
    preview = preview.fillna("")

    n_cols = preview.shape[1]
    for i in range(len(preview)):
        row = preview.iloc[i]
        non_blank = sum(1 for v in row if str(v).strip() != "")
        # Header row heuristic: most columns are non-blank text
        if n_cols > 0 and (non_blank / n_cols) >= 0.5:
            return i
    return 0  # fallback: assume no rows to skip


def get_display_headers(df):
    display_headers = []
    for i, col in enumerate(df.columns):
        letter = ""
        temp = i
        while temp >= 0:
            letter = chr(65 + (temp % 26)) + letter
            temp = (temp // 26) - 1
        clean_name = str(col).strip()
        if "Unnamed:" in clean_name or clean_name == "":
            display_headers.append(f"Column {letter}")
        else:
            # Collapse multi-line SAP field descriptions to first line only, for readability
            first_line = clean_name.split("\n")[0].strip()
            display_headers.append(f"Column {letter} ({first_line})")
    return display_headers


def propose_index(df, display_list, keywords, fallback_letter):
    """
    Propose a default column selection. Keywords are tried in priority
    order (keywords[0] is the most specific/preferred match) — ALL
    columns are checked against keywords[0] before moving to keywords[1],
    so a specific phrase like "cumulated acquisition value" always wins
    over a loosely-related column like "Fiscal Year of Original
    Acquisition" just because it contains the word "acquisition" and
    happens to appear earlier in the sheet.
    Columns that are entirely blank are skipped, since a blank column is
    never a valid mapping target even if its header text matches.
    """
    non_blank_mask = [(df.iloc[:, i].astype(str).str.strip() != "").any() for i in range(len(display_list))]

    # Try each keyword in priority order, across ALL columns, before moving to the next keyword
    for keyword in keywords:
        for idx, name in enumerate(display_list):
            if keyword.lower() in name.lower() and non_blank_mask[idx]:
                return idx
    # Fallback: same priority order, but allow blank columns (rare edge case)
    for keyword in keywords:
        for idx, name in enumerate(display_list):
            if keyword.lower() in name.lower():
                return idx
    # Final fallback: lettered column if non-blank
    for idx, name in enumerate(display_list):
        if (name.startswith(f"Column {fallback_letter} ") or name == f"Column {fallback_letter}") and non_blank_mask[idx]:
            return idx
    return 0


if uploaded_file:
    try:
        xl_file = pd.ExcelFile(uploaded_file)
        available_sheets = xl_file.sheet_names
        st.success(f"📖 System detected {len(available_sheets)} tabs in your workbook.")

        c_tabs1, c_tabs2 = st.columns(2)
        with c_tabs1:
            default_ifrs_idx = next((i for i, s in enumerate(available_sheets) if "IFRS" in s.upper()), 0)
            area01_tab = st.selectbox("Select Tab for Depreciation Area 01 (e.g., IFRS / Tab1):", options=available_sheets, index=default_ifrs_idx)
        with c_tabs2:
            default_igaap_idx = next((i for i, s in enumerate(available_sheets) if "GAAP" in s.upper() or "IGAAP" in s.upper()), min(1, len(available_sheets) - 1))
            area15_tab = st.selectbox("Select Tab for Depreciation Area 15 (e.g., IGAAP / Tab2):", options=available_sheets, index=default_igaap_idx)

        # --- AUTO-DETECTED HEADER ROWS ---
        auto_skip_01 = detect_header_row(uploaded_file, area01_tab)
        auto_skip_15 = detect_header_row(uploaded_file, area15_tab)

        st.caption("Header rows are auto-detected. Override below only if the wrong row was picked.")
        sk1, sk2 = st.columns(2)
        with sk1:
            skip_rows_01 = st.number_input(
                f"Header title rows to skip for {area01_tab} (auto-detected: {auto_skip_01}):",
                min_value=0, value=auto_skip_01
            )
        with sk2:
            skip_rows_15 = st.number_input(
                f"Header title rows to skip for {area15_tab} (auto-detected: {auto_skip_15}):",
                min_value=0, value=auto_skip_15
            )

        # --- STEP 2: FIELD DISCOVERY ---
        uploaded_file.seek(0)
        df_01 = pd.read_excel(uploaded_file, sheet_name=area01_tab, skiprows=skip_rows_01, dtype=str).dropna(how='all').fillna("")
        uploaded_file.seek(0)
        df_15 = pd.read_excel(uploaded_file, sheet_name=area15_tab, skiprows=skip_rows_15, dtype=str).dropna(how='all').fillna("")

        df_01.columns = [str(c).strip() for c in df_01.columns]
        df_15.columns = [str(c).strip() for c in df_15.columns]

        cols_01_display = get_display_headers(df_01)
        cols_15_display = get_display_headers(df_15)

        st.markdown("---")
        st.subheader("📋 Step 2: Adaptive Source-to-Target Field Mapping Window")
        st.info("The system has analyzed your file and proposed selections based on text heuristics, skipping any column that is entirely blank. Verify or alter them below:")

        map_c1, map_c2 = st.columns(2)

        with map_c1:
            st.markdown(f"### 🌐 Area 01 Configuration ({area01_tab})")
            if_comp = st.selectbox("Company Code (COMPANY_CODE Target):", options=cols_01_display, index=propose_index(df_01, cols_01_display, ["cocd", "company", "comp"], "A"))
            if_asset = st.selectbox("Asset Number (ASSET_NUMBER Target Key):", options=cols_01_display, index=propose_index(df_01, cols_01_display, ["legacy asset number", "asset no", "asset_no", "asset number"], "B"))
            if_sub = st.selectbox("Asset Subnumber (ASSET_SUBNUMBER Target):", options=cols_01_display, index=propose_index(df_01, cols_01_display, ["asset subnumber"], "E"))
            if_legacy = st.selectbox("Legacy Asset Number (LEGACY_ASSET Target):", options=cols_01_display, index=propose_index(df_01, cols_01_display, ["legacy"], "D"))
            if_acq = st.selectbox("Cumulated Acquisition Value (ACQ_VALUE Target):", options=cols_01_display, index=propose_index(df_01, cols_01_display, ["cumulated acquisition value", "cum. acq", "acq. val", "acquisition"], "AL"))
            if_dep = st.selectbox("Accumulated Ordinary Depreciation (ORD_DEPR Target):", options=cols_01_display, index=propose_index(df_01, cols_01_display, ["accumulated ordinary depreciation", "accum. ord", "ordinary dep", "depreciation"], "AN"))

        with map_c2:
            st.markdown(f"### 🇮🇳 Area 15 Configuration ({area15_tab})")
            ig_asset = st.selectbox("Asset Number Cross-Match Key:", options=cols_15_display, index=propose_index(df_15, cols_15_display, ["legacy asset number", "asset no", "asset_no", "asset number"], "B"))
            ig_acq = st.selectbox("Cumulated Acquisition Value (ACQ_VALUE Target):", options=cols_15_display, index=propose_index(df_15, cols_15_display, ["cumulated acquisition value", "book_val", "acq", "value15"], "AI"))
            ig_dep = st.selectbox("Accumulated Ordinary Depreciation (ORD_DEPR Target):", options=cols_15_display, index=propose_index(df_15, cols_15_display, ["accumulated ordinary depreciation", "accum_dep", "dep_area15", "depr15"], "AJ"))

        # --- STEP 3: GLOBAL PARAMS (fiscal year & currency are fixed per spec) ---
        st.markdown("---")
        st.subheader("⚙️ Step 3: Global Migration Parameters")
        gc1, gc2, gc3 = st.columns(3)
        with gc1:
            st.text_input("Target Fiscal Year (fixed)", value=FISCAL_YEAR_FIXED, disabled=True)
        with gc2:
            st.text_input("Local Currency Key (fixed)", value=CURRENCY_FIXED, disabled=True)
        with gc3:
            target_tab_name = st.text_input("LTMC Target Worksheet Name", value="Cumulative Values")

        # --- STEP 3.5: LIVE PREVIEW & VALIDATION ---
        st.markdown("---")
        st.subheader("🔍 Step 3.5: Preview & Validation")

        idx_01_asset = cols_01_display.index(if_asset)
        idx_15_asset = cols_15_display.index(ig_asset)

        keys_01_series = df_01.iloc[:, idx_01_asset].astype(str).str.strip().str.split('.').str[0]
        keys_15_series = df_15.iloc[:, idx_15_asset].astype(str).str.strip().str.split('.').str[0]

        keys_01 = set(k for k in keys_01_series if k and k != "nan")
        keys_15 = set(k for k in keys_15_series if k and k != "nan")
        overlap = keys_01 & keys_15

        blank_01 = int((keys_01_series.str.strip() == "").sum() + (keys_01_series == "nan").sum())
        blank_15 = int((keys_15_series.str.strip() == "").sum() + (keys_15_series == "nan").sum())
        dupes_01 = int(keys_01_series[keys_01_series != ""].duplicated().sum())

        vcol1, vcol2, vcol3 = st.columns(3)
        vcol1.metric(f"Unique keys in {area01_tab}", len(keys_01))
        vcol2.metric(f"Unique keys in {area15_tab}", len(keys_15))
        vcol3.metric("Matching keys (overlap)", len(overlap))

        if blank_01 > 0:
            st.warning(f"⚠️ {blank_01} rows in {area01_tab} have a blank Asset Number in the selected column — these rows will be skipped. If this number looks too high, you likely picked the wrong Asset Number column.")
        if blank_15 > 0:
            st.warning(f"⚠️ {blank_15} rows in {area15_tab} have a blank Asset Number in the selected column.")
        if dupes_01 > 0:
            st.warning(f"⚠️ {dupes_01} duplicate Asset Numbers found in {area01_tab}. Later duplicates will overwrite earlier lookups for {area15_tab} matching.")
        if len(keys_01) > 0 and len(overlap) == 0:
            st.error("❌ Zero assets match between the two tabs with the current Asset Number columns. Double-check both dropdowns point to the same kind of ID.")
        elif len(keys_01) > 0:
            match_rate = len(overlap) / len(keys_01) * 100
            st.success(f"✅ {match_rate:.1f}% of {area01_tab} assets have a match in {area15_tab}.")

        with st.expander("Preview first 5 mapped rows"):
            preview_rows = []
            idx_01_comp = cols_01_display.index(if_comp)
            idx_01_sub = cols_01_display.index(if_sub)
            idx_01_legacy = cols_01_display.index(if_legacy)
            idx_01_acq = cols_01_display.index(if_acq)
            idx_01_dep = cols_01_display.index(if_dep)
            idx_15_acq = cols_15_display.index(ig_acq)
            idx_15_dep = cols_15_display.index(ig_dep)

            area15_lookup_preview = {}
            for _, row_15 in df_15.iterrows():
                raw_key = str(row_15.iloc[idx_15_asset]).strip().split('.')[0]
                if raw_key and raw_key != "nan":
                    area15_lookup_preview[raw_key] = row_15

            shown = 0
            for _, row_01 in df_01.iterrows():
                if shown >= 5:
                    break
                asset_key = str(row_01.iloc[idx_01_asset]).strip().split('.')[0]
                if not asset_key or asset_key == "nan":
                    continue
                comp_code = str(row_01.iloc[idx_01_comp]).strip()
                legacy_key = str(row_01.iloc[idx_01_legacy]).strip()
                val_01_acq = str(row_01.iloc[idx_01_acq]).strip()
                val_01_dep = str(row_01.iloc[idx_01_dep]).strip()
                matched_15 = area15_lookup_preview.get(asset_key)
                val_15_acq = str(matched_15.iloc[idx_15_acq]).strip() if matched_15 is not None else "0.00 (no match)"
                val_15_dep = str(matched_15.iloc[idx_15_dep]).strip() if matched_15 is not None else "0.00 (no match)"
                preview_rows.append({
                    "Company Code": comp_code, "Legacy Asset": legacy_key, "Asset Number": asset_key,
                    f"{area01_tab} Acq Value": val_01_acq, f"{area01_tab} Depr": val_01_dep,
                    f"{area15_tab} Acq Value": val_15_acq, f"{area15_tab} Depr": val_15_dep,
                })
                shown += 1
            if preview_rows:
                st.dataframe(pd.DataFrame(preview_rows), use_container_width=True)
            else:
                st.warning("No valid rows found to preview with the current column selections.")

        # --- STEP 4: GENERATE XML (now writes the full 16-column Cumulative Values structure) ---
        st.markdown("---")
        if st.button("Generate Production LTMC Upload XML", type="primary"):
            status_placeholder = st.empty()

            if len(keys_01) > 0 and len(overlap) == 0:
                status_placeholder.error("⚠️ Aborted: 0 assets match between tabs. Fix the Asset Number column selections first (see validation above).")
            else:
                idx_01_comp = cols_01_display.index(if_comp)
                idx_01_asset = cols_01_display.index(if_asset)
                idx_01_sub = cols_01_display.index(if_sub)
                idx_01_legacy = cols_01_display.index(if_legacy)
                idx_01_acq = cols_01_display.index(if_acq)
                idx_01_dep = cols_01_display.index(if_dep)

                idx_15_asset = cols_15_display.index(ig_asset)
                idx_15_acq = cols_15_display.index(ig_acq)
                idx_15_dep = cols_15_display.index(ig_dep)

                output_filename = "LTMC_Perfect_Generated_Output.xml"

                def clean_xml(val):
                    return html.escape(str(val).strip())

                area15_lookup = {}
                for _, row_15 in df_15.iterrows():
                    raw_key = str(row_15.iloc[idx_15_asset]).strip().split('.')[0]
                    if raw_key and raw_key != "nan" and raw_key != "":
                        area15_lookup[raw_key] = row_15

                total_records_written = 0
                skipped_blank = 0
                skipped_banner = 0

                with open(output_filename, "wb") as f:
                    f.write(b'<?xml version="1.0" encoding="UTF-8"?>\n')
                    f.write(b'<?mso-application progid="Excel.Sheet"?>\n')
                    f.write(b'<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"\n')
                    f.write(b' xmlns:o="urn:schemas-microsoft-com:office:office"\n')
                    f.write(b' xmlns:x="urn:schemas-microsoft-com:office:excel"\n')
                    f.write(b' xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">\n')
                    f.write(f' <Worksheet ss:Name="{clean_xml(target_tab_name)}">\n'.encode('utf-8'))
                    f.write(b'  <Table>\n')

                    # Header row — now full 16 columns matching the real LTMC "Cumulative Values" tab
                    f.write(b'   <Row>\n')
                    for field in TARGET_COLUMNS:
                        f.write(f'    <Cell><Data ss:Type="String">{clean_xml(field)}</Data></Cell>\n'.encode('utf-8'))
                    f.write(b'   </Row>\n')

                    for idx, row_01 in df_01.iterrows():
                        comp_code = str(row_01.iloc[idx_01_comp]).strip()
                        asset_key = str(row_01.iloc[idx_01_asset]).strip().split('.')[0]
                        sub_num = str(row_01.iloc[idx_01_sub]).strip()
                        if not sub_num or sub_num == "nan":
                            sub_num = "0"

                        if not asset_key or asset_key == "nan":
                            skipped_blank += 1
                            continue
                        if "Asset No" in asset_key or asset_key.startswith("Asset Master"):
                            skipped_banner += 1
                            continue

                        val_01_acq = str(row_01.iloc[idx_01_acq]).strip()
                        val_01_dep = str(row_01.iloc[idx_01_dep]).strip()

                        matched_row_15 = area15_lookup.get(asset_key, None)
                        if matched_row_15 is not None:
                            val_15_acq = str(matched_row_15.iloc[idx_15_acq]).strip()
                            val_15_dep = str(matched_row_15.iloc[idx_15_dep]).strip()
                        else:
                            val_15_acq = "0.00"
                            val_15_dep = "0.00"

                        # Row block: Depreciation Area 01, then 15 — grouped per asset, per spec
                        for area, acq_val, dep_val in [("01", val_01_acq, val_01_dep), ("15", val_15_acq, val_15_dep)]:
                            row_values = [
                                comp_code,          # Company Code
                                asset_key,           # External/Legacy Asset Number
                                sub_num,              # Asset Subnumber
                                area,                 # Depreciation Area
                                FISCAL_YEAR_FIXED,    # Current Fiscal Year (always 2026)
                                acq_val,              # Cumulated Acquisition Value
                                "",                   # Cum. Revaluation on the Replacem. Value
                                "",                   # Cumulative Investment Grants
                                dep_val,              # Accumulated Ordinary Depreciation
                                "",                   # Cumulative Special Depreciation
                                "",                   # Cumulative Unplanned Depreciation
                                "",                   # Cumulative Transfer of Reserves
                                "",                   # Cumulative Interest
                                "",                   # Cumulative Revaluation of Ordinary Depr.
                                "",                   # Cumulative Down Payments
                                CURRENCY_FIXED,       # Currency Key (always INR)
                            ]
                            f.write(b'   <Row>\n')
                            for v in row_values:
                                f.write(f'    <Cell><Data ss:Type="String">{clean_xml(v)}</Data></Cell>\n'.encode('utf-8'))
                            f.write(b'   </Row>\n')
                            total_records_written += 1

                    f.write(b'  </Table>\n')
                    f.write(b' </Worksheet>\n')
                    f.write(b'</Workbook>\n')

                if total_records_written > 0:
                    msg = f"✅ Complete! Built {total_records_written:,} structured target records ({total_records_written // 2:,} assets × 2 depreciation areas)."
                    if skipped_blank or skipped_banner:
                        msg += f" (Skipped {skipped_blank} blank-key rows, {skipped_banner} banner/header rows.)"
                    status_placeholder.success(msg)
                    with open(output_filename, "rb") as file_to_download:
                        st.download_button(
                            label="📥 Download Perfect LTMC XML File",
                            data=file_to_download,
                            file_name="LTMC_Universal_Fixed_Assets.xml",
                            mime="text/xml"
                        )
                else:
                    status_placeholder.error(
                        f"⚠️ 0 valid data rows. {skipped_blank} rows had a blank Asset Number, "
                        f"{skipped_banner} looked like banner/header rows. "
                        f"Check the Asset Number column selection above."
                    )
                os.remove(output_filename)

        # --- STEP 5: ADDITIONAL LTMC TABS (Master Details / Inventory / Origin) ---
        # Sourced from the I GAAP tab ONLY (df_15). Target column POSITIONS are the
        # real, fixed SAP LTMC schema (never reordered/collapsed) — but which SOURCE
        # column feeds each target position is resolved dynamically by header-text
        # matching, so different clients' files (different column order, extra or
        # missing columns, renamed headers) all map correctly without breaking
        # alignment. Any target field we can't find in the source is left blank in
        # its correct position rather than shifting everything after it.
        st.markdown("---")
        st.subheader("📦 Step 5: Additional LTMC Tabs (Master Details / Inventory / Origin)")
        st.caption(f"All three tabs below pull directly from **{area15_tab}** only. Target column positions match the real SAP LTMC template exactly; source columns are located dynamically by header text so this adapts to different client file layouts.")

        def find_col_dynamic(df, keywords):
            """
            Locate a source column by header text, trying each keyword in
            priority order (most specific first) across ALL columns before
            trying the next keyword — so a specific phrase always wins over
            a loose partial match found earlier in the sheet. Returns None
            if nothing matches (caller leaves that field blank).
            """
            for kw in keywords:
                for i, c in enumerate(df.columns):
                    first_line = str(c).split("\n")[0].strip().rstrip("*").strip()
                    if kw.lower() == first_line.lower():
                        return i
            for kw in keywords:
                for i, c in enumerate(df.columns):
                    first_line = str(c).split("\n")[0].strip().rstrip("*").strip()
                    if kw.lower() in first_line.lower():
                        return i
            return None

        def get_val(row, idx, default=""):
            if idx is None:
                return default
            v = str(row.iloc[idx]).strip()
            return v if v and v != "nan" else default

        # Required anchor fields — these three must resolve or we can't build any of the 3 tabs
        anchor_comp = find_col_dynamic(df_15, ["Company Code"])
        anchor_asset = find_col_dynamic(df_15, ["External/Legacy Asset Number", "Asset Number"])
        anchor_sub = find_col_dynamic(df_15, ["Asset Subnumber"])

        if anchor_comp is None or anchor_asset is None:
            st.error(f"❌ Could not locate Company Code and/or Asset Number in {area15_tab}. These 3 tabs cannot be generated until resolved.")
        else:
            with st.expander("Resolved anchor columns (required for all 3 tabs)"):
                st.write(f"- Company Code → {df_15.columns[anchor_comp].split(chr(10))[0][:50]}")
                st.write(f"- Asset Number → {df_15.columns[anchor_asset].split(chr(10))[0][:50]}")
                st.write(f"- Asset Subnumber → {(df_15.columns[anchor_sub].split(chr(10))[0][:50] if anchor_sub is not None else 'not found — will default to 0')}")

            def write_positional_tab_xml(worksheet_name, field_specs, out_filename):
                """
                field_specs: list of (header_label, resolved_col_index_or_fixed_value_or_None)
                Each entry writes one column, in order, at its fixed target
                position — resolved_col_index_or_fixed_value_or_None can be:
                  - an int  -> pull from that source column index
                  - a str   -> literal fixed value (e.g. "X", "0")
                  - None    -> leave blank
                """
                def clean_xml(val):
                    return html.escape(str(val).strip())
                count = 0
                with open(out_filename, "wb") as f:
                    f.write(b'<?xml version="1.0" encoding="UTF-8"?>\n')
                    f.write(b'<?mso-application progid="Excel.Sheet"?>\n')
                    f.write(b'<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"\n')
                    f.write(b' xmlns:o="urn:schemas-microsoft-com:office:office"\n')
                    f.write(b' xmlns:x="urn:schemas-microsoft-com:office:excel"\n')
                    f.write(b' xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">\n')
                    f.write(f' <Worksheet ss:Name="{clean_xml(worksheet_name)}">\n'.encode("utf-8"))
                    f.write(b'  <Table>\n')
                    f.write(b'   <Row>\n')
                    for label, _ in field_specs:
                        f.write(f'    <Cell><Data ss:Type="String">{clean_xml(label)}</Data></Cell>\n'.encode("utf-8"))
                    f.write(b'   </Row>\n')
                    for _, src_row in df_15.iterrows():
                        asset_key = get_val(src_row, anchor_asset).split(".")[0]
                        if not asset_key:
                            continue
                        f.write(b'   <Row>\n')
                        for label, spec in field_specs:
                            if isinstance(spec, int):
                                val = get_val(src_row, spec)
                                if label == "External/Legacy Asset Number":
                                    val = val.split(".")[0]
                            elif isinstance(spec, str):
                                val = spec
                            else:
                                val = ""
                            f.write(f'    <Cell><Data ss:Type="String">{clean_xml(val)}</Data></Cell>\n'.encode("utf-8"))
                        f.write(b'   </Row>\n')
                        count += 1
                    f.write(b'  </Table>\n')
                    f.write(b' </Worksheet>\n')
                    f.write(b'</Workbook>\n')
                return count

            btn1, btn2, btn3 = st.columns(3)

            # --- Master Details (real SAP template: 12 columns, fixed positions) ---
            with btn1:
                st.markdown("**Master Details**")
                if st.button("Generate Master Details XML"):
                    md_class = find_col_dynamic(df_15, ["Asset Class"])
                    md_desc = find_col_dynamic(df_15, ["Asset Description"])
                    md_desc2 = find_col_dynamic(df_15, ["Asset Description 2"])
                    md_serial = find_col_dynamic(df_15, ["Serial Number"])
                    md_invnr = find_col_dynamic(df_15, ["Inventory Number"])
                    md_qty = find_col_dynamic(df_15, ["Quantity"])
                    md_uom = find_col_dynamic(df_15, ["Base Unit of Measure"])
                    md_mainnum = find_col_dynamic(df_15, ["Asset Main Number Text"])

                    field_specs = [
                        ("Company Code", anchor_comp),
                        ("External/Legacy Asset Number", anchor_asset),
                        ("Asset Subnumber", anchor_sub if anchor_sub is not None else "0"),
                        ("Asset Class", md_class),
                        ("Asset Description", md_desc),
                        ("Asset Description 2", md_desc2),
                        ("Serial Number", md_serial),
                        ("Inventory Number", md_invnr),
                        ("Quantity", md_qty),
                        ("Base Unit of Measure", md_uom),
                        ("Asset Is Managed Historically", "X"),  # fixed per spec
                        ("Asset Main Number Text", md_mainnum),
                    ]
                    with st.expander("Master Details — resolved source columns"):
                        for label, spec in field_specs:
                            if isinstance(spec, int):
                                st.write(f"- {label} → {df_15.columns[spec].split(chr(10))[0][:50]}")
                            elif isinstance(spec, str):
                                st.write(f"- {label} → fixed value '{spec}'")
                            else:
                                st.write(f"- {label} → ⚠️ not found in source, will be left blank")

                    out_fn = "LTMC_Master_Details.xml"
                    n = write_positional_tab_xml("Master Details", field_specs, out_fn)
                    st.success(f"✅ {n:,} rows written, {len(field_specs)} columns (matches real LTMC template positions).")
                    with open(out_fn, "rb") as fh:
                        st.download_button("📥 Download Master Details XML", data=fh, file_name=out_fn, mime="text/xml", key="dl_master")
                    os.remove(out_fn)

            # --- Inventory (real SAP template: 6 columns, fixed positions) ---
            with btn2:
                st.markdown("**Inventory**")
                if st.button("Generate Inventory XML"):
                    inv_lastdate = find_col_dynamic(df_15, ["Last Inventory Date"])
                    inv_suppl = find_col_dynamic(df_15, ["Supplementary Inventory Specifications"])
                    inv_ind = find_col_dynamic(df_15, ["Inventory Indicator"])

                    field_specs = [
                        ("Company Code", anchor_comp),
                        ("External/Legacy Asset Number", anchor_asset),
                        ("Asset Subnumber", anchor_sub if anchor_sub is not None else "0"),
                        ("Last Inventory Date", inv_lastdate),
                        ("Supplementary Inventory Specifications", inv_suppl),
                        ("Inventory Indicator", inv_ind),
                    ]
                    with st.expander("Inventory — resolved source columns"):
                        for label, spec in field_specs:
                            if isinstance(spec, int):
                                st.write(f"- {label} → {df_15.columns[spec].split(chr(10))[0][:50]}")
                            elif isinstance(spec, str):
                                st.write(f"- {label} → fixed value '{spec}'")
                            else:
                                st.write(f"- {label} → not present in source, left blank")

                    out_fn = "LTMC_Inventory.xml"
                    n = write_positional_tab_xml("Inventory", field_specs, out_fn)
                    st.success(f"✅ {n:,} rows written, {len(field_specs)} columns (matches real LTMC template positions).")
                    with open(out_fn, "rb") as fh:
                        st.download_button("📥 Download Inventory XML", data=fh, file_name=out_fn, mime="text/xml", key="dl_inventory")
                    os.remove(out_fn)

            # --- Origin (real SAP template: 14 columns, fixed positions) ---
            with btn3:
                st.markdown("**Origin**")
                if st.button("Generate Origin XML"):
                    org_vendor_no = find_col_dynamic(df_15, ["Account Number of Supplier", "Vendor Number"])
                    org_vendor = find_col_dynamic(df_15, ["Name of Asset Supplier", "Vendor"])
                    org_mfr = find_col_dynamic(df_15, ["Manufacturer of Asset", "Manufacturer"])
                    org_purchnew = find_col_dynamic(df_15, ["Indicator: Asset Purchased New", "Asset Purchased New"])
                    org_tradeid = find_col_dynamic(df_15, ["Company ID of Trading Partner", "Trading Partner"])
                    org_country = find_col_dynamic(df_15, ["Country/Region of Origin", "Country"])
                    org_typename = find_col_dynamic(df_15, ["Asset Type Name"])
                    org_acqdate = find_col_dynamic(df_15, ["Original Acquis. Date", "Original Acquisition Date"])
                    org_acqyr = find_col_dynamic(df_15, ["Fiscal Year of Original Acquisition"])
                    org_acqval = find_col_dynamic(df_15, ["Original Acquisition Value"])
                    org_inhouse = find_col_dynamic(df_15, ["In-House Production Percentage"])

                    field_specs = [
                        ("Company Code", anchor_comp),
                        ("External/Legacy Asset Number", anchor_asset),
                        ("Asset Subnumber", anchor_sub if anchor_sub is not None else "0"),
                        ("Account Number of Supplier", org_vendor_no),
                        ("Name of Asset Supplier", org_vendor),
                        ("Manufacturer of Asset", org_mfr),
                        ("Indicator: Asset Purchased New", org_purchnew),
                        ("Company ID of Trading Partner", org_tradeid),
                        ("Asset's Country/Region of Origin", org_country),
                        ("Asset Type Name", org_typename),
                        ("Original Acquis. Date of AuC/Transf. Ass", org_acqdate),
                        ("Fiscal Year of Original Acquisition", org_acqyr),
                        ("Original Acquisition Value", org_acqval),
                        ("In-House Production Percentage", org_inhouse),
                    ]
                    with st.expander("Origin — resolved source columns"):
                        for label, spec in field_specs:
                            if isinstance(spec, int):
                                st.write(f"- {label} → {df_15.columns[spec].split(chr(10))[0][:50]}")
                            elif isinstance(spec, str):
                                st.write(f"- {label} → fixed value '{spec}'")
                            else:
                                st.write(f"- {label} → not present in source, left blank")

                    out_fn = "LTMC_Origin.xml"
                    n = write_positional_tab_xml("Origin", field_specs, out_fn)
                    st.success(f"✅ {n:,} rows written, {len(field_specs)} columns (matches real LTMC template positions).")
                    with open(out_fn, "rb") as fh:
                        st.download_button("📥 Download Origin XML", data=fh, file_name=out_fn, mime="text/xml", key="dl_origin")
                    os.remove(out_fn)

    except Exception as e:
        st.error(f"Critical Runtime Exception: {str(e)}")