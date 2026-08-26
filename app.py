import streamlit as st
import pandas as pd

from config import (
    FISCAL_YEAR_FIXED,
    CURRENCY_FIXED,
    TARGET_COLUMNS,
    ASSET_MANAGED_HISTORICALLY_FIXED,
    MASTER_DETAILS_FIELDS,
    INVENTORY_FIELDS,
    ORIGIN_FIELDS,
    MODE_LANDING,
    MODE_MASTER,
    MODE_SUB
)
from utils import (
    parse_date_safe,
    find_col_dynamic,
    negate_num,
    letter_to_idx,
    clean_xml
)

st.set_page_config(page_title="Universal SAP LTMC Mapper", layout="wide")

# Initialize session state for mode navigation if not present
if "app_mode" not in st.session_state:
    st.session_state.app_mode = MODE_LANDING


def detect_header_row(uploaded_file, sheet_name, max_scan=15):
    uploaded_file.seek(0)
    preview = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=None, dtype=str, nrows=max_scan)
    preview = preview.fillna("")

    n_cols = preview.shape[1]
    for i in range(len(preview)):
        row = preview.iloc[i]
        non_blank = sum(1 for v in row if str(v).strip() != "")
        if n_cols > 0 and (non_blank / n_cols) >= 0.5:
            return i
    return 0


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
            first_line = clean_name.split("\n")[0].strip()
            display_headers.append(f"Column {letter} ({first_line})")
    return display_headers


def propose_index(df, display_list, keywords, fallback_letter):
    non_blank_mask = [(df.iloc[:, i].astype(str).str.strip() != "").any() for i in range(len(display_list))]

    for keyword in keywords:
        for idx, name in enumerate(display_list):
            if keyword.lower() in name.lower() and non_blank_mask[idx]:
                return idx
    for keyword in keywords:
        for idx, name in enumerate(display_list):
            if keyword.lower() in name.lower():
                return idx
    for idx, name in enumerate(display_list):
        if (name.startswith(f"Column {fallback_letter} ") or name == f"Column {fallback_letter}") and non_blank_mask[idx]:
            return idx
    return 0


# --- LANDING PAGE ROUTER ---
if st.session_state.app_mode == MODE_LANDING:
    st.title("🚀 Dynamic Universal SAP LTMC Data Mapper")
    st.write("Please select your target migration mode to proceed:")
    
    col1, col2 = st.columns(2)
    with col1:
        if st.button("🏢 Master Asset Migration Mode", use_container_width=True, type="primary"):
            st.session_state.app_mode = MODE_MASTER
            st.rerun()
            
    with col2:
        if st.button("📦 Sub Asset Migration Mode", use_container_width=True, type="primary"):
            st.session_state.app_mode = MODE_SUB
            st.rerun()

else:
    # --- ACTIVE PIPELINE WORKFLOW ---
    mode_title = "Master Asset Pipeline" if st.session_state.app_mode == MODE_MASTER else "Sub Asset Pipeline"
    
    if st.button("⬅️ Back to Mode Selection"):
        st.session_state.app_mode = MODE_LANDING
        st.rerun()
        
    st.title(f"🚀 Universal SAP LTMC Data Mapper — {mode_title}")
    st.write("Upload any Excel workbook, select your source sheets, map fields dynamically via UI dropdowns, and compile to standard LTMC XML.")

    # --- STEP 1: DYNAMIC WORKBOOK INSPECTION ---
    st.subheader("📁 Step 1: Upload Source Workbook & Select Tabs")
    uploaded_file = st.file_uploader("Upload Source Excel File (.xlsx)", type=["xlsx"])

    if uploaded_file:
        try:
            xl_file = pd.ExcelFile(uploaded_file)
            available_sheets = xl_file.sheet_names
            st.success(f"📖 System detected {len(available_sheets)} tabs in your workbook.")

            c_tabs1, c_tabs2 = st.columns(2)
            with c_tabs1:
                default_ifrs_idx = next((i for i, s in enumerate(available_sheets) if "IFRS" in s.upper()), 0)
                area01_tab = st.selectbox("Select Tab for Depreciation Area 01 (e.g., IFRS / Tab1):", options=available_sheets, index=default_ifrs_idx, key="sb_area01_tab")
            with c_tabs2:
                default_igaap_idx = next((i for i, s in enumerate(available_sheets) if "GAAP" in s.upper() or "IGAAP" in s.upper()), min(1, len(available_sheets) - 1))
                area15_tab = st.selectbox("Select Tab for Depreciation Area 15 (e.g., IGAAP / Tab2):", options=available_sheets, index=default_igaap_idx, key="sb_area15_tab")

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
                if_comp = st.selectbox("Company Code (COMPANY_CODE Target):", options=cols_01_display, index=propose_index(df_01, cols_01_display, ["cocd", "company", "comp"], "A"), key="sb_if_comp")
                if_asset = st.selectbox("Asset Number (ASSET_NUMBER Target Key):", options=cols_01_display, index=propose_index(df_01, cols_01_display, ["legacy asset number", "asset no", "asset_no", "asset number"], "B"), key="sb_if_asset")
                if_sub = st.selectbox("Asset Subnumber (ASSET_SUBNUMBER Target):", options=cols_01_display, index=propose_index(df_01, cols_01_display, ["asset subnumber"], "E"), key="sb_if_sub")
                if_legacy = st.selectbox("Legacy Asset Number (LEGACY_ASSET Target):", options=cols_01_display, index=propose_index(df_01, cols_01_display, ["legacy"], "D"), key="sb_if_legacy")
                if_acq = st.selectbox("Cumulated Acquisition Value (ACQ_VALUE Target):", options=cols_01_display, index=propose_index(df_01, cols_01_display, ["cumulated acquisition value", "cum. acq", "acq. val", "acquisition"], "AL"), key="sb_if_acq")
                if_dep = st.selectbox("Accumulated Ordinary Depreciation (ORD_DEPR Target):", options=cols_01_display, index=propose_index(df_01, cols_01_display, ["accumulated ordinary depreciation", "accum. ord", "ordinary dep", "depreciation"], "AN"), key="sb_if_dep")

            with map_c2:
                st.markdown(f"### 🇮🇳 Area 15 Configuration ({area15_tab})")
                ig_asset = st.selectbox("Asset Number Cross-Match Key:", options=cols_15_display, index=propose_index(df_15, cols_15_display, ["legacy asset number", "asset no", "asset_no", "asset number"], "B"), key="sb_ig_asset")
                ig_acq = st.selectbox("Cumulated Acquisition Value (ACQ_VALUE Target):", options=cols_15_display, index=propose_index(df_15, cols_15_display, ["cumulated acquisition value", "book_val", "acq", "value15"], "AI"), key="sb_ig_acq")
                ig_dep = st.selectbox("Accumulated Ordinary Depreciation (ORD_DEPR Target):", options=cols_15_display, index=propose_index(df_15, cols_15_display, ["accumulated ordinary depreciation", "accum_dep", "dep_area15", "depr15"], "AJ"), key="sb_ig_dep")

            idx_01_comp = cols_01_display.index(if_comp)
            idx_01_asset = cols_01_display.index(if_asset)
            idx_01_sub = cols_01_display.index(if_sub)
            idx_01_legacy = cols_01_display.index(if_legacy)
            idx_01_acq = cols_01_display.index(if_acq)
            idx_01_dep = cols_01_display.index(if_dep)
            idx_15_asset = cols_15_display.index(ig_asset)
            idx_15_acq = cols_15_display.index(ig_acq)
            idx_15_dep = cols_15_display.index(ig_dep)

            # --- FILTER DATAFRAMES BASED ON SELECTED MODE ---
            # Master Mode: Keep rows where subnumber is "0" or blank
            # Sub Mode: Keep rows where subnumber is NOT "0" and NOT blank
            sub_series_01 = df_01.iloc[:, idx_01_sub].astype(str).str.strip().str.split('.').str[0]
            if st.session_state.app_mode == MODE_MASTER:
                mask = (sub_series_01 == "") | (sub_series_01 == "0") | (sub_series_01 == "nan")
            else:
                mask = (sub_series_01 != "") & (sub_series_01 != "0") & (sub_series_01 != "nan")
            
            df_01 = df_01[mask].reset_index(drop=True)

            cap_date_01 = find_col_dynamic(df_01, ["Asset Capitalization Date"])
            cap_date_15 = find_col_dynamic(df_15, ["Asset Capitalization Date"])

            def cap_date_is_2026(row, cap_col):
                if cap_col is None:
                    return False
                d = parse_date_safe(row.iloc[cap_col])
                return pd.notna(d) and d.year == 2026

            st.markdown("---")
            st.subheader("⚙️ Step 3: Global Migration Parameters")
            gc1, gc2, gc3 = st.columns(3)
            with gc1:
                st.text_input("Target Fiscal Year (fixed)", value=FISCAL_YEAR_FIXED, disabled=True)
            with gc2:
                st.text_input("Local Currency Key (fixed)", value=CURRENCY_FIXED, disabled=True)
            with gc3:
                target_tab_name = st.text_input("LTMC Target Worksheet Name", value="Cumulative Values")

            st.markdown("---")
            st.subheader("🔍 Step 3.5: Preview & Validation")

            keys_01_series = df_01.iloc[:, idx_01_asset].astype(str).str.strip().str.split('.').str[0]
            keys_15_series = df_15.iloc[:, idx_15_asset].astype(str).str.strip().str.split('.').str[0]

            keys_01 = set(k for k in keys_01_series if k and k != "nan")
            keys_15 = set(k for k in keys_15_series if k and k != "nan")
            overlap = keys_01 & keys_15

            blank_01 = int((keys_01_series.str.strip() == "").sum() + (keys_01_series == "nan").sum())
            blank_15 = int((keys_15_series.str.strip() == "").sum() + (keys_15_series == "nan").sum())
            dupes_01 = int(keys_01_series[keys_01_series != ""].duplicated().sum())

            vcol1, vcol2, vcol3 = st.columns(3)
            vcol1.metric(f"Unique keys in {area01_tab} ({st.session_state.app_mode})", len(keys_01))
            vcol2.metric(f"Unique keys in {area15_tab}", len(keys_15))
            vcol3.metric("Matching keys (overlap)", len(overlap))

            if blank_01 > 0:
                st.warning(f"⚠️ {blank_01} rows in {area01_tab} have a blank Asset Number in the selected column.")
            if blank_15 > 0:
                st.warning(f"⚠️ {blank_15} rows in {area15_tab} have a blank Asset Number in the selected column.")
            if dupes_01 > 0:
                st.warning(f"⚠️ {dupes_01} duplicate Asset Numbers found in {area01_tab}.")
            if len(keys_01) > 0 and len(overlap) == 0:
                st.error("❌ Zero assets match between the two tabs with the current Asset Number columns.")
            elif len(keys_01) > 0:
                match_rate = len(overlap) / len(keys_01) * 100
                st.success(f"✅ {match_rate:.1f}% of {area01_tab} assets have a match in {area15_tab}.")

            st.markdown("---")
            if st.button("Generate Production LTMC Upload XML", type="primary"):
                status_placeholder = st.empty()

                if len(keys_01) > 0 and len(overlap) == 0:
                    status_placeholder.error("⚠️ Aborted: 0 assets match between tabs.")
                else:
                    output_filename = f"LTMC_{st.session_state.app_mode.capitalize()}_Output.xml"

                    area15_lookup = {}
                    for _, row_15 in df_15.iterrows():
                        raw_key = str(row_15.iloc[idx_15_asset]).strip().split('.')[0]
                        if raw_key and raw_key != "nan" and raw_key != "":
                            area15_lookup[raw_key] = row_15

                    total_records_written = 0
                    skipped_blank = 0
                    skipped_banner = 0
                    skipped_2026_cap = 0

                    with open(output_filename, "wb") as f:
                        f.write(b'<?xml version="1.0" encoding="UTF-8"?>\n')
                        f.write(b'<?mso-application progid="Excel.Sheet"?>\n')
                        f.write(b'<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"\n')
                        f.write(b' xmlns:o="urn:schemas-microsoft-com:office:office"\n')
                        f.write(b' xmlns:x="urn:schemas-microsoft-com:office:excel"\n')
                        f.write(b' xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">\n')
                        f.write(f' <Worksheet ss:Name="{clean_xml(target_tab_name)}">\n'.encode('utf-8'))
                        f.write(b'  <Table>\n')

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
                            val_01_dep = negate_num(row_01.iloc[idx_01_dep])

                            matched_row_15 = area15_lookup.get(asset_key, None)
                            if matched_row_15 is not None:
                                val_15_acq = str(matched_row_15.iloc[idx_15_acq]).strip()
                                val_15_dep = negate_num(matched_row_15.iloc[idx_15_dep])
                            else:
                                val_15_acq = "0.00"
                                val_15_dep = "0.00"

                            for area, acq_val, dep_val in [("01", val_01_acq, val_01_dep), ("15", val_15_acq, val_15_dep)]:
                                cap_source_row = row_01 if area == "01" else matched_row_15
                                cap_source_col = cap_date_01 if area == "01" else cap_date_15
                                if cap_source_row is not None and cap_date_is_2026(cap_source_row, cap_source_col):
                                    skipped_2026_cap += 1
                                    continue

                                row_values = [
                                    comp_code,
                                    asset_key,
                                    sub_num,
                                    area,
                                    FISCAL_YEAR_FIXED,
                                    acq_val,
                                    "",
                                    "",
                                    dep_val,
                                    "",
                                    "",
                                    "",
                                    "",
                                    "",
                                    "",
                                    CURRENCY_FIXED,
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
                        status_placeholder.success(f"✅ Complete! Built {total_records_written:,} structured target records for {st.session_state.app_mode}.")

        except Exception as e:
            st.error(f"An error occurred while processing the file: {e}")