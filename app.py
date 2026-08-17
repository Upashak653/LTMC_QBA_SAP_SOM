import streamlit as st
import pandas as pd
import html
import os

st.set_page_config(page_title="Universal SAP LTMC Mapper", layout="wide")

st.title("🚀 Dynamic Universal SAP LTMC Data Mapper")
st.write("Upload any Excel workbook, select your source sheets, map fields dynamically via UI dropdowns, and compile to standard LTMC XML.")

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
    Propose a default column selection. Skips columns that are entirely
    blank, since a blank column is never a valid mapping target even if
    its header text matches the keywords.
    """
    non_blank_mask = [ (df.iloc[:, i].astype(str).str.strip() != "").any() for i in range(len(display_list)) ]

    # First pass: keyword match AND column has data
    for idx, name in enumerate(display_list):
        if any(k.lower() in name.lower() for k in keywords) and non_blank_mask[idx]:
            return idx
    # Second pass: keyword match regardless of blank (rare fallback)
    for idx, name in enumerate(display_list):
        if any(k.lower() in name.lower() for k in keywords):
            return idx
    # Third pass: fallback to lettered column if non-blank
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

        # --- AUTO-DETECTED HEADER ROWS (replaces manual guessing) ---
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
            if_legacy = st.selectbox("Legacy Asset Number (LEGACY_ASSET Target):", options=cols_01_display, index=propose_index(df_01, cols_01_display, ["legacy"], "D"))
            if_acq = st.selectbox("Acquisition Value (ACQ_VALUE Target):", options=cols_01_display, index=propose_index(df_01, cols_01_display, ["cum. acq", "acq. val", "acquisition"], "AL"))
            if_dep = st.selectbox("Ordinary Depreciation (ORD_DEPR Target):", options=cols_01_display, index=propose_index(df_01, cols_01_display, ["accum. ord", "ordinary dep", "depreciation"], "AN"))

        with map_c2:
            st.markdown(f"### 🇮🇳 Area 15 Configuration ({area15_tab})")
            ig_asset = st.selectbox("Asset Number Cross-Match Key:", options=cols_15_display, index=propose_index(df_15, cols_15_display, ["legacy asset number", "asset no", "asset_no", "asset number"], "B"))
            ig_acq = st.selectbox("Acquisition Value (ACQ_VALUE Target):", options=cols_15_display, index=propose_index(df_15, cols_15_display, ["book_val", "acq", "value15"], "AI"))
            ig_dep = st.selectbox("Ordinary Depreciation (ORD_DEPR Target):", options=cols_15_display, index=propose_index(df_15, cols_15_display, ["accum_dep", "dep_area15", "depr15"], "AJ"))

        # --- STEP 3: GLOBAL PARAMS ---
        st.markdown("---")
        st.subheader("⚙️ Step 3: Global Migration Parameters")
        gc1, gc2, gc3 = st.columns(3)
        with gc1:
            fiscal_year = st.text_input("Target Fiscal Year", value="2026")
        with gc2:
            currency_key = st.text_input("Local Currency Key", value="INR")
        with gc3:
            target_tab_name = st.text_input("LTMC Target Worksheet Name", value="cumulative values")

        # --- STEP 3.5: LIVE PREVIEW & VALIDATION (new) ---
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

        # --- STEP 4: GENERATE XML ---
        st.markdown("---")
        if st.button("Generate Production LTMC Upload XML", type="primary"):
            status_placeholder = st.empty()

            if len(keys_01) > 0 and len(overlap) == 0:
                status_placeholder.error("⚠️ Aborted: 0 assets match between tabs. Fix the Asset Number column selections first (see validation above).")
            else:
                idx_01_comp = cols_01_display.index(if_comp)
                idx_01_asset = cols_01_display.index(if_asset)
                idx_01_legacy = cols_01_display.index(if_legacy)
                idx_01_acq = cols_01_display.index(if_acq)
                idx_01_dep = cols_01_display.index(if_dep)

                idx_15_asset = cols_15_display.index(ig_asset)
                idx_15_acq = cols_15_display.index(ig_acq)
                idx_15_dep = cols_15_display.index(ig_dep)

                output_filename = "LTMC_Perfect_Generated_Output.xml"
                TARGET_FIELDS = ["COMPANY_CODE", "LEGACY_ASSET", "ASSET_NUMBER", "DEP_AREA", "FISCAL_YEAR", "ACQ_VALUE", "ORD_DEPR", "CURRENCY"]

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

                    f.write(b'   <Row>\n')
                    for field in TARGET_FIELDS:
                        f.write(f'    <Cell><Data ss:Type="String">{field}</Data></Cell>\n'.encode('utf-8'))
                    f.write(b'   </Row>\n')

                    for idx, row_01 in df_01.iterrows():
                        comp_code = str(row_01.iloc[idx_01_comp]).strip()
                        asset_key = str(row_01.iloc[idx_01_asset]).strip().split('.')[0]
                        legacy_key = str(row_01.iloc[idx_01_legacy]).strip()

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

                        f.write(b'   <Row>\n')
                        f.write(f'    <Cell><Data ss:Type="String">{clean_xml(comp_code)}</Data></Cell>\n'.encode('utf-8'))
                        f.write(f'    <Cell><Data ss:Type="String">{clean_xml(legacy_key)}</Data></Cell>\n'.encode('utf-8'))
                        f.write(f'    <Cell><Data ss:Type="String">{clean_xml(asset_key)}</Data></Cell>\n'.encode('utf-8'))
                        f.write(b'    <Cell><Data ss:Type="String">01</Data></Cell>\n')
                        f.write(f'    <Cell><Data ss:Type="String">{clean_xml(fiscal_year)}</Data></Cell>\n'.encode('utf-8'))
                        f.write(f'    <Cell><Data ss:Type="String">{clean_xml(val_01_acq)}</Data></Cell>\n'.encode('utf-8'))
                        f.write(f'    <Cell><Data ss:Type="String">{clean_xml(val_01_dep)}</Data></Cell>\n'.encode('utf-8'))
                        f.write(f'    <Cell><Data ss:Type="String">{clean_xml(currency_key)}</Data></Cell>\n'.encode('utf-8'))
                        f.write(b'   </Row>\n')
                        total_records_written += 1

                        f.write(b'   <Row>\n')
                        f.write(f'    <Cell><Data ss:Type="String">{clean_xml(comp_code)}</Data></Cell>\n'.encode('utf-8'))
                        f.write(f'    <Cell><Data ss:Type="String">{clean_xml(legacy_key)}</Data></Cell>\n'.encode('utf-8'))
                        f.write(f'    <Cell><Data ss:Type="String">{clean_xml(asset_key)}</Data></Cell>\n'.encode('utf-8'))
                        f.write(b'    <Cell><Data ss:Type="String">15</Data></Cell>\n')
                        f.write(f'    <Cell><Data ss:Type="String">{clean_xml(fiscal_year)}</Data></Cell>\n'.encode('utf-8'))
                        f.write(f'    <Cell><Data ss:Type="String">{clean_xml(val_15_acq)}</Data></Cell>\n'.encode('utf-8'))
                        f.write(f'    <Cell><Data ss:Type="String">{clean_xml(val_15_dep)}</Data></Cell>\n'.encode('utf-8'))
                        f.write(f'    <Cell><Data ss:Type="String">{clean_xml(currency_key)}</Data></Cell>\n'.encode('utf-8'))
                        f.write(b'   </Row>\n')
                        total_records_written += 1

                    f.write(b'  </Table>\n')
                    f.write(b' </Worksheet>\n')
                    f.write(b'</Workbook>\n')

                if total_records_written > 0:
                    msg = f"✅ Complete! Built {total_records_written:,} structured target records."
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

    except Exception as e:
        st.error(f"Critical Runtime Exception: {str(e)}")