import streamlit as st
import pandas as pd
import html
import os
import re

st.set_page_config(page_title="Universal SAP LTMC Mapper", layout="wide")

st.title("🚀 Dynamic Universal SAP LTMC Data Mapper")
st.write("Upload any Excel workbook, select your source sheets, map fields dynamically via UI dropdowns, and compile to standard LTMC XML.")

# --- STEP 1: DYNAMIC WORKBOOK INSPECTION ---
st.subheader("📁 Step 1: Upload Source Workbook & Select Tabs")
uploaded_file = st.file_uploader("Upload Source Excel File (.xlsx)", type=["xlsx"])

if uploaded_file:
    try:
        # Dynamically inspect the file to discover tabs
        xl_file = pd.ExcelFile(uploaded_file)
        available_sheets = xl_file.sheet_names
        st.success(f"📖 System detected {len(available_sheets)} tabs in your workbook.")
        
        # Let the user choose which tabs represent Area 01 and Area 15 dynamically
        c_tabs1, c_tabs2 = st.columns(2)
        with c_tabs1:
            # Try to auto-detect a tab with 'IFRS' or default to the first tab
            default_ifrs_idx = next((i for i, s in enumerate(available_sheets) if "IFRS" in s.upper()), 0)
            area01_tab = st.selectbox("Select Tab for Depreciation Area 01 (e.g., IFRS / Tab1):", options=available_sheets, index=default_ifrs_idx)
        with c_tabs2:
            # Try to auto-detect a tab with 'GAAP' or default to the second tab
            default_igaap_idx = next((i for i, s in enumerate(available_sheets) if "GAAP" in s.upper() or "IGAAP" in s.upper()), min(1, len(available_sheets)-1))
            area15_tab = st.selectbox("Select Tab for Depreciation Area 15 (e.g., IGAAP / Tab2):", options=available_sheets, index=default_igaap_idx)

        # Allow user to specify if the primary tab has decorative title rows to skip
        skip_rows_01 = st.number_input("Number of header title rows to skip for Area 01 Tab (0 if headers are on Row 1):", min_value=0, value=1 if "IFRS" in area01_tab.upper() else 0)
        skip_rows_15 = st.number_input("Number of header title rows to skip for Area 15 Tab (0 if headers are on Row 1):", min_value=0, value=0)

        # --- STEP 2: DYNAMIC FIELD DISCOVERY & HEURISTIC PROPOSAL ---
        # Read the full data sheets based on runtime UI choices
        uploaded_file.seek(0)
        df_01 = pd.read_excel(uploaded_file, sheet_name=area01_tab, skiprows=skip_rows_01, dtype=str).dropna(how='all').fillna("")
        
        uploaded_file.seek(0)
        df_15 = pd.read_excel(uploaded_file, sheet_name=area15_tab, skiprows=skip_rows_15, dtype=str).dropna(how='all').fillna("")
        
        # Clean column spaces safely
        df_01.columns = [str(c).strip() for c in df_01.columns]
        df_15.columns = [str(c).strip() for c in df_15.columns]

        # Helper utility to generate user-friendly "Column letter + Header name" lists
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
                    display_headers.append(f"Column {letter} ({clean_name})")
            return display_headers

        cols_01_display = get_display_headers(df_01)
        cols_15_display = get_display_headers(df_15)

        st.markdown("---")
        st.subheader("📋 Step 2: Adaptive Source-to-Target Field Mapping Window")
        st.info("The system has analyzed your file and proposed selections based on text heuristics. Verify or alter them below:")

        # Intelligent fuzzy lookup matching function to auto-propose selections
        def propose_index(display_list, keywords, fallback_letter):
            for idx, name in enumerate(display_list):
                if any(k.lower() in name.lower() for k in keywords):
                    return idx
            for idx, name in enumerate(display_list):
                if name.startswith(f"Column {fallback_letter} ") or name == f"Column {fallback_letter}":
                    return idx
            return 0

        map_c1, map_c2 = st.columns(2)
        
        with map_c1:
            st.markdown(f"### 🌐 Area 01 Configuration ({area01_tab})")
            if_comp = st.selectbox("Company Code (COMPANY_CODE Target):", options=cols_01_display, index=propose_index(cols_01_display, ["cocd", "company", "comp"], "A"))
            if_asset = st.selectbox("Asset Number (ASSET_NUMBER Target Key):", options=cols_01_display, index=propose_index(cols_01_display, ["asset no", "asset_no", "asset number"], "C"))
            if_legacy = st.selectbox("Legacy Asset Number (LEGACY_ASSET Target):", options=cols_01_display, index=propose_index(cols_01_display, ["legacy"], "D"))
            if_acq = st.selectbox("Acquisition Value (ACQ_VALUE Target):", options=cols_01_display, index=propose_index(cols_01_display, ["cum. acq", "acq. val", "acquisition"], "AL"))
            if_dep = st.selectbox("Ordinary Depreciation (ORD_DEPR Target):", options=cols_01_display, index=propose_index(cols_01_display, ["accum. ord", "ordinary dep", "depreciation"], "AN"))

        with map_c2:
            st.markdown(f"### 🇮🇳 Area 15 Configuration ({area15_tab})")
            ig_asset = st.selectbox("Asset Number Cross-Match Key:", options=cols_15_display, index=propose_index(cols_15_display, ["asset no", "asset_no", "asset number"], "C"))
            ig_acq = st.selectbox("Acquisition Value (ACQ_VALUE Target):", options=cols_15_display, index=propose_index(cols_15_display, ["book_val", "acq", "value15"], "AI"))
            ig_dep = st.selectbox("Ordinary Depreciation (ORD_DEPR Target):", options=cols_15_display, index=propose_index(cols_15_display, ["accum_dep", "dep_area15", "depr15"], "AJ"))

        # --- STEP 3: MIGRATION DATA CONTEXT PARAMETERS ---
        st.markdown("---")
        st.subheader("⚙️ Step 3: Global Migration Parameters")
        gc1, gc2, gc3 = st.columns(3)
        with gc1:
            fiscal_year = st.text_input("Target Fiscal Year", value="2026")
        with gc2:
            currency_key = st.text_input("Local Currency Key", value="INR")
        with gc3:
            target_tab_name = st.text_input("LTMC Target Worksheet Name", value="cumulative values")

        # --- STEP 4: TRANSACTION COMPILATION ENGINE ---
        st.markdown("---")
        if st.button("Generate Production LTMC Upload XML", type="primary"):
            status_placeholder = st.empty()
            status_placeholder.info("Resolving user layout indices dynamically...")
            
            # Translate text dropdown strings back to exact positional column indices
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

            # Optimized memory hashing map for Area 15 Cross-Tab matching
            area15_lookup = {}
            for _, row_15 in df_15.iterrows():
                # Extract clean key format without decimal formatting errors (.0)
                raw_key = str(row_15.iloc[idx_15_asset]).strip().split('.')[0]
                if raw_key and raw_key != "nan" and raw_key != "":
                    area15_lookup[raw_key] = row_15

            total_records_written = 0
            
            with open(output_filename, "wb") as f:
                # SAP LTMC Excel XML Namespace Standards
                f.write(b'<?xml version="1.0" encoding="UTF-8"?>\n')
                f.write(b'<?mso-application progid="Excel.Sheet"?>\n')
                f.write(b'<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"\n')
                f.write(b' xmlns:o="urn:schemas-microsoft-com:office:office"\n')
                f.write(b' xmlns:x="urn:schemas-microsoft-com:office:excel"\n')
                f.write(b' xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">\n')
                f.write(f' <Worksheet ss:Name="{clean_xml(target_tab_name)}">\n'.encode('utf-8'))
                f.write(b'  <Table>\n')
                
                # Rigid Technical Header Row 
                f.write(b'   <Row>\n')
                for field in TARGET_FIELDS:
                    f.write(f'    <Cell><Data ss:Type="String">{field}</Data></Cell>\n'.encode('utf-8'))
                f.write(b'   </Row>\n')
                
                # Transform and pair row streams
                for idx, row_01 in df_01.iterrows():
                    comp_code = str(row_01.iloc[idx_01_comp]).strip()
                    asset_key = str(row_01.iloc[idx_01_asset]).strip().split('.')[0]
                    legacy_key = str(row_01.iloc[idx_01_legacy]).strip()
                    
                    # Filter out sub-headers, banners, or empty lines from data payload
                    if not asset_key or asset_key == "" or "Asset No" in asset_key or asset_key == "nan" or asset_key.startswith("Asset Master"):
                        continue
                    
                    val_01_acq = str(row_01.iloc[idx_01_acq]).strip()
                    val_01_dep = str(row_01.iloc[idx_01_dep]).strip()
                    
                    # Fast-hash matching into Area 15 dictionary lookup matrix
                    matched_row_15 = area15_lookup.get(asset_key, None)
                    if matched_row_15 is not None:
                        val_15_acq = str(matched_row_15.iloc[idx_15_acq]).strip()
                        val_15_dep = str(matched_row_15.iloc[idx_15_dep]).strip()
                    else:
                        val_15_acq = "0.00"
                        val_15_dep = "0.00"
                        
                    # --- WRITE ROW BLOCK A: AREA 01 (IFRS) ---
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
                    
                    # --- WRITE ROW BLOCK B: AREA 15 (IGAAP) ---
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
                status_placeholder.success(f"✅ Re-arrangement engine complete! Successfully built {total_records_written:,} structured target records.")
                with open(output_filename, "rb") as file_to_download:
                    st.download_button(
                        label="📥 Download Perfect LTMC XML File",
                        data=file_to_download,
                        file_name="LTMC_Universal_Fixed_Assets.xml",
                        mime="text/xml"
                    )
            else:
                status_placeholder.error("⚠️ The conversion returned 0 valid data rows. Check your tab selections or skipped row values.")
            os.remove(output_filename)
            
    except Exception as e:
        st.error(f"Critical Runtime Exception: {str(e)}")