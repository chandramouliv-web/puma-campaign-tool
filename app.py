import streamlit as st
import pandas as pd
import io
import re
import zipfile
from datetime import datetime, time, timedelta

# Set up page configuration
st.set_page_config(page_title="Marketplace Price Automator", page_icon="🚀", layout="wide")

# =================================================================
# ⚙️ HELPER FUNCTIONS & CLEANING PLUGINS
# =================================================================

def normalize_sku(sku):
    if pd.isna(sku) or not sku:
        return ""
    return str(sku).strip().replace('-', '_').lower()

def clean_id_str(val):
    if pd.isna(val):
        return None
    if isinstance(val, float):
        return str(int(val)) if val.is_integer() else str(val)
    s = str(val).strip()
    if re.match(r"^-?\d+\.0+$", s):
        s = s.split(".")[0]
    return s

def get_clean_headers_and_df(uploaded_file):
    try:
        uploaded_file.seek(0)
        if uploaded_file.name.endswith('.csv'):
            full_df = pd.read_csv(uploaded_file, header=None)
        else:
            full_df = pd.read_excel(uploaded_file, header=None)

        row3_names = full_df.iloc[2].fillna("").astype(str).tolist()

        def get_excel_col_letter(n):
            result = ""
            while n > 0:
                n, remainder = divmod(n - 1, 26)
                result = chr(65 + remainder) + result
            return result

        clean_headers = []
        for index, name in enumerate(row3_names):
            clean_name = name.strip()
            col_letter = get_excel_col_letter(index + 1)
            if not clean_name or "Unnamed:" in clean_name or clean_name == "nan":
                clean_name = "Blank Column"
            clean_headers.append(f"[{col_letter}] {clean_name}")

        full_df.columns = clean_headers
        full_df = full_df.iloc[3:].reset_index(drop=True)
        return clean_headers, full_df
    except Exception as e:
        st.error(f"Error processing layout headers from {uploaded_file.name}: {e}")
        return [], None

def read_full_file_standard(uploaded_file):
    uploaded_file.seek(0)
    if uploaded_file.name.endswith('.csv'):
        return pd.read_csv(uploaded_file)
    return pd.read_excel(uploaded_file)

def _read_shopee_zip(uploaded_file):
    dfs = []
    file_bytes = uploaded_file.read()
    with zipfile.ZipFile(io.BytesIO(file_bytes)) as zf:
        names = sorted(n for n in zf.namelist() if n.endswith(".xlsx"))
        if not names:
            raise ValueError("No active .xlsx data files located inside the uploaded ZIP bundle.")
        bar = st.progress(0, text="Reading Shopee export files…")
        for i, name in enumerate(names):
            with zf.open(name) as f:
                dfs.append(pd.read_excel(f, engine="calamine", header=2, skiprows=[3, 4]))
            bar.progress((i + 1) / len(names), text=f"Reading Shopee file {i+1}/{len(names)}…")
        bar.empty()
    return pd.concat(dfs, ignore_index=True)

def _find_col(df, *keyword_sets):
    for kws in keyword_sets:
        for c in df.columns:
            cl = str(c).strip().lower()
            if all(kw in cl for kw in kws): return c
    return None

# ==========================================
# 🎨 STREAMLIT INTERACTIVE UI
# ==========================================
st.title("🚀 Marketplace Price Automator")
st.write("Upload your structural files, map columns dynamically, and generate unified multi-sheet pricing exports.")

mode = st.selectbox(
    "Select Automation Mode",
    ["🔄 Run All Marketplace Channels", "🛍 Shopee Only", "🏪 Lazada Only", "👗 Zalora Only"],
    index=0
)

st.markdown("---")
col1, col2 = st.columns(2)

with col1:
    st.subheader("📋 Core Data Settings")
    tracker_file = st.file_uploader("1. Upload Master Tracker File (.csv, .xlsx)", type=["csv", "xlsx"])
    tracker_pim, tracker_rrp, tracker_md = None, None, None
    df_tracker = None
    
    if tracker_file:
        tracker_headers, df_tracker = get_clean_headers_and_df(tracker_file)
        if df_tracker is not None:
            st.success("💡 Cleaned tracking headers via specified index layout rules.")
            tracker_pim = st.selectbox("Map PIM ID Column", [""] + tracker_headers, key="t_pim")
            tracker_rrp = st.selectbox("Map Regular RRP Column", [""] + tracker_headers, key="t_rrp")
            tracker_md = st.selectbox("Map Special / Campaign / Markdown Price Column", [""] + tracker_headers, key="t_md")

    st.markdown("---")

    sku_file = st.file_uploader("2. Upload SKU Map File (.csv, .xlsx)", type=["csv", "xlsx"])
    sku_sku, sku_pim = None, None
    if sku_file:
        try:
            sku_headers = list(read_full_file_standard(sku_file).columns)
            sku_sku = st.selectbox("Map Seller SKU Column", [""] + sku_headers, key="s_sku")
            sku_pim = st.selectbox("Map PIM ID Column", [""] + sku_headers, key="s_pim")
        except Exception as e:
            st.error(f"Could not parse SKU Map layout: {e}")

with col2:
    st.subheader("🛍 Marketplace Templates")
    
    # --- SHOPEE SECTION ---
    shopee_file = None
    shopee_sku, shopee_promo, shopee_orig, shopee_start, shopee_end = None, None, None, None, None
    shopee_start_str, shopee_end_str = "", ""
    if mode in ["🛍 Shopee Only", "🔄 Run All Marketplace Channels"]:
        shopee_file = st.file_uploader("3. Upload Shopee Master Template (.zip)", type=["zip"])
        if shopee_file:
            try:
                df_shopee_preview = _read_shopee_zip(shopee_file)
                shopee_headers = list(df_shopee_preview.columns)
                
                shopee_sku = st.selectbox("Map Shopee SKU Column", [""] + shopee_headers, key="sh_sku")
                shopee_promo = st.selectbox("Map Shopee Promotion/Discount Price Column", [""] + shopee_headers, key="sh_promo")
                shopee_orig = st.selectbox("Map Shopee Original Price Column", [""] + shopee_headers, key="sh_orig")
                shopee_start = st.selectbox("Map Shopee Start Date Column (Optional)", [""] + shopee_headers, key="sh_start")
                shopee_end = st.selectbox("Map Shopee End Date Column (Optional)", [""] + shopee_headers, key="sh_end")
                
                if shopee_start or shopee_end:
                    st.caption("🗓️ **Set Shopee Campaign Run Windows**")
                    dates_col1, dates_col2 = st.columns(2)
                    with dates_col1:
                        sh_d1 = st.date_input("Shopee Start Date", datetime(2026, 7, 9))
                        sh_t1 = st.time_input("Shopee Start Time", time(23, 30, 0))
                        shopee_start_str = f"{sh_d1} {sh_t1.strftime('%H:%M:%S')}"
                    with dates_col2:
                        sh_d2 = st.date_input("Shopee End Date", datetime(2026, 8, 31))
                        sh_t2 = st.time_input("Shopee End Time", time(23, 59, 59))
                        shopee_end_str = f"{sh_d2} {sh_t2.strftime('%H:%M:%S')}"
            except Exception as e:
                st.error(f"Could not open or parse Shopee ZIP package stream: {e}")

    # --- LAZADA SECTION ---
    lazada_file = None
    lazada_sku, lazada_price = None, None
    if mode in ["🏪 Lazada Only", "🔄 Run All Marketplace Channels"]:
        if mode == "🔄 Run All Marketplace Channels": st.markdown("---")
        lazada_file = st.file_uploader("4. Upload Lazada Master Template (.csv, .xlsx)", type=["csv", "xlsx"])
        if lazada_file:
            try:
                lazada_headers = list(read_full_file_standard(lazada_file).columns)
                lazada_sku = st.selectbox("Map Lazada Seller SKU Column", [""] + lazada_headers, key="lz_sku")
                lazada_price = st.selectbox("Map Lazada Special/Campaign Price Column", [""] + lazada_headers, key="lz_price")
            except Exception as e:
                st.error(f"Could not parse Lazada layout: {e}")

    # --- ZALORA SECTION ---
    zalora_template_file = None
    zalora_upload_file = None
    zalora_sku, zalora_promo, zalora_orig, zalora_start, zalora_end = None, None, None, None, None
    zalora_start_str, zalora_end_str = "", ""
    if mode in ["👗 Zalora Only", "🔄 Run All Marketplace Channels"]:
        st.markdown("---")
        zalora_template_file = st.file_uploader("5a. Upload Zalora REFERENCE Template File (.xlsx)", type=["xlsx"], key="zal_temp")
        zalora_upload_file = st.file_uploader("5b. Upload Zalora TARGET Upload File (.xlsx)", type=["xlsx"], key="zal_upl")
        
        if zalora_upload_file:
            try:
                zalora_headers = list(read_full_file_standard(zalora_upload_file).columns)
                zalora_sku = st.selectbox("Map Zalora SKU Column", [""] + zalora_headers, key="zal_sku")
                zalora_promo = st.selectbox("Map Zalora Sale Price Column", [""] + zalora_headers, key="zal_promo")
                zalora_orig = st.selectbox("Map Zalora RRP Price Column", [""] + zalora_headers, key="zal_orig")
                zalora_start = st.selectbox("Map Zalora Sale Start Date Column", [""] + zalora_headers, key="zal_start")
                zalora_end = st.selectbox("Map Zalora Sale End Date Column", [""] + zalora_headers, key="zal_end")
                
                st.caption("🗓️ **Set Manual Zalora Campaign Windows**")
                zal_dates_col1, zal_dates_col2 = st.columns(2)
                with zal_dates_col1:
                    zal_d1 = st.date_input("Zalora Start Date", datetime(2026, 7, 9))
                    zal_t1 = st.time_input("Zalora Start Time", time(23, 30, 0))
                    zalora_start_str = f"{zal_d1} {zal_t1.strftime('%H:%M:%S')}"
                with zal_dates_col2:
                    zal_d2 = st.date_input("Zalora End Date", datetime(2026, 8, 31))
                    zal_t2 = st.time_input("Zalora End Time", time(23, 59, 59))
                    zalora_end_str = f"{zal_d2} {zal_t2.strftime('%H:%M:%S')}"
            except Exception as e:
                st.error(f"Could not parse Zalora layout: {e}")

st.markdown("---")

# ==========================================
# ⚙️ PROCESSING EXECUTION CORE
# ==========================================
if st.button("🚀 Run Automation Process", type="primary", use_container_width=True):
    error_found = False
    if not tracker_file or not tracker_pim or not tracker_rrp or not tracker_md:
        st.error("❌ Tracker layout mappings are invalid or unassigned."); error_found = True
    if not sku_file or not sku_sku or not sku_pim:
        st.error("❌ Universal reference SKU mapping parameters must be fully bound."); error_found = True
    if mode in ["🛍 Shopee Only", "🔄 Run All Marketplace Channels"] and (not shopee_file or not shopee_sku or not shopee_promo or not shopee_orig):
        st.error("❌ Shopee engine selected, but tracking dimensions are unassigned."); error_found = True
    if mode in ["🏪 Lazada Only", "🔄 Run All Marketplace Channels"] and (not lazada_file or not lazada_sku or not lazada_price):
        st.error("❌ Lazada engine selected, but data columns remain unassigned."); error_found = True
    if mode in ["👗 Zalora Only", "🔄 Run All Marketplace Channels"] and (not zalora_template_file or not zalora_upload_file or not zalora_sku or not zalora_promo or not zalora_orig or not zalora_start or not zalora_end):
        st.error("❌ Zalora engine selected, but columns remain unassigned."); error_found = True

    if not error_found:
        with st.spinner("Executing system pipeline mappings..."):
            try:
                df_sku = read_full_file_standard(sku_file)
                
                # 1. Parse Tracker Dictionary Mapping
                tracker_map = {}
                rrp_map = {}
                for _, row in df_tracker.iterrows():
                    pim = row[tracker_pim]
                    if pd.isna(pim) or str(pim).strip() == "": continue
                    
                    rrp = pd.to_numeric(row[tracker_rrp], errors='coerce') or 0
                    md = pd.to_numeric(row[tracker_md], errors='coerce') or 0
                    
                    new_price = round(md) if md != 0 and not pd.isna(md) else round(rrp)
                    tracker_map[pim] = new_price
                    rrp_map[pim] = round(rrp)

                # 2. Cross-reference Platform SKU Tables
                price_map = {}
                sku_to_pim = {}
                for _, row in df_sku.iterrows():
                    norm_sku = normalize_sku(row[sku_sku])
                    pim = row[sku_pim]
                    if pim in tracker_map:
                        price_map[norm_sku] = tracker_map[pim]
                        sku_to_pim[norm_sku] = pim

                output_buffer = io.BytesIO()
                sheets_written = 0
                
                with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
                    
                    # 3. Process Shopee Channel Data
                    if shopee_file and mode in ["🛍 Shopee Only", "🔄 Run All Marketplace Channels"]:
                        df_shopee = _read_shopee_zip(shopee_file)
                        mismatch_rows = []
                        upload_rows = []
                        df_shopee['QC Comment'] = ""
                        
                        for idx, row in df_shopee.iterrows():
                            sku = row[shopee_sku]
                            norm_sku = normalize_sku(sku)
                            existing_promo = row[shopee_promo]
                            orig_price = pd.to_numeric(row[shopee_orig], errors='coerce')
                            
                            new_price = price_map.get(norm_sku, existing_promo)
                            pim = sku_to_pim.get(norm_sku)
                            rrp = rrp_map.get(pim)
                            
                            comment = ""
                            is_mismatch = False
                            
                            if rrp is not None and orig_price != rrp:
                                comment = "RRP Mismatch"
                                is_mismatch = True
                                mismatch_rows.append({
                                    "Seller SKU": sku, "Marketplace Status": "Active", "Marketplace Message": "RRP Mismatch",
                                    "RRP": rrp, "Sale Amount": new_price, 
                                    "Sale Start Date(SGT)": shopee_start_str if shopee_start else '', 
                                    "Sale End Date(SGT)": shopee_end_str if shopee_end else ''
                                })
                            elif pd.isna(new_price) or new_price == "":
                                comment = "Discount Price is Blank"
                            elif rrp is not None and rrp == new_price:
                                comment = "Remove: RRP = Discount"
                            
                            df_shopee.at[idx, shopee_promo] = new_price
                            if shopee_start: df_shopee.at[idx, shopee_start] = shopee_start_str
                            if shopee_end: df_shopee.at[idx, shopee_end] = shopee_end_str
                            df_shopee.at[idx, 'QC Comment'] = comment
                            
                            if not is_mismatch and not (pd.isna(new_price) or new_price == "") and not (rrp is not None and rrp == new_price):
                                upload_row = row.copy()
                                upload_row[shopee_promo] = new_price
                                if shopee_start: upload_row[shopee_start] = shopee_start_str
                                if shopee_end: upload_row[shopee_end] = shopee_end_str
                                upload_rows.append(upload_row)

                        df_shopee.to_excel(writer, sheet_name="Shopee_Master_Updated", index=False)
                        df_mismatch = pd.DataFrame(mismatch_rows) if mismatch_rows else pd.DataFrame([{"Message": "No RRP Mismatches Found"}])
                        df_mismatch.to_excel(writer, sheet_name="Shopee_RRP_Mismatches", index=False)
                        
                        if upload_rows:
                            df_upload = pd.DataFrame(upload_rows).drop(columns=['QC Comment'], errors='ignore')
                            df_upload.to_excel(writer, sheet_name="Shopee_Upload", index=False)
                        sheets_written += 3

                    # 4. Process Lazada Channel Data
                    if lazada_file and mode in ["🏪 Lazada Only", "🔄 Run All Marketplace Channels"]:
                        df_lazada = read_full_file_standard(lazada_file)
                        for idx, row in df_lazada.iterrows():
                            norm_sku = normalize_sku(row[lazada_sku])
                            existing_price = row[lazada_price]
                            new_price = price_map.get(norm_sku, existing_price)
                            df_lazada.at[idx, lazada_price] = new_price
                        df_lazada.to_excel(writer, sheet_name="Lazada_Upload", index=False)
                        sheets_written += 1

                    # 5. Process Zalora Channel Data
                    if zalora_template_file and zalora_upload_file and mode in ["👗 Zalora Only", "🔄 Run All Marketplace Channels"]:
                        df_zalora_temp = read_full_file_standard(zalora_template_file)
                        df_zalora_upl = read_full_file_standard(zalora_upload_file)
                        
                        temp_sku_col = _find_col(df_zalora_temp, ("sku",), ("item",)) or df_zalora_temp.columns[0]
                        temp_rrp_col = _find_col(df_zalora_temp, ("rrp",), ("original",)) or df_zalora_temp.columns[1]
                        temp_srp_col = _find_col(df_zalora_temp, ("srp",), ("sale",), ("promo",)) or df_zalora_temp.columns[2]
                        
                        temp_data = {}
                        for _, r in df_zalora_temp.iterrows():
                            tsku = normalize_sku(r[temp_sku_col])
                            if tsku:
                                temp_data[tsku] = {
                                    "rrp": pd.to_numeric(r[temp_rrp_col], errors='coerce'),
                                    "srp": pd.to_numeric(r[temp_srp_col], errors='coerce')
                                }
                        
                        enriched_temp = df_zalora_temp.copy()
                        enriched_temp["ALU_NO"] = enriched_temp[temp_sku_col].apply(clean_id_str)
                        enriched_temp["Current RRP"] = ""
                        enriched_temp["RRP Check"] = ""
                        enriched_temp["Current SRP"] = ""
                        enriched_temp["SRP Check"] = ""
                        enriched_temp["Comments"] = ""
                        
                        final_upload_rows = []
                        invalid_skus = []  # Tracking error logs to pinpoint validation dropouts
                        
                        for idx, row in df_zalora_upl.iterrows():
                            sku = row[zalora_sku]
                            norm_sku = normalize_sku(sku)
                            pim = sku_to_pim.get(norm_sku)
                            t_info = temp_data.get(norm_sku)
                            
                            # Rule 4 Diagnostics check
                            if not pim or norm_sku not in price_map or not t_info or pd.isna(pim):
                                invalid_skus.append({
                                    "Row Index": idx + 4,
                                    "Input SKU": sku,
                                    "Normalized SKU": norm_sku,
                                    "Reason for Removal": "Missing from SKU_Map or Master Tracker mapping framework (#N/A)"
                                })
                                continue
                                
                            tracker_rrp_val = rrp_map.get(pim, 0)
                            tracker_srp_val = price_map.get(norm_sku, 0)
                            upload_rrp = pd.to_numeric(row[zalora_orig], errors='coerce') or 0
                            upload_srp = pd.to_numeric(row[zalora_promo], errors='coerce') or 0
                            
                            rrp_check = (t_info["rrp"] == upload_rrp)
                            srp_check = (t_info["srp"] == upload_srp)
                            
                            if tracker_srp_val == 0:
                                row[zalora_promo] = ""
                                row[zalora_start] = ""
                                row[zalora_end] = ""
                                final_upload_rows.append(row)
                                continue
                                
                            if rrp_check and srp_check:
                                current_end_dt = row[zalora_end]
                                try:
                                    parsed_end = pd.to_datetime(current_end_dt)
                                    if parsed_end < datetime.now() + timedelta(days=30):
                                        row[zalora_start] = zalora_start_str
                                        row[zalora_end] = zalora_end_str
                                except:
                                    row[zalora_start] = zalora_start_str
                                    row[zalora_end] = zalora_end_str
                            else:
                                row[zalora_orig] = tracker_rrp_val
                                row[zalora_promo] = tracker_srp_val
                                row[zalora_start] = zalora_start_str
                                row[zalora_end] = zalora_end_str
                                
                            final_upload_rows.append(row)
                        
                        for t_idx, t_row in enriched_temp.iterrows():
                            t_sku_val = normalize_sku(t_row[temp_sku_col])
                            t_pim = sku_to_pim.get(t_sku_val)
                            if t_pim:
                                enriched_temp.at[t_idx, "Current RRP"] = rrp_map.get(t_pim, "")
                                enriched_temp.at[t_idx, "Current SRP"] = price_map.get(t_sku_val, "")
                        
                        enriched_temp.to_excel(writer, sheet_name="Zalora_Template_Enriched", index=False)
                        sheets_written += 1
                        
                        if final_upload_rows:
                            df_zalora_final = pd.DataFrame(final_upload_rows)
                            df_zalora_final.to_excel(writer, sheet_name="Zalora_Final_Upload", index=False)
                            sheets_written += 1
                        
                        # Write the error log diagnostic tab
                        df_errors = pd.DataFrame(invalid_skus) if invalid_skus else pd.DataFrame([{"Message": "No records were dropped. All items mapped successfully!"}])
                        df_errors.to_excel(writer, sheet_name="Zalora_Errors", index=False)
                        sheets_written += 1

                    if sheets_written == 0:
                        pd.DataFrame([{
                            "Execution Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                            "Status": "Skipped",
                            "Diagnostic Note": "No marketplace template files were uploaded matching the chosen processing mode configuration."
                        }]).to_excel(writer, sheet_name="Execution_Summary", index=False)

                output_buffer.seek(0)
                st.success("🎉 Automation executed successfully!")
                st.download_button(
                    label="📥 Download Consolidated Marketplace Workbook",
                    data=output_buffer,
                    file_name="Consolidated_Marketplace_Pricing.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            except Exception as e:
                st.error(f"A systematic error occurred during calculations: {e}")
