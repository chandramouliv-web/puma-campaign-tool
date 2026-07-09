import streamlit as st
import pandas as pd
import io
import re
import zipfile
from datetime import datetime, time

# Set up page configuration
st.set_page_config(page_title="Marketplace Price Automator", page_icon="🚀", layout="wide")

# =================================================================
# ⚙️ ULTRA HIGH-PERFORMANCE LOW-LATENCY ENGINES
# =================================================================

def normalize_sku(sku):
    if pd.isna(sku) or not sku:
        return ""
    return str(sku).strip().replace('-', '_').lower()

def clean_id_str(val):
    if pd.isna(val):
        return ""
    if isinstance(val, float):
        return str(int(val)) if val.is_integer() else str(val)
    s = str(val).strip()
    if re.match(r"^-?\d+\.0+$", s):
        s = s.split(".")[0]
    return s if s != "nan" else ""

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
            if any(kw in cl for kw in kws): return c
    return None

def is_last_day_of_month(date_val):
    try:
        if pd.isna(date_val) or str(date_val).strip() == "":
            return False
        dt = pd.to_datetime(date_val)
        return dt.day == dt.days_in_month
    except:
        return False

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
            st.info("💡 Auto-suggesting mappings from your SKU Map columns.")
            
            def_sku_idx = sku_headers.index("EAN") if "EAN" in sku_headers else 0
            def_pim_idx = sku_headers.index("Color_No") if "Color_No" in sku_headers else 0
            
            sku_sku = st.selectbox("Map Seller SKU / EAN Column", [""] + sku_headers, index=def_sku_idx + 1, key="s_sku")
            sku_pim = st.selectbox("Map PIM ID / Color_No Column", [""] + sku_headers, index=def_pim_idx + 1, key="s_pim")
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
    zalora_file = None
    zalora_sku, zalora_promo, zalora_orig, zalora_start, zalora_end = None, None, None, None, None
    zalora_start_str, zalora_end_str = "", ""
    if mode in ["👗 Zalora Only", "🔄 Run All Marketplace Channels"]:
        st.markdown("---")
        zalora_file = st.file_uploader("5. Upload Master Zalora Price Template (.xlsx)", type=["xlsx"], key="zal_master")
        
        if zalora_file:
            try:
                zalora_headers = list(read_full_file_standard(zalora_file).columns)
                zalora_sku = st.selectbox("Map Zalora SKU Column (e.g. SellerSku)", [""] + zalora_headers, key="zal_sku")
                zalora_promo = st.selectbox("Map Zalora Sale Price Column (e.g. SalePrice)", [""] + zalora_headers, key="zal_promo")
                zalora_orig = st.selectbox("Map Zalora RRP Price Column (e.g. Price)", [""] + zalora_headers, key="zal_orig")
                zalora_start = st.selectbox("Map Zalora Sale Start Date Column (e.g. SaleStartDate)", [""] + zalora_headers, key="zal_start")
                zalora_end = st.selectbox("Map Zalora Sale End Date Column (e.g. SaleEndDate)", [""] + zalora_headers, key="zal_end")
                
                st.caption("🗓️ **Set Manual Campaign Run Windows**")
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
    if mode in ["👗 Zalora Only", "🔄 Run All Marketplace Channels"] and (not zalora_file or not zalora_sku or not zalora_promo or not zalora_orig or not zalora_start or not zalora_end):
        st.error("❌ Zalora engine selected, but columns remain unassigned."); error_found = True

    if not error_found:
        with st.spinner("Executing real-time dictionary hashing mappings..."):
            try:
                # ⚡ INSTANT IN-MEMORY CONVERSION: Ingest SKU file as pure arrays
                df_sku = read_full_file_standard(sku_file)
                sku_raw_arr = df_sku[sku_sku].astype(str).values
                pim_raw_arr = df_sku[sku_pim].astype(str).values
                
                # Pre-build double-sided mapping indices for O(1) cascading fallbacks
                ean_to_pim = {}
                alu_to_pim = {}
                
                for idx_s in range(len(sku_raw_arr)):
                    raw_s = sku_raw_arr[idx_s].strip()
                    norm_s = normalize_sku(raw_s)
                    pim_clean = clean_id_str(pim_raw_arr[idx_s])
                    
                    if norm_s: ean_to_pim[norm_s] = pim_clean
                    if raw_s: alu_to_pim[raw_s] = pim_clean

                # ⚡ INSTANT IN-MEMORY TRACKER CONVERSION
                tracker_map = {}
                rrp_map = {}
                
                pim_tracker_raw = df_tracker[tracker_pim].values
                rrp_tracker_raw = pd.to_numeric(df_tracker[tracker_rrp], errors='coerce').fillna(0).values
                md_tracker_raw = pd.to_numeric(df_tracker[tracker_md], errors='coerce').fillna(0).values
                
                for t_idx in range(len(df_tracker)):
                    pim_val = clean_id_str(pim_tracker_raw[t_idx])
                    if not pim_val: continue
                    tracker_map[pim_val] = round(md_tracker_raw[t_idx])
                    rrp_map[pim_val] = round(rrp_tracker_raw[t_idx])

                output_buffer = io.BytesIO()
                with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
                    
                    pd.DataFrame([{
                        "Run Date": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "Automation Mode Selection": mode,
                        "Engine State": "Executed Instantly via Primitive List Dictionaries"
                    }]).to_excel(writer, sheet_name="Dashboard_Summary", index=False)
                    
                    # 3. Process Shopee Channel Data
                    if shopee_file and mode in ["🛍 Shopee Only", "🔄 Run All Marketplace Channels"]:
                        df_shopee = _read_shopee_zip(shopee_file)
                        mismatch_rows = []
                        upload_rows = []
                        
                        shopee_skus = df_shopee[shopee_sku].values
                        shopee_promos = df_shopee[shopee_promo].values
                        shopee_origs = pd.to_numeric(df_shopee[shopee_orig], errors='coerce').fillna(0).values
                        shopee_records = df_shopee.to_dict('records')
                        
                        for idx, row in enumerate(shopee_records):
                            sku = shopee_skus[idx]
                            norm_sku = normalize_sku(sku)
                            existing_promo = shopee_promos[idx]
                            orig_price = shopee_origs[idx]
                            
                            pim = ean_to_pim.get(norm_sku)
                            new_price = tracker_map.get(pim, existing_promo)
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
                            
                            row[shopee_promo] = new_price
                            if shopee_start: row[shopee_start] = shopee_start_str
                            if shopee_end: row[shopee_end] = shopee_end_str
                            row['QC Comment'] = comment
                            
                            if not is_mismatch and not (pd.isna(new_price) or new_price == "") and not (rrp is not None and rrp == new_price):
                                upload_rows.append(row.copy())

                        pd.DataFrame(shopee_records).to_excel(writer, sheet_name="Shopee_Master_Updated", index=False)
                        df_mismatch = pd.DataFrame(mismatch_rows) if mismatch_rows else pd.DataFrame([{"Message": "No RRP Mismatches Found"}])
                        df_mismatch.to_excel(writer, sheet_name="Shopee_RRP_Mismatches", index=False)
                        if upload_rows:
                            pd.DataFrame(upload_rows).drop(columns=['QC Comment'], errors='ignore').to_excel(writer, sheet_name="Shopee_Upload", index=False)

                    # 4. Process Lazada Channel Data
                    if lazada_file and mode in ["🏪 Lazada Only", "🔄 Run All Marketplace Channels"]:
                        df_lazada = read_full_file_standard(lazada_file)
                        lazada_skus = df_lazada[lazada_sku].values
                        lazada_prices = df_lazada[lazada_price].values
                        
                        final_lazada_prices = [tracker_map.get(ean_to_pim.get(normalize_sku(lazada_skus[i])), lazada_prices[i]) for i in range(len(df_lazada))]
                        df_lazada[lazada_price] = final_lazada_prices
                        df_lazada.to_excel(writer, sheet_name="Lazada_Upload", index=False)

                    # 5. Process Zalora Channel Data (Enforced Instant Lookups Architecture)
                    if zalora_file and mode in ["👗 Zalora Only", "🔄 Run All Marketplace Channels"]:
                        df_zalora_raw = read_full_file_standard(zalora_file)
                        df_zalora_raw.to_excel(writer, sheet_name="Direct Download From Zalora", index=False)
                        
                        temp_sku_col = _find_col(df_zalora_raw, ["sku", "item", "sellersku"]) or df_zalora_raw.columns[0]
                        temp_rrp_col = _find_col(df_zalora_raw, ["price", "rrp", "original"]) or df_zalora_raw.columns[1]
                        temp_srp_col = _find_col(df_zalora_raw, ["saleprice", "srp", "sale"]) or df_zalora_raw.columns[2]
                        
                        working_flow_rows = []
                        final_to_upload_rows = []
                        
                        # Ingest metrics into clean list dictionaries
                        zalora_records = df_zalora_raw.to_dict('records')
                        
                        for row_dict in zalora_records:
                            sku = str(row_dict[temp_sku_col]).strip()
                            norm_sku = normalize_sku(sku)
                            
                            # CASCADING CATCH: Check direct structural ALU first, fallback instantly to normalized EAN hash map
                            pim = alu_to_pim.get(sku) or ean_to_pim.get(norm_sku)
                            
                            # RULE 1: Item Not Found in Tracker
                            if not pim or pim not in rrp_map:
                                row_dict.update({
                                    "ALU_NO/Color_No": "", "RRP/PH EC RRP": "", "RRP check (I=D)": "False",
                                    "SRP/PH MD Price": "", "SRP check (K=E)": "False",
                                    "Comments": "Ignore – Item Not Found in Tracker."
                                })
                                working_flow_rows.append(row_dict)
                                continue
                                
                            tracker_rrp_val = rrp_map[pim]
                            tracker_srp_val = tracker_map[pim]  # Directly assigned from tracker array matrix
                            
                            current_rrp = pd.to_numeric(row_dict[temp_rrp_col], errors='coerce') or 0
                            current_srp = pd.to_numeric(row_dict[temp_srp_col], errors='coerce') or 0
                            
                            rrp_match = (tracker_rrp_val == current_rrp)
                            srp_match = (tracker_srp_val == current_srp)
                            
                            row_dict.update({
                                "ALU_NO/Color_No": str(pim), 
                                "RRP/PH EC RRP": str(tracker_rrp_val),
                                "RRP check (I=D)": str(rrp_match), 
                                "SRP/PH MD Price": str(tracker_srp_val),  
                                "SRP check (K=E)": str(srp_match)
                            })
                            
                            has_no_dates = pd.isna(row_dict[zalora_start]) or str(row_dict[zalora_start]).strip() == ""
                            is_last_day = is_last_day_of_month(row_dict[zalora_end])
                            
                            # RULE 2: No Changes Required
                            if rrp_match and srp_match and has_no_dates:
                                row_dict["Comments"] = "All Good – RRP and SRP Match the Tracker. No Update Required."
                                working_flow_rows.append(row_dict)
                                continue
                                
                            # RULE 3: Month-End Sale Window Exception Check
                            if rrp_match and srp_match and is_last_day:
                                row_dict["Comments"] = "All Good – Sale Ends at Month End. No Update Required."
                                working_flow_rows.append(row_dict)
                                continue
                                
                            # RULE 4: Sale Price Mismatch Pipeline
                            if rrp_match and not srp_match:
                                if tracker_srp_val == 0:
                                    row_dict.update({temp_srp_col: "", zalora_start: "", zalora_end: "", "Comments": "Sale Price Updated."})
                                else:
                                    row_dict.update({
                                        temp_srp_col: str(tracker_srp_val),
                                        zalora_start: str(zalora_start_str), zalora_end: str(zalora_end_str),
                                        "Comments": "Sale Price Updated."
                                    })
                                working_flow_rows.append(row_dict)
                                final_to_upload_rows.append(row_dict.copy())
                                continue
                                
                            # RULE 5: Base RRP Overwrite variance check conditions
                            if not rrp_match:
                                row_dict[temp_rrp_col] = str(tracker_rrp_val)
                                if tracker_srp_val == 0:
                                    row_dict.update({temp_srp_col: "", zalora_start: "", zalora_end: "", "Comments": "RRP and Sale Price Updated."})
                                else:
                                    row_dict.update({
                                        temp_srp_col: str(tracker_srp_val),
                                        zalora_start: str(zalora_start_str), zalora_end: str(zalora_end_str),
                                        "Comments": "RRP and Sale Price Updated."
                                    })
                                working_flow_rows.append(row_dict)
                                final_to_upload_rows.append(row_dict.copy())
                                continue

                            working_flow_rows.append(row_dict)

                        # Write rows out directly to excel sheets
                        pd.DataFrame(working_flow_rows).to_excel(writer, sheet_name="Working Flow", index=False)
                        
                        if final_to_upload_rows:
                            df_to_upload = pd.DataFrame(final_to_upload_rows)
                            cols_to_drop = ["ALU_NO/Color_No", "RRP/PH EC RRP", "RRP check (I=D)", "SRP/PH MD Price", "SRP check (K=E)", "Comments"]
                            df_to_upload.drop(columns=cols_to_drop, errors='ignore').to_excel(writer, sheet_name="To Upload", index=False)
                        else:
                            pd.DataFrame([{"Message": "No items required updates; all records skipped from upload sheet."}]).to_excel(writer, sheet_name="To Upload", index=False)

                output_buffer.seek(0)
                st.success("🚀 Execution Instantaneous! Array scans completed smoothly.")
                st.download_button(
                    label="📥 Download Consolidated Marketplace Workbook",
                    data=output_buffer,
                    file_name="Consolidated_Marketplace_Pricing.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            except Exception as e:
                st.error(f"A systematic error occurred during calculations: {e}")
