import streamlit as st
import pandas as pd
import io
import re
import zipfile
import math
from datetime import datetime, time
import openpyxl
from openpyxl.styles import PatternFill, Font, Border, Alignment
from openpyxl.utils import get_column_letter

# Set up page configuration
st.set_page_config(page_title="Marketplace Price Automator", page_icon="🚀", layout="wide")

# =================================================================
# ⚙️ ULTRA HIGH-PERFORMANCE DATA ENGINES & PLUGINS (CACHED)
# =================================================================

@st.cache_data(show_spinner="Analyzing file structure...")
def get_cached_sheet_names(file_bytes, file_name):
    if file_name.endswith(('.xls', '.xlsx')):
        try:
            xl = pd.ExcelFile(io.BytesIO(file_bytes), engine='calamine')
            return xl.sheet_names
        except:
            try:
                xl = pd.ExcelFile(io.BytesIO(file_bytes))
                return xl.sheet_names
            except:
                return ["Sheet1"]
    return ["Default"]

@st.cache_data(show_spinner="Parsing tracking headers...")
def get_cached_clean_headers_and_df(file_bytes, file_name, target_sheet=None):
    try:
        if file_name.endswith('.csv'):
            full_df = pd.read_csv(io.BytesIO(file_bytes), header=None)
        else:
            try:
                full_df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=target_sheet, header=None, engine='calamine')
            except:
                full_df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=target_sheet, header=None)

        if full_df.empty or len(full_df) < 3:
            return [], full_df

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
        data_df = full_df.iloc[3:].reset_index(drop=True)
        return clean_headers, data_df
    except Exception as e:
        return [], None

@st.cache_data(show_spinner="Ingesting flat files...")
def read_cached_file_standard(file_bytes, file_name):
    if file_name.endswith('.csv'):
        return pd.read_csv(io.BytesIO(file_bytes))
    try:
        return pd.read_excel(io.BytesIO(file_bytes), engine='calamine')
    except:
        return pd.read_excel(io.BytesIO(file_bytes))

@st.cache_data(show_spinner="Processing master template...")
def _read_shopee_stream_flexible(file_bytes, file_name):
    if file_name.endswith('.zip'):
        dfs = []
        with zipfile.ZipFile(io.BytesIO(file_bytes)) as zf:
            names = sorted(n for n in zf.namelist() if n.endswith(('.xlsx', '.xls', '.csv')))
            if not names:
                raise ValueError("No active dataset files located inside the ZIP bundle.")
            for name in names:
                with zf.open(name) as f:
                    if name.endswith('.csv'):
                        df = pd.read_csv(f, header=None)
                    else:
                        try:
                            df = pd.read_excel(io.BytesIO(f.read()), header=None, engine='calamine')
                        except:
                            df = pd.read_excel(io.BytesIO(f.read()), header=None)
                    
                    df_cleaned = auto_detect_shopee_header(df)
                    if isinstance(df_cleaned, pd.DataFrame) and not df_cleaned.empty:
                        dfs.append(df_cleaned)
        if not dfs:
            return pd.DataFrame()
        consolidated_df = pd.concat(dfs, ignore_index=True)
        consolidated_df.drop_duplicates(inplace=True)
        return consolidated_df
        
    if file_name.endswith('.csv'):
        df = pd.read_csv(io.BytesIO(file_bytes), header=None)
    else:
        try:
            df = pd.read_excel(io.BytesIO(file_bytes), header=None, engine='calamine')
        except:
            df = pd.read_excel(io.BytesIO(file_bytes), header=None)
    return auto_detect_shopee_header(df)

def auto_detect_shopee_header(df):
    for idx in range(min(15, len(df))):
        row_str = df.iloc[idx].fillna("").astype(str).str.lower().str.strip().values
        if any("product id" in r or "variation id" in r or "sku ref" in r or "parent sku" in r for r in row_str):
            headers = df.iloc[idx].fillna("").astype(str).str.strip().values
            clean_headers = [h if (h and h.lower() != "nan") else f"Blank_Col_{i}" for i, h in enumerate(headers)]
            new_df = df.iloc[idx+1:].copy()
            new_df.columns = clean_headers
            return new_df.reset_index(drop=True)
            
    if len(df) > 0:
        df.columns = [str(c).strip() for c in df.iloc[0].fillna("").values]
        return df.iloc[1:].reset_index(drop=True)
    return df

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

def _apply_rounding_strategy(val, strategy_mode):
    try:
        f_val = float(val)
        if strategy_mode == "Round Up":
            return math.ceil(f_val)
        elif strategy_mode == "Round":
            return round(f_val)
        else:
            return round(f_val, 2)
    except:
        return 0

def _to_numeric_safe(val, strategy_mode):
    if pd.isna(val):
        return 0
    try:
        num = pd.to_numeric(val, errors='coerce')
        if pd.isna(num):
            return 0
        return _apply_rounding_strategy(num, strategy_mode)
    except:
        return 0

def _find_col(df, keyword_sets):
    columns_lower = {str(c).strip().lower(): c for c in df.columns}
    for kw in keyword_sets:
        kw_lower = kw.lower()
        for c_lower, orig_col in columns_lower.items():
            if kw_lower in c_lower:
                return orig_col
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
st.write("Upload your structural files, map columns dynamically, and generate unified marketplace pricing exports.")

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
    tracker_sheet = None
    tracker_pim, tracker_rrp, tracker_md = None, None, None
    df_tracker = None
    
    if tracker_file:
        t_bytes = tracker_file.getvalue()
        if tracker_file.name.endswith(('.xlsx', '.xls')):
            t_sheets = get_cached_sheet_names(t_bytes, tracker_file.name)
            tracker_sheet = st.selectbox("Select Target Tracker Worksheet", t_sheets, key="t_sheet_selector")
        
        if "tracker_round_prefs" not in st.session_state:
            st.session_state.tracker_round_prefs = {}
            
        current_sheet_key = f"{tracker_file.name}_{tracker_sheet or 'Default'}"
        saved_pref_index = 0
        if current_sheet_key in st.session_state.tracker_round_prefs:
            saved_pref = st.session_state.tracker_round_prefs[current_sheet_key]
            saved_pref_index = ["Follow Exact Tracker Value", "Round", "Round Up"].index(saved_pref)

        sheet_rounding_strategy = st.selectbox(
            f"Select Price Rounding Method for Sheet [{tracker_sheet or 'Default'}]",
            ["Follow Exact Tracker Value", "Round", "Round Up"],
            index=saved_pref_index,
            key="sheet_rounding_dropdown"
        )
        st.session_state.tracker_round_prefs[current_sheet_key] = sheet_rounding_strategy
            
        tracker_headers, df_tracker = get_cached_clean_headers_and_df(t_bytes, tracker_file.name, target_sheet=tracker_sheet)
        if df_tracker is not None:
            st.success(f"💡 Layout parsed with [{sheet_rounding_strategy}] rounding active for this sheet.")
            tracker_pim = st.selectbox("Map PIM ID Column", [""] + tracker_headers, key="t_pim")
            tracker_rrp = st.selectbox("Map Regular RRP Column", [""] + tracker_headers, key="t_rrp")
            tracker_md = st.selectbox("Map Special / Campaign / Markdown Price Column", [""] + tracker_headers, key="t_md")

    st.markdown("---")

    sku_file = st.file_uploader("2. Upload SKU Map File (.csv, .xlsx)", type=["csv", "xlsx"])
    sku_sku, sku_pim = None, None
    if sku_file:
        s_bytes = sku_file.getvalue()
        try:
            sku_headers = list(read_cached_file_standard(s_bytes, sku_file.name).columns)
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
    shopee_sku, shopee_parent, shopee_orig = None, None, None
    if mode in ["🛍 Shopee Only", "🔄 Run All Marketplace Channels"]:
        shopee_file = st.file_uploader("3. Upload Shopee Master Template (.zip, .xlsx, .csv)", type=["zip", "xlsx", "csv"])
        if shopee_file:
            sh_bytes = shopee_file.getvalue()
            try:
                df_shopee_preview = _read_shopee_stream_flexible(sh_bytes, shopee_file.name)
                shopee_headers = list(df_shopee_preview.columns)
                
                shopee_sku = st.selectbox("Map Shopee SKU Column (e.g. SKU)", [""] + shopee_headers, key="sh_sku")
                shopee_parent = st.selectbox("Map Shopee Parent SKU Column", [""] + shopee_headers, key="sh_parent")
                shopee_orig = st.selectbox("Map Shopee Original Price Column", [""] + shopee_headers, key="sh_orig")
            except Exception as e:
                st.error(f"Could not open or parse Shopee raw data stream: {e}")

    # --- LAZADA SECTION ---
    lazada_file = None
    lazada_sku, lazada_price = None, None
    if mode in ["🏪 Lazada Only", "🔄 Run All Marketplace Channels"]:
        if mode == "🔄 Run All Marketplace Channels": st.markdown("---")
        lazada_file = st.file_uploader("4. Upload Lazada Master Template (.csv, .xlsx)", type=["csv", "xlsx"])
        if lazada_file:
            lz_bytes = lazada_file.getvalue()
            try:
                lazada_headers = list(read_cached_file_standard(lz_bytes, lazada_file.name).columns)
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
            zal_bytes = zalora_file.getvalue()
            try:
                zalora_headers = list(read_cached_file_standard(zal_bytes, zalora_file.name).columns)
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
    if mode in ["🛍 Shopee Only", "🔄 Run All Marketplace Channels"] and (not shopee_file or not shopee_sku or not shopee_orig):
        st.error("❌ Shopee engine selected, but tracking dimensions are unassigned."); error_found = True
    if mode in ["🏪 Lazada Only", "🔄 Run All Marketplace Channels"] and (not lazada_file or not lazada_sku or not lazada_price):
        st.error("❌ Lazada engine selected, but data columns remain unassigned."); error_found = True
    if mode in ["👗 Zalora Only", "🔄 Run All Marketplace Channels"] and (not zalora_file or not zalora_sku or not zalora_promo or not zalora_orig or not zalora_start or not zalora_end):
        st.error("❌ Zalora engine selected, but columns remain unassigned."); error_found = True

    if not error_found:
        s_bytes = sku_file.getvalue()
        t_bytes = tracker_file.getvalue()
        
        with st.spinner("Executing system pipeline lookups..."):
            try:
                # 1. High Performance In-Memory SKU Mapping Ingestion
                df_sku = read_cached_file_standard(s_bytes, sku_file.name)
                sku_raw_arr = df_sku[sku_sku].astype(str).values
                pim_raw_arr = df_sku[sku_pim].astype(str).values
                
                ean_to_pim = {}
                alu_to_pim = {}
                
                for idx_s in range(len(sku_raw_arr)):
                    raw_s = sku_raw_arr[idx_s].strip()
                    norm_s = normalize_sku(raw_s)
                    pim_clean = clean_id_str(pim_raw_arr[idx_s])
                    
                    if norm_s: ean_to_pim[norm_s] = pim_clean
                    if raw_s: alu_to_pim[raw_s] = pim_clean

                # 2. Parse Tracker Data once into memory hash maps
                tracker_map = {}
                rrp_map = {}
                
                pim_tracker_raw = df_tracker[tracker_pim].values
                rrp_tracker_raw = pd.to_numeric(df_tracker[tracker_rrp], errors='coerce').fillna(0).values
                md_tracker_raw = pd.to_numeric(df_tracker[tracker_md], errors='coerce').fillna(0).values
                
                for t_idx in range(len(df_tracker)):
                    pim_val = clean_id_str(pim_tracker_raw[t_idx])
                    if not pim_val: continue
                    tracker_map[pim_val] = _apply_rounding_strategy(md_tracker_raw[t_idx], sheet_rounding_strategy)
                    rrp_map[pim_val] = _apply_rounding_strategy(rrp_tracker_raw[t_idx], sheet_rounding_strategy)

                output_buffer = io.BytesIO()
                with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
                    
                    # 3. Process Shopee Channel Data
                    if shopee_file and mode in ["🛍 Shopee Only", "🔄 Run All Marketplace Channels"]:
                        sh_bytes = shopee_file.getvalue()
                        df_shopee_raw = _read_shopee_stream_flexible(sh_bytes, shopee_file.name)
                        if not df_shopee_raw.empty:
                            df_shopee_raw.to_excel(writer, sheet_name="Consolidated File", index=False)
                            
                            shopee_promo = _find_col(df_shopee_raw, ["discount price", "promo", "campaign price"]) or "Discount price"
                            
                            shopee_working_flow = []
                            shopee_mismatches = []
                            shopee_final_uploads = []
                            
                            shopee_records = df_shopee_raw.to_dict('records')
                            for row in shopee_records:
                                raw_sku = str(row.get(shopee_sku, '')).strip()
                                raw_parent = str(row.get(shopee_parent, '')).strip()
                                
                                if not raw_sku or raw_sku.lower() == 'nan' or raw_sku == '':
                                    raw_sku = raw_parent
                                    row[shopee_sku] = raw_sku
                                    
                                if (not raw_sku or raw_sku.lower() == 'nan' or raw_sku == '') and (not raw_parent or raw_parent.lower() == 'nan' or raw_parent == ''):
                                    row.update({
                                        "ALU_NO": "", "RRP": "", "RRP Check": "False", "SRP": "",
                                        "Comments": "Missing SKU and Parent SKU."
                                    })
                                    shopee_working_flow.append(row)
                                    continue
                                    
                                norm_sku = normalize_sku(raw_sku)
                                pim = ean_to_pim.get(norm_sku) or alu_to_pim.get(raw_sku)
                                
                                if not pim or pim not in rrp_map:
                                    row.update({
                                        "ALU_NO": "", "RRP": "", "RRP Check": "False", "SRP": "",
                                        "Comments": "Ignore – Item Not Found in Tracker."
                                    })
                                    shopee_working_flow.append(row)
                                    continue
                                    
                                tracker_rrp_val = rrp_map[pim]
                                tracker_srp_val = tracker_map[pim]
                                current_price = _to_numeric_safe(row.get(shopee_orig, 0), sheet_rounding_strategy)
                                
                                rrp_match = (tracker_rrp_val == current_price)
                                
                                row.update({
                                    "ALU_NO": str(pim), "RRP": str(tracker_rrp_val),
                                    "RRP Check": str(rrp_match), "SRP": str(tracker_srp_val)
                                })
                                
                                if not rrp_match:
                                    row['Comments'] = "RRP Mismatch"
                                    shopee_working_flow.append(row)
                                    shopee_mismatches.append({
                                        "Seller SKU": raw_sku, "Marketplace Status": "", "Marketplace Message": "", "RRP": tracker_rrp_val
                                    })
                                    continue
                                
                                if tracker_srp_val == 0:
                                    row['Comments'] = "ignore - SRP is Zero"
                                    shopee_working_flow.append(row)
                                    continue
                                
                                row['Comments'] = "RRP is true and SRP is not equal to RRP - To be in Upload File"
                                row[shopee_promo] = tracker_srp_val
                                
                                shopee_working_flow.append(row)
                                shopee_final_uploads.append(row.copy())
                                
                            pd.DataFrame(shopee_working_flow).to_excel(writer, sheet_name="Working File", index=False)
                            
                            df_sh_mismatch = pd.DataFrame(shopee_mismatches) if shopee_mismatches else pd.DataFrame([{"Message": "No mismatches detected"}])
                            df_sh_mismatch.to_excel(writer, sheet_name="RRP Mismatches", index=False)
                            
                            if shopee_final_uploads:
                                df_sh_upload = pd.DataFrame(shopee_final_uploads)
                                drop_cols = ['ALU_NO', 'RRP', 'RRP Check', 'SRP', 'Comments']
                                df_sh_upload.drop(columns=drop_cols, errors='ignore').to_excel(writer, sheet_name="To Upload", index=False)
                            else:
                                pd.DataFrame([{"Message": "All entries skipped based on filtering rules."}]).to_excel(writer, sheet_name="To Upload", index=False)

                    # 4. Process Lazada Channel Data
                    if lazada_file and mode in ["🏪 Lazada Only", "🔄 Run All Marketplace Channels"]:
                        lz_bytes = lazada_file.getvalue()
                        df_lazada = read_cached_file_standard(lz_bytes, lazada_file.name)
                        lazada_skus = df_lazada[lazada_sku].values
                        lazada_prices = df_lazada[lazada_price].values
                        
                        final_lazada_prices = []
                        for i in range(len(df_lazada)):
                            norm_lz_sku = normalize_sku(lazada_skus[i])
                            pim = ean_to_pim.get(norm_lz_sku) or alu_to_pim.get(str(lazada_skus[i]).strip())
                            
                            if pim and pim in rrp_map:
                                tracker_rrp_val = rrp_map[pim]
                                tracker_srp_val = tracker_map[pim]
                                computed_campaign_price = _apply_rounding_strategy(tracker_rrp_val, sheet_rounding_strategy) if tracker_srp_val == 0 else _apply_rounding_strategy(tracker_srp_val, sheet_rounding_strategy)
                                final_lazada_prices.append(computed_campaign_price)
                            else:
                                final_lazada_prices.append(lazada_prices[i])
                                
                        df_lazada[lazada_price] = final_lazada_prices
                        df_lazada.to_excel(writer, sheet_name="Lazada_Upload", index=False)

                    # 5. Process Zalora Channel Data
                    if zalora_file and mode in ["👗 Zalora Only", "🔄 Run All Marketplace Channels"]:
                        zal_bytes = zalora_file.getvalue()
                        df_zalora_raw = read_cached_file_standard(zal_bytes, zalora_file.name)
                        df_zalora_raw.to_excel(writer, sheet_name="Direct Download From Zalora", index=False)
                        
                        temp_sku_col = _find_col(df_zalora_raw, ["sku", "item", "sellersku"]) or df_zalora_raw.columns[0]
                        temp_rrp_col = _find_col(df_zalora_raw, ["price", "rrp", "original"]) or df_zalora_raw.columns[1]
                        temp_srp_col = _find_col(df_zalora_raw, ["saleprice", "srp", "sale"]) or df_zalora_raw.columns[2]
                        
                        working_flow_rows = []
                        final_to_upload_rows = []
                        
                        zalora_records = df_zalora_raw.to_dict('records')
                        user_selected_end_dt = pd.to_datetime(zalora_end_str)
                        
                        for row_dict in zalora_records:
                            sku = str(row_dict[temp_sku_col]).strip()
                            norm_sku = normalize_sku(sku)
                            pim = alu_to_pim.get(sku) or ean_to_pim.get(norm_sku)
                            
                            if not pim or pim not in rrp_map:
                                row_dict.update({
                                    "ALU_NO/Color_No": "", "RRP/PH EC RRP": "", "RRP check (I=D)": "False",
                                    "SRP/PH MD Price": "", "SRP check (K=E)": "False",
                                    "Comments": "Ignore – Item Not Found in Tracker."
                                })
                                working_flow_rows.append(row_dict)
                                continue
                                
                            tracker_rrp_val = rrp_map[pim]
                            tracker_srp_val = tracker_map[pim]  
                            
                            current_rrp = _to_numeric_safe(row_dict[temp_rrp_col], sheet_rounding_strategy)
                            current_srp = _to_numeric_safe(row_dict[temp_srp_col], sheet_rounding_strategy)
                            
                            initial_rrp_match = (tracker_rrp_val == current_rrp)
                            initial_srp_match = (tracker_srp_val == current_srp)
                            
                            has_no_dates = pd.isna(row_dict[zalora_start]) or str(row_dict[zalora_start]).strip() == ""
                            is_last_day = is_last_day_of_month(row_dict[zalora_end])
                            
                            if initial_rrp_match and initial_srp_match and has_no_dates:
                                row_dict.update({
                                    "ALU_NO/Color_No": str(pim), "RRP/PH EC RRP": str(tracker_rrp_val),
                                    "RRP check (I=D)": "True", "SRP/PH MD Price": str(tracker_srp_val),
                                    "SRP check (K=E)": "True", "Comments": "All Good – RRP and SRP Match the Tracker. No Update Required."
                                })
                                working_flow_rows.append(row_dict)
                                continue
                                
                            if initial_rrp_match and initial_srp_match and is_last_day:
                                try:
                                    existing_end_dt = pd.to_datetime(row_dict[zalora_end])
                                    if existing_end_dt >= user_selected_end_dt:
                                        row_dict.update({
                                            "ALU_NO/Color_No": str(pim), "RRP/PH EC RRP": str(tracker_rrp_val),
                                            "RRP check (I=D)": "True", "SRP/PH MD Price": str(tracker_srp_val),
                                            "SRP check (K=E)": "True", "Comments": "All Good – Sale Ends at Month End. No Update Required."
                                        })
                                        working_flow_rows.append(row_dict)
                                        continue
                                    else:
                                        row_dict[zalora_end] = str(zalora_end_str)
                                        row_dict["Comments"] = "Sale End Date Updated."
                                        
                                        post_rrp_match = (tracker_rrp_val == _to_numeric_safe(row_dict[temp_rrp_col], sheet_rounding_strategy))
                                        post_srp_match = (tracker_srp_val == _to_numeric_safe(row_dict[temp_srp_col], sheet_rounding_strategy))
                                        row_dict.update({
                                            "ALU_NO/Color_No": str(pim), "RRP/PH EC RRP": str(tracker_rrp_val),
                                            "RRP check (I=D)": str(post_rrp_match), "SRP/PH MD Price": str(tracker_srp_val),
                                            "SRP check (K=E)": str(post_srp_match)
                                        })
                                        working_flow_rows.append(row_dict)
                                        final_to_upload_rows.append(row_dict.copy())
                                        continue
                                except:
                                    pass
                                
                            if initial_rrp_match and not initial_srp_match:
                                if tracker_srp_val == 0:
                                    row_dict.update({temp_srp_col: "", zalora_start: "", zalora_end: "", "Comments": "Sale Price Updated."})
                                else:
                                    row_dict.update({
                                        temp_srp_col: str(tracker_srp_val),
                                        zalora_start: str(zalora_start_str), zalora_end: str(zalora_end_str),
                                        "Comments": "Sale Price Updated."
                                    })
                                
                                post_rrp_match = (tracker_rrp_val == _to_numeric_safe(row_dict[temp_rrp_col], sheet_rounding_strategy))
                                post_srp_match = (tracker_srp_val == _to_numeric_safe(row_dict[temp_srp_col], sheet_rounding_strategy))
                                row_dict.update({
                                    "ALU_NO/Color_No": str(pim), "RRP/PH EC RRP": str(tracker_rrp_val),
                                    "RRP check (I=D)": str(post_rrp_match), "SRP/PH MD Price": str(tracker_srp_val),
                                    "SRP check (K=E)": str(post_srp_match)
                                })
                                working_flow_rows.append(row_dict)
                                final_to_upload_rows.append(row_dict.copy())
                                continue
                                
                            if not initial_rrp_match:
                                row_dict[temp_rrp_col] = str(tracker_rrp_val)
                                if tracker_srp_val == 0:
                                    row_dict.update({temp_srp_col: "", zalora_start: "", zalora_end: "", "Comments": "RRP and Sale Price Updated."})
                                else:
                                    row_dict.update({
                                        temp_srp_col: str(tracker_srp_val),
                                        zalora_start: str(zalora_start_str), zalora_end: str(zalora_end_str),
                                        "Comments": "RRP and Sale Price Updated."
                                    })
                                
                                post_rrp_match = (tracker_rrp_val == _to_numeric_safe(row_dict[temp_rrp_col], sheet_rounding_strategy))
                                post_srp_match = (tracker_srp_val == _to_numeric_safe(row_dict[temp_srp_col], sheet_rounding_strategy))
                                row_dict.update({
                                    "ALU_NO/Color_No": str(pim), "RRP/PH EC RRP": str(tracker_rrp_val),
                                    "RRP check (I=D)": str(post_rrp_match), "SRP/PH MD Price": str(tracker_srp_val),
                                    "SRP check (K=E)": str(post_srp_match)
                                })
                                working_flow_rows.append(row_dict)
                                final_to_upload_rows.append(row_dict.copy())
                                continue

                            working_flow_rows.append(row_dict)

                        pd.DataFrame(working_flow_rows).to_excel(writer, sheet_name="Working Flow", index=False)
                        
                        if final_to_upload_rows:
                            df_to_upload = pd.DataFrame(final_to_upload_rows)
                            cols_to_drop = ["ALU_NO/Color_No", "RRP/PH EC RRP", "RRP check (I=D)", "SRP/PH MD Price", "SRP check (K=E)", "Comments"]
                            df_to_upload.drop(columns=cols_to_drop, errors='ignore').to_excel(writer, sheet_name="To Upload", index=False)
                        else:
                            pd.DataFrame([{"Message": "No items required updates; all records skipped from upload file."}]).to_excel(writer, sheet_name="To Upload", index=False)

                # =========================================================
                # 🎨 POST-PROCESSING GLOBAL FORMATTING OVERRIDE MATRIX (BATCH LOOP)
                # =========================================================
                output_buffer.seek(0)
                wb = openpyxl.load_workbook(output_buffer)
                
                # Setup universal formatting objects
                consolas_font_data = Font(name="Consolas", size=10, bold=False)
                center_vertical_align = Alignment(vertical="center")
                
                for s_name in wb.sheetnames:
                    ws = wb[s_name]
                    
                    # 1. Faster Row Height Allocation via sequential indexing loop
                    for r_idx in range(1, ws.max_row + 1):
                        ws.row_dimensions[r_idx].height = 15
                        
                    # 2. Optimized Row-Level Font and Alignment Broadcast using iter_rows
                    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
                        is_header_row = (row[0].row == 1)
                        for cell in row:
                            cell.alignment = center_vertical_align
                            if is_header_row:
                                # Overwrite font family/size while keeping original background/text color intact
                                cell.font = Font(name="Consolas", size=10, bold=cell.font.bold, color=cell.font.color, italic=cell.font.italic)
                            else:
                                cell.font = consolas_font_data
                                
                    # 3. Dynamic Column Width Fitter
                    for col in ws.columns:
                        max_len = 15  # Baseline width minimum requirement
                        for cell in col:
                            val_str = str(cell.value or '')
                            if len(val_str) > max_len:
                                max_len = len(val_str)
                        col_letter = get_column_letter(col[0].column)
                        ws.column_dimensions[col_letter].width = min(max_len + 3, 50)
                
                new_buffer = io.BytesIO()
                wb.save(new_buffer)
                output_buffer = new_buffer

                output_buffer.seek(0)
                st.success("🎉 Process Complete! Consolas font rules applied globally to all generated worksheets.")
                st.download_button(
                    label="📥 Download Consolidated Marketplace Workbook",
                    data=output_buffer,
                    file_name="Consolidated_Marketplace_Pricing.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            except Exception as e:
                st.error(f"A systematic error occurred during calculations: {e}")
