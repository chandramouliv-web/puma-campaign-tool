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
# ⚙️ CACHED DATA ENGINES (PREVENTS RE-READING AND RE-PROCESSING)
# =================================================================

@st.cache_data(show_spinner="Analyzing file structure...")
def get_cached_sheet_names(file_bytes, file_name):
    """Dynamically extracts worksheet tabs from binary files."""
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
    """Loads and returns layout headers and the raw unmapped DataFrame matrix."""
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
    """Natively caches flat standard tables."""
    if file_name.endswith('.csv'):
        return pd.read_csv(io.BytesIO(file_bytes))
    try:
        return pd.read_excel(io.BytesIO(file_bytes), engine='calamine')
    except:
        return pd.read_excel(io.BytesIO(file_bytes))

@st.cache_data(show_spinner="Processing master Shopee archive...")
def _read_shopee_stream_flexible(file_bytes, file_name):
    """
    PHASE 1: EXTRACTION & INITIAL CONSOLIDATION
    Extracts all spreadsheet files inside an uploaded ZIP cluster or a single standalone file.
    Runs case-insensitive header-matching schemas inside memory arrays.
    """
    dfs = []
    if file_name.endswith('.zip'):
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
    """Scans structural headers to detect standard platform transaction export lines."""
    for idx in range(min(15, len(df))):
        row_str = df.iloc[idx].fillna("").astype(str).str.lower().str.strip().values
        if any("product id" in r or "variation id" in r or "sku ref" in r or "parent sku" in r or "seller sku" in r for r in row_str):
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
    """Applies whitespace strip, converts characters to lowercase, and changes hyphens to underscores."""
    if pd.isna(sku) or not str(sku).strip():
        return ""
    return str(sku).strip().lower().replace('-', '_')

def clean_id_str(val):
    if pd.isna(val):
        return ""
    if isinstance(val, float):
        return str(int(val)) if val.is_integer() else str(val)
    s = str(val).strip()
    if re.match(r"^-?\d+\.0+$", s):
        s = s.split(".")[0]
    return s if s != "nan" else ""

def _to_numeric_safe(val):
    if pd.isna(val):
        return 0
    try:
        num = pd.to_numeric(val, errors='coerce')
        if pd.isna(num):
            return 0
        return round(num, 2)
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
# 🎨 STREAMLIT HIGH-RESPONSIVENESS UI
# ==========================================
st.title("🚀 Marketplace Price Automator")
st.write("Dynamic pricing optimization and multi-channel structural automation.")

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
        
        tracker_headers, df_tracker = get_cached_clean_headers_and_df(t_bytes, tracker_file.name, target_sheet=tracker_sheet)
        if df_tracker is not None:
            tracker_pim = st.selectbox("Map PIM ID Column (e.g. Color_No)", [""] + tracker_headers, key="t_pim")
            tracker_rrp = st.selectbox("Map Regular RRP Column (e.g. PH EC RRP)", [""] + tracker_headers, key="t_rrp")
            tracker_md = st.selectbox("Map Campaign Markdown Price Column (e.g. PH MD Price)", [""] + tracker_headers, key="t_md")

    st.markdown("---")

    sku_file = st.file_uploader("2. Upload SKU Map File (.csv, .xlsx)", type=["csv", "xlsx"])
    sku_sku, sku_pim = None, None
    if sku_file:
        s_bytes = sku_file.getvalue()
        try:
            sku_headers = list(read_cached_file_standard(s_bytes, sku_file.name).columns)
            
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
    shopee_sku, shopee_parent, shopee_orig, shopee_promo = None, None, None, None
    shopee_start_col, shopee_end_col = None, None
    if mode in ["🛍 Shopee Only", "🔄 Run All Marketplace Channels"]:
        shopee_file = st.file_uploader("3. Upload Shopee Master Template (.zip, .xlsx, .csv)", type=["zip", "xlsx", "csv"])
        if shopee_file:
            sh_bytes = shopee_file.getvalue()
            try:
                df_shopee_preview = _read_shopee_stream_flexible(sh_bytes, shopee_file.name)
                shopee_headers = list(df_shopee_preview.columns)
                
                shopee_sku = st.selectbox("Map Shopee Seller SKU Column", [""] + shopee_headers, key="sh_sku")
                shopee_parent = st.selectbox("Map Shopee Parent SKU Column", [""] + shopee_headers, key="sh_parent")
                shopee_orig = st.selectbox("Map Shopee Original Price / RRP Column", [""] + shopee_headers, key="sh_orig")
                shopee_promo = st.selectbox("Map Shopee Promotion Price Column", [""] + shopee_headers, key="sh_promo")
                shopee_start_col = st.selectbox("Map Shopee Start Date Column", [""] + shopee_headers, key="sh_start")
                shopee_end_col = st.selectbox("Map Shopee End Date Column", [""] + shopee_headers, key="sh_end")
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
# ⚙️ PROCESSING EXECUTION CORE (VECTORIZED)
# ==========================================
if st.button("🚀 Run Automation Process", type="primary", use_container_width=True):
    error_found = False
    if not tracker_file or not tracker_pim or not tracker_rrp or not tracker_md:
        st.error("❌ Tracker layout mappings are invalid."); error_found = True
    if not sku_file or not sku_sku or not sku_pim:
        st.error("❌ Reference SKU mapping parameters must be fully bound."); error_found = True
    if mode in ["🛍 Shopee Only", "🔄 Run All Marketplace Channels"] and (not shopee_file or not shopee_sku or not shopee_orig or not shopee_promo):
        st.error("❌ Shopee template columns (SKU, Original Price, and Promotion Price) remain unassigned."); error_found = True
    if mode in ["🏪 Lazada Only", "🔄 Run All Marketplace Channels"] and (not lazada_file or not lazada_sku or not lazada_price):
        st.error("❌ Lazada columns remain unassigned."); error_found = True
    if mode in ["👗 Zalora Only", "🔄 Run All Marketplace Channels"] and (not zalora_file or not zalora_sku or not zalora_promo or not zalora_orig or not zalora_start or not zalora_end):
        st.error("❌ Zalora tracking dimensions are unassigned."); error_found = True

    if not error_found:
        s_bytes = sku_file.getvalue()
        t_bytes = tracker_file.getvalue()
        
        with st.spinner("Executing optimized system pipeline lookups..."):
            try:
                # -------------------------------------------------------------
                # PHASE 2: MASTER TRACKER & SKU MAPPING
                # -------------------------------------------------------------
                df_sku = read_cached_file_standard(s_bytes, sku_file.name)
                sku_raw_arr = df_sku[sku_sku].astype(str).values
                pim_raw_arr = df_sku[sku_pim].astype(str).values
                
                normalized_sku_to_pim = {}
                normalized_sku_to_new_price = {}
                
                tracker_map_new_price = {}
                tracker_map_rrp = {}
                
                pim_tracker_raw = df_tracker[tracker_pim].values
                rrp_tracker_raw = pd.to_numeric(df_tracker[tracker_rrp], errors='coerce').fillna(0).values
                md_tracker_raw = pd.to_numeric(df_tracker[tracker_md], errors='coerce').fillna(0).values
                
                for t_idx in range(len(df_tracker)):
                    pim_val = clean_id_str(pim_tracker_raw[t_idx])
                    if not pim_val: 
                        continue
                    
                    raw_rrp = rrp_tracker_raw[t_idx]
                    raw_md = md_tracker_raw[t_idx]
                    
                    # Logic Condition: Rule for calculating New Price
                    if raw_md == 0 or pd.isna(raw_md):
                        new_price = round(raw_rrp)
                    else:
                        new_price = round(raw_md)
                        
                    tracker_map_new_price[pim_val] = new_price
                    tracker_map_rrp[pim_val] = round(raw_rrp)

                # Map Normalized SKU -> Targets
                for idx_s in range(len(sku_raw_arr)):
                    raw_s = sku_raw_arr[idx_s].strip()
                    norm_s = normalize_sku(raw_s)
                    pim_clean = clean_id_str(pim_raw_arr[idx_s])
                    
                    if norm_s and pim_clean:
                        normalized_sku_to_pim[norm_s] = pim_clean
                        if pim_clean in tracker_map_new_price:
                            normalized_sku_to_new_price[norm_s] = tracker_map_new_price[pim_clean]

                output_buffer = io.BytesIO()
                with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
                    
                    # -------------------------------------------------------------
                    # PHASE 1 & 3: SHOPEE ENGINE REFACTOR
                    # -------------------------------------------------------------
                    if shopee_file and mode in ["🛍 Shopee Only", "🔄 Run All Marketplace Channels"]:
                        sh_bytes = shopee_file.getvalue()
                        df_shopee_raw = _read_shopee_stream_flexible(sh_bytes, shopee_file.name)
                        
                        if not df_shopee_raw.empty:
                            shopee_records = df_shopee_raw.to_dict('records')
                            
                            shopee_consolidated_output = []
                            shopee_rrp_mismatches = []
                            shopee_upload_list = []
                            
                            for row in shopee_records:
                                # Fallback logic requirement: If empty Seller SKU, copy Parent SKU
                                raw_sku = str(row.get(shopee_sku, '')).strip()
                                raw_parent = str(row.get(shopee_parent, '')).strip()
                                
                                if not raw_sku or raw_sku.lower() == 'nan':
                                    raw_sku = raw_parent
                                    row[shopee_sku] = raw_sku
                                
                                norm_sku = normalize_sku(raw_sku)
                                target_pim = normalized_sku_to_pim.get(norm_sku, "")
                                
                                qc_comment = "Ignore – Item Not Found in Tracker."
                                belongs_in_upload = False
                                
                                start_dt = row.get(shopee_start_col, "") if shopee_start_col else ""
                                end_dt = row.get(shopee_end_col, "") if shopee_end_col else ""
                                current_shopee_promo = row.get(shopee_promo, 0)
                                
                                if target_pim in tracker_map_rrp:
                                    tracker_rrp_val = tracker_map_rrp[target_pim]
                                    tracker_new_price = tracker_map_new_price[target_pim]
                                    current_orig_price = _to_numeric_safe(row.get(shopee_orig, 0))
                                    
                                    # Order of Operations Checks from Phase 3:
                                    # Check 1: RRP Deviation Checks
                                    if current_orig_price != tracker_rrp_val:
                                        qc_comment = "RRP Mismatch"
                                        shopee_rrp_mismatches.append({
                                            "Seller SKU": raw_sku,
                                            "Marketplace Status": "Active",
                                            "Marketplace Message": "RRP Mismatch",
                                            "RRP": tracker_rrp_val,
                                            "Sale Amount": current_shopee_promo,
                                            "Sale Start Date": start_dt,
                                            "Sale End Date": end_dt
                                        })
                                    # Check 2: Missing/Blank Target Pricing
                                    elif not tracker_new_price or pd.isna(tracker_new_price) or tracker_new_price == 0:
                                        qc_comment = "Discount Price is Blank"
                                    # Check 3: Promo Equalizer Constraint
                                    elif tracker_rrp_val == tracker_new_price:
                                        qc_comment = "Remove: RRP = Discount"
                                    else:
                                        # Clean Upload Pipeline
                                        qc_comment = "Ready for Upload"
                                        row[shopee_promo] = tracker_new_price
                                        belongs_in_upload = True
                                
                                # Append directly to Phase 4 outputs
                                row_consolidated = row.copy()
                                row_consolidated["QC Comment"] = qc_comment
                                shopee_consolidated_output.append(row_consolidated)
                                
                                if belongs_in_upload:
                                    shopee_upload_list.append(row.copy())

                            # Write Phase 4 worksheets
                            pd.DataFrame(shopee_consolidated_output).to_excel(writer, sheet_name="Consolidated File", index=False)
                            
                            df_mismatches_final = pd.DataFrame(shopee_rrp_mismatches) if shopee_rrp_mismatches else pd.DataFrame(columns=["Seller SKU", "Marketplace Status", "Marketplace Message", "RRP", "Sale Amount", "Sale Start Date", "Sale End Date"])
                            df_mismatches_final.to_excel(writer, sheet_name="Shopee_RRP_Mismatches", index=False)
                            
                            df_upload_final = pd.DataFrame(shopee_upload_list) if shopee_upload_list else pd.DataFrame(columns=df_shopee_raw.columns)
                            df_upload_final.to_excel(writer, sheet_name="Shopee_Upload", index=False)

                    # -------------------------------------------------------------
                    # 4. Process Lazada Channel Data
                    # -------------------------------------------------------------
                    if lazada_file and mode in ["🏪 Lazada Only", "🔄 Run All Marketplace Channels"]:
                        lz_bytes = lazada_file.getvalue()
                        df_lazada = read_cached_file_standard(lz_bytes, lazada_file.name)
                        lazada_skus = df_lazada[lazada_sku].values
                        lazada_prices = df_lazada[lazada_price].values
                        
                        final_lazada_prices = []
                        for i in range(len(df_lazada)):
                            norm_lz_sku = normalize_sku(lazada_skus[i])
                            pim = normalized_sku_to_pim.get(norm_lz_sku)
                            
                            if pim and pim in tracker_map_rrp:
                                tracker_rrp_val = tracker_map_rrp[pim]
                                tracker_srp_val = tracker_map_new_price[pim]
                                computed_campaign_price = tracker_rrp_val if tracker_srp_val == 0 else tracker_srp_val
                                final_lazada_prices.append(computed_campaign_price)
                            else:
                                final_lazada_prices.append(lazada_prices[i])
                                
                        df_lazada[lazada_price] = final_lazada_prices
                        df_lazada.to_excel(writer, sheet_name="Lazada_Upload", index=False)

                    # -------------------------------------------------------------
                    # 5. Process Zalora Channel Data
                    # -------------------------------------------------------------
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
                            pim = normalized_sku_to_pim.get(norm_sku)
                            
                            if not pim or pim not in tracker_map_rrp:
                                row_dict.update({"ALU_NO/Color_No": "", "RRP/PH EC RRP": "", "RRP check (I=D)": "False", "SRP/PH MD Price": "", "SRP check (K=E)": "False", "Comments": "Ignore – Item Not Found in Tracker."})
                                working_flow_rows.append(row_dict)
                                continue
                                
                            tracker_rrp_val = tracker_map_rrp[pim]
                            tracker_srp_val = tracker_map_new_price[pim]  
                            current_rrp = _to_numeric_safe(row_dict[temp_rrp_col])
                            current_srp = _to_numeric_safe(row_dict[temp_srp_col])
                            
                            initial_rrp_match = (tracker_rrp_val == current_rrp)
                            initial_srp_match = (tracker_srp_val == current_srp)
                            has_no_dates = pd.isna(row_dict[zalora_start]) or str(row_dict[zalora_start]).strip() == ""
                            is_last_day = is_last_day_of_month(row_dict[zalora_end])
                            
                            if initial_rrp_match and initial_srp_match and has_no_dates:
                                row_dict.update({"ALU_NO/Color_No": str(pim), "RRP/PH EC RRP": str(tracker_rrp_val), "RRP check (I=D)": "True", "SRP/PH MD Price": str(tracker_srp_val), "SRP check (K=E)": "True", "Comments": "All Good – RRP and SRP Match the Tracker. No Update Required."})
                                working_flow_rows.append(row_dict)
                                continue
                                
                            if initial_rrp_match and initial_srp_match and is_last_day:
                                try:
                                    existing_end_dt = pd.to_datetime(row_dict[zalora_end])
                                    if existing_end_dt >= user_selected_end_dt:
                                        row_dict.update({"ALU_NO/Color_No": str(pim), "RRP/PH EC RRP": str(tracker_rrp_val), "RRP check (I=D)": "True", "SRP/PH MD Price": str(tracker_srp_val), "SRP check (K=E)": "True", "Comments": "All Good – Sale Ends at Month End. No Update Required."})
                                        working_flow_rows.append(row_dict)
                                        continue
                                    else:
                                        row_dict[zalora_end] = str(zalora_end_str)
                                        row_dict["Comments"] = "Sale End Date Updated."
                                        post_rrp_match = (tracker_rrp_val == _to_numeric_safe(row_dict[temp_rrp_col]))
                                        post_srp_match = (tracker_srp_val == _to_numeric_safe(row_dict[temp_srp_col]))
                                        row_dict.update({"ALU_NO/Color_No": str(pim), "RRP/PH EC RRP": str(tracker_rrp_val), "RRP check (I=D)": str(post_rrp_match), "SRP/PH MD Price": str(tracker_srp_val), "SRP check (K=E)": str(post_srp_match)})
                                        working_flow_rows.append(row_dict)
                                        final_to_upload_rows.append(row_dict.copy())
                                        continue
                                except:
                                    pass
                                
                            if initial_rrp_match and not initial_srp_match:
                                if tracker_srp_val == 0:
                                    row_dict.update({temp_srp_col: "", zalora_start: "", zalora_end: "", "Comments": "Sale Price Updated."})
                                else:
                                    row_dict.update({temp_srp_col: str(tracker_srp_val), zalora_start: str(zalora_start_str), zalora_end: str(zalora_end_str), "Comments": "Sale Price Updated."})
                                
                                post_rrp_match = (tracker_rrp_val == _to_numeric_safe(row_dict[temp_rrp_col]))
                                post_srp_match = (tracker_srp_val == _to_numeric_safe(row_dict[temp_srp_col]))
                                row_dict.update({"ALU_NO/Color_No": str(pim), "RRP/PH EC RRP": str(tracker_rrp_val), "RRP check (I=D)": str(post_rrp_match), "SRP/PH MD Price": str(tracker_srp_val), "SRP check (K=E)": str(post_srp_match)})
                                working_flow_rows.append(row_dict)
                                final_to_upload_rows.append(row_dict.copy())
                                continue
                                
                            if not initial_rrp_match:
                                row_dict[temp_rrp_col] = str(tracker_rrp_val)
                                if tracker_srp_val == 0:
                                    row_dict.update({temp_srp_col: "", zalora_start: "", zalora_end: "", "Comments": "RRP and Sale Price Updated."})
                                else:
                                    row_dict.update({temp_srp_col: str(tracker_srp_val), zalora_start: str(zalora_start_str), zalora_end: str(zalora_end_str), "Comments": "RRP and Sale Price Updated."})
                                
                                post_rrp_match = (tracker_rrp_val == _to_numeric_safe(row_dict[temp_rrp_col]))
                                post_srp_match = (tracker_srp_val == _to_numeric_safe(row_dict[temp_srp_col]))
                                row_dict.update({"ALU_NO/Color_No": str(pim), "RRP/PH EC RRP": str(tracker_rrp_val), "RRP check (I=D)": str(post_rrp_match), "SRP/PH MD Price": str(tracker_srp_val), "SRP check (K=E)": str(post_srp_match)})
                                working_flow_rows.append(row_dict)
                                final_to_upload_rows.append(row_dict.copy())
                                continue

                            working_flow_rows.append(row_dict)

                        pd.DataFrame(working_flow_rows).to_excel(writer, sheet_name="Working Flow", index=False)
                        
                        if final_to_upload_rows:
                            df_to_upload = pd.DataFrame(final_to_upload_rows)
                            df_to_upload.drop(columns=["ALU_NO/Color_No", "RRP/PH EC RRP", "RRP check (I=D)", "SRP/PH MD Price", "SRP check (K=E)", "Comments"], errors='ignore').to_excel(writer, sheet_name="To Upload", index=False)
                        else:
                            pd.DataFrame([{"Message": "No updates required."}]).to_excel(writer, sheet_name="To Upload", index=False)

                # =========================================================
                # 🎨 POST-PROCESSING STYLE ENGINE: ENFORCE EX_ORANGE HEADERS
                # =========================================================
                output_buffer.seek(0)
                wb = openpyxl.load_workbook(output_buffer)
                
                target_styled_sheets = ["Shopee_Upload", "To Upload", "Lazada_Upload", "Shopee_RRP_Mismatches"]
                for sheet_name in wb.sheetnames:
                    if sheet_name in target_styled_sheets or "Upload" in sheet_name:
                        ws = wb[sheet_name]
                        orange_fill = PatternFill(start_color="FF5722", end_color="FF5722", fill_type="solid")
                        white_bold_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
                        thin_side = openpyxl.styles.Side(style='thin', color='CCCCCC')
                        clean_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
                        center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
                        
                        ws.row_dimensions[1].height = 28
                        for col_idx in range(1, ws.max_column + 1):
                            cell = ws.cell(row=1, column=col_idx)
                            cell.fill = orange_fill
                            cell.font = white_bold_font
                            cell.border = clean_border
                            cell.alignment = center_align
                        
                        for col in ws.columns:
                            max_len = max(len(str(cell.value or '')) for cell in col)
                            col_letter = get_column_letter(col[0].column)
                            ws.column_dimensions[col_letter].width = max(max_len + 3, 14)
                            
                new_buffer = io.BytesIO()
                wb.save(new_buffer)
                output_buffer = new_buffer

                output_buffer.seek(0)
                st.success("🎉 Process Complete! Platform-ready files compiled successfully.")
                st.download_button(
                    label="📥 Download Consolidated Marketplace Workbook",
                    data=output_buffer,
                    file_name="Consolidated_Marketplace_Pricing.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            except Exception as e:
                st.error(f"A systematic error occurred during processing loops: {e}")
