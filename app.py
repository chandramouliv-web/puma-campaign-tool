import streamlit as st
import pandas as pd
import io
import re
import zipfile

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
    """
    Clean an ID value (Product ID, Shop SKU, etc.) for exact output.
    Prevents floating-point trailing digits (.0) from leaking into string exports.
    """
    if pd.isna(val):
        return None
    if isinstance(val, float):
        return str(int(val)) if val.is_integer() else str(val)
    s = str(val).strip()
    if re.match(r"^-?\d+\.0+$", s):
        s = s.split(".")[0]
    return s

def _extract_ean(sku_val, parent_val):
    """Safely extracts a 13-digit EAN from layout configurations."""
    for v in (sku_val, parent_val):
        if pd.notna(v):
            try:
                s = str(int(float(v)))
                if re.match(r"^\d{13}$", s): return s
            except (ValueError, TypeError):
                s = str(v).strip()
                if re.match(r"^\d{13}$", s): return s
    return None

def get_clean_headers_and_df(uploaded_file):
    """
    Reads Row 3 directly as clean column names and maps them 
    with their Excel column letters to avoid duplicate name confusion.
    """
    try:
        uploaded_file.seek(0)
        if uploaded_file.name.endswith('.csv'):
            full_df = pd.read_csv(uploaded_file, header=None)
        else:
            full_df = pd.read_excel(uploaded_file, header=None)

        # Extract row 3 (Index 2) as target header strings
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
    """Standard single-row header reader for SKU Maps or Excel Templates."""
    uploaded_file.seek(0)
    if uploaded_file.name.endswith('.csv'):
        return pd.read_csv(uploaded_file)
    return pd.read_excel(uploaded_file)

def _read_shopee_zip(uploaded_file):
    """
    Unpacks an uploaded Shopee ZIP stream containing multiple standard 
    export tables using the optimized calamine engine.
    """
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
    
    # Tracker Upload Container
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

    # SKU Map Upload Container
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
    
    # Shopee ZIP Upload block
    shopee_file = None
    shopee_sku, shopee_promo, shopee_orig, shopee_start, shopee_end = None, None, None, None, None
    if mode in ["🛍 Shopee Only", "🔄 Run All Marketplace Channels"]:
        shopee_file = st.file_uploader("3. Upload Shopee Master Template (.zip)", type=["zip"])
        if shopee_file:
            try:
                df_shopee_preview = _read_shopee_zip(shopee_file)
                shopee_headers = list(df_shopee_preview.columns)
                
                shopee_sku = st.selectbox("Map Shopee SKU Column", [""] + shopee_headers, key="sh_sku")
                shopee_promo = st.selectbox("Map Shopee Promotion/Discount Price Column", [""] + shopee_headers, key="sh_promo")
                shopee_orig = st.selectbox("Map Shopee Original Price Column", [""] + shopee_headers, key="sh_orig")
                shopee_start = st.selectbox("Map Shopee Start Date (Optional)", [""] + shopee_headers, key="sh_start")
                shopee_end = st.selectbox("Map Shopee End Date (Optional)", [""] + shopee_headers, key="sh_end")
            except Exception as e:
                st.error(f"Could not open or parse Shopee ZIP package stream: {e}")

    # Lazada Upload block
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

    # Zalora Upload block
    zalora_file = None
    zalora_sku, zalora_price = None, None
    if mode in ["👗 Zalora Only", "🔄 Run All Marketplace Channels"]:
        st.markdown("---")
        zalora_file = st.file_uploader("5. Upload Zalora Master Template (.csv, .xlsx)", type=["csv", "xlsx"])
        if zalora_file:
            try:
                zalora_headers = list(read_full_file_standard(zalora_file).columns)
                zalora_sku = st.selectbox("Map Zalora SKU / Item Column", [""] + zalora_headers, key="zal_sku")
                zalora_price = st.selectbox("Map Zalora Target Price Column", [""] + zalora_headers, key="zal_price")
            except Exception as e:
                st.error(f"Could not parse Zalora layout: {e}")

st.markdown("---")

# ==========================================
# ⚙️ PROCESSING EXECUTION CORE
# ==========================================
if st.button("🚀 Run Automation Process", type="primary", use_container_width=True):
    
    # Form Validation Passports
    error_found = False
    if not tracker_file or not tracker_pim or not tracker_rrp or not tracker_md:
        st.error("❌ Tracker layout mappings are invalid or unassigned."); error_found = True
    if not sku_file or not sku_sku or not sku_pim:
        st.error("❌ Universal reference SKU mapping parameters must be fully bound."); error_found = True
    if mode in ["🛍 Shopee Only", "🔄 Run All Marketplace Channels"] and (not shopee_file or not shopee_sku or not shopee_promo or not shopee_orig):
        st.error("❌ Shopee engine selected, but tracking dimensions are unassigned."); error_found = True
    if mode in ["🏪 Lazada Only", "🔄 Run All Marketplace Channels"] and (not lazada_file or not lazada_sku or not lazada_price):
        st.error("❌ Lazada engine selected, but data columns remain unassigned."); error_found = True
    if mode in ["👗 Zalora Only", "🔄 Run All Marketplace Channels"] and (not zalora_file or not zalora_sku or not zalora_price):
        st.error("❌ Zalora engine selected, but data columns remain unassigned."); error_found = True

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
                with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
                    
                    # 3. Process Shopee Channel Data via unpacked ZIP arrays
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
                                    "Sale Start Date(SGT)": row.get(shopee_start, '') if shopee_start else '', 
                                    "Sale End Date(SGT)": row.get(shopee_end, '') if shopee_end else ''
                                })
                            elif pd.isna(new_price) or new_price == "":
                                comment = "Discount Price is Blank"
                            elif rrp is not None and rrp == new_price:
                                comment = "Remove: RRP = Discount"
                            
                            df_shopee.at[idx, shopee_promo] = new_price
                            df_shopee.at[idx, 'QC Comment'] = comment
                            
                            if not is_mismatch and not (pd.isna(new_price) or new_price == "") and not (rrp is not None and rrp == new_price):
                                upload_row = row.copy()
                                upload_row[shopee_promo] = new_price
                                upload_rows.append(upload_row)

                        df_shopee.to_excel(writer, sheet_name="Shopee_Master_Updated", index=False)
                        
                        df_mismatch = pd.DataFrame(mismatch_rows) if mismatch_rows else pd.DataFrame([{"Message": "No RRP Mismatches Found"}])
                        df_mismatch.to_excel(writer, sheet_name="Shopee_RRP_Mismatches", index=False)
                        
                        if upload_rows:
                            df_upload = pd.DataFrame(upload_rows).drop(columns=['QC Comment'], errors='ignore')
                            df_upload.to_excel(writer, sheet_name="Shopee_Upload", index=False)

                    # 4. Process Lazada Channel Data
                    if lazada_file and mode in ["🏪 Lazada Only", "🔄 Run All Marketplace Channels"]:
                        df_lazada = read_full_file_standard(lazada_file)
                        for idx, row in df_lazada.iterrows():
                            norm_sku = normalize_sku(row[lazada_sku])
                            existing_price = row[lazada_price]
                            new_price = price_map.get(norm_sku, existing_price)
                            df_lazada.at[idx, lazada_price] = new_price
                        
                        df_lazada.to_excel(writer, sheet_name="Lazada_Upload", index=False)

                    # 5. Process Zalora Channel Data
                    if zalora_file and mode in ["👗 Zalora Only", "🔄 Run All Marketplace Channels"]:
                        df_zalora = read_full_file_standard(zalora_file)
                        for idx, row in df_zalora.iterrows():
                            norm_sku = normalize_sku(row[zalora_sku])
                            existing_price = row[zalora_price]
                            new_price = price_map.get(norm_sku, existing_price)
                            df_zalora.at[idx, zalora_price] = new_price
                        
                        df_zalora.to_excel(writer, sheet_name="Zalora_Upload", index=False)

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
