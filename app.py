import streamlit as st
import pandas as pd
import io

# Set up page configuration
st.set_page_config(page_title="Marketplace Price Automator", page_icon="🚀", layout="wide")

def normalize_sku(sku):
    if pd.isna(sku) or not sku:
        return ""
    return str(sku).strip().replace('-', '_').lower()

def get_clean_headers_and_df(uploaded_file):
    """
    Reads the file, detects multi-row headers, flattens them into intuitive 
    strings, and returns the cleaned list of headers along with the processed DataFrame.
    """
    try:
        # Load the first 5 rows to analyze the header structure
        if uploaded_file.name.endswith('.csv'):
            preview_df = pd.read_csv(uploaded_file, nrows=5, header=None)
        else:
            preview_df = pd.read_excel(uploaded_file, nrows=5, header=None)
        
        # Read full file without headers initially so we can build them manually
        uploaded_file.seek(0)
        if uploaded_file.name.endswith('.csv'):
            full_df = pd.read_csv(uploaded_file, header=None)
        else:
            full_df = pd.read_excel(uploaded_file, header=None)

        # Force all header rows to be pure strings to eliminate 'float' errors on empty cells
        row1 = preview_df.iloc[0].fillna("").astype(str).tolist()
        row2 = preview_df.iloc[1].fillna("").astype(str).tolist()
        row3 = preview_df.iloc[2].fillna("").astype(str).tolist()

        # Forward-fill merged row values to handle gaps left by merged cells
        last_r1 = ""
        last_r2 = ""
        combined_headers = []
        
        for r1, r2, r3 in zip(row1, row2, row3):
            # Clean up default pandas naming artifacts and whitespace
            r1_clean = r1.strip() if r1 != "nan" and "Unnamed:" not in r1 else ""
            r2_clean = r2.strip() if r2 != "nan" and "Unnamed:" not in r2 else ""
            r3_clean = r3.strip() if r3 != "nan" and "Unnamed:" not in r3 else ""
            
            if r1_clean: last_r1 = r1_clean
            if r2_clean: last_r2 = r2_clean
            
            # Construct a clean, meaningful hierarchy
            parts = []
            if last_r1 and "Onwards" not in last_r1: parts.append(last_r1)
            if last_r2: parts.append(last_r2)
            if r3_clean: parts.append(r3_clean)
            
            # Use a fallback placeholder if the column is entirely unnamed
            header_name = " - ".join(parts) if parts else "Unnamed Column"
            combined_headers.append(header_name)

        # Reassign the combined names to the dataframe and drop the multi-row header rows
        full_df.columns = combined_headers
        full_df = full_df.iloc[3:].reset_index(drop=True)
        
        return combined_headers, full_df

    except Exception as e:
        st.error(f"Error flattening headers from {uploaded_file.name}: {e}")
        return [], None

def read_full_file_standard(uploaded_file):
    """Standard file reader for simple single-row header sheets like SKU map."""
    uploaded_file.seek(0)
    if uploaded_file.name.endswith('.csv'):
        return pd.read_csv(uploaded_file)
    return pd.read_excel(uploaded_file)

# ==========================================
# 🎨 STREAMLIT UI
# ==========================================
st.title("🚀 Marketplace Price Automator")
st.write("Upload your structural files, map the columns dynamically, and generate production-ready pricing exports.")

# 1. Global Mode Selection
mode = st.selectbox(
    "Select Automation Mode",
    ["🔄 Run Both Marketplace Channels", "🛍 Shopee Only", "🏪 Lazada Only"],
    index=0
)

st.markdown("---")
col1, col2 = st.columns(2)

# 2. Sidebar / Column Configuration Containers
with col1:
    st.subheader("📋 Core Data Settings")
    
    # Tracker Upload with Flattener applied
    tracker_file = st.file_uploader("1. Upload Master Tracker File (.csv, .xlsx)", type=["csv", "xlsx"])
    tracker_pim, tracker_rrp, tracker_md = None, None, None
    df_tracker = None
    
    if tracker_file:
        tracker_headers, df_tracker = get_clean_headers_and_df(tracker_file)
        if df_tracker is not None:
            st.success("💡 Detected multi-row headers and successfully cleaned them up!")
            tracker_pim = st.selectbox("Map PIM ID Column", [""] + tracker_headers, key="t_pim")
            tracker_rrp = st.selectbox("Map RRP Column (e.g. PH EC RRP)", [""] + tracker_headers, key="t_rrp")
            tracker_md = st.selectbox("Map Markdown Price Column (e.g. PH MD Price)", [""] + tracker_headers, key="t_md")

    st.markdown("---")

    # SKU Map Upload
    sku_file = st.file_uploader("2. Upload SKU Map File (.csv, .xlsx)", type=["csv", "xlsx"])
    sku_sku, sku_pim = None, None
    if sku_file:
        try:
            sku_headers = list(read_full_file_standard(sku_file).columns)
            sku_sku = st.selectbox("Map Seller SKU Column", [""] + sku_headers, key="s_sku")
            sku_pim = st.selectbox("Map PIM ID Column", [""] + sku_headers, key="s_pim")
        except:
            st.error("Could not parse SKU Map layout.")

with col2:
    st.subheader("🛍 Marketplace Templates")
    
    # Shopee File block
    shopee_file = None
    shopee_sku, shopee_promo, shopee_orig, shopee_start, shopee_end = None, None, None, None, None
    if "Shopee" in mode or "Both" in mode:
        shopee_file = st.file_uploader("3. Upload Shopee Master Template (.csv, .xlsx)", type=["csv", "xlsx"])
        if shopee_file:
            try:
                shopee_headers = list(read_full_file_standard(shopee_file).columns)
                shopee_sku = st.selectbox("Map SKU Column", [""] + shopee_headers, key="sh_sku")
                shopee_promo = st.selectbox("Map Promotion/Discount Price Column", [""] + shopee_headers, key="sh_promo")
                shopee_orig = st.selectbox("Map Original Price Column", [""] + shopee_headers, key="sh_orig")
                shopee_start = st.selectbox("Map Start Date (Optional)", [""] + shopee_headers, key="sh_start")
                shopee_end = st.selectbox("Map End Date (Optional)", [""] + shopee_headers, key="sh_end")
            except:
                st.error("Could not parse Shopee layout.")

    # Lazada File block
    lazada_file = None
    lazada_sku, lazada_price = None, None
    if "Lazada" in mode or "Both" in mode:
        if "Both" in mode: st.markdown("---")
        lazada_file = st.file_uploader("4. Upload Lazada Master Template (.csv, .xlsx)", type=["csv", "xlsx"])
        if lazada_file:
            try:
                lazada_headers = list(read_full_file_standard(lazada_file).columns)
                lazada_sku = st.selectbox("Map Lazada Seller SKU Column", [""] + lazada_headers, key="lz_sku")
                lazada_price = st.selectbox("Map Special/Campaign Price Column", [""] + lazada_headers, key="lz_price")
            except:
                st.error("Could not parse Lazada layout.")

st.markdown("---")

# ==========================================
# ⚙️ BACKEND PROCESSING CORE
# ==========================================
if st.button("🚀 Run Automation Process", type="primary", use_container_width=True):
    
    # Validation Safeguards
    error_found = False
    if not tracker_file or not tracker_pim or not tracker_rrp or not tracker_md:
        st.error("❌ Please upload the Tracker file and completely map PIM, RRP, and MD columns."); error_found = True
    if not sku_file or not sku_sku or not sku_pim:
        st.error("❌ Please upload the SKU Map file and completely map SKU and PIM columns."); error_found = True
    if ("Shopee" in mode or "Both" in mode) and (not shopee_file or not shopee_sku or not shopee_promo or not shopee_orig):
        st.error("❌ Shopee is selected, but dynamic structural column mapping is incomplete."); error_found = True
    if ("Lazada" in mode or "Both" in mode) and (not lazada_file or not lazada_sku or not lazada_price):
        st.error("❌ Lazada is selected, but dynamic structural column mapping is incomplete."); error_found = True

    if not error_found:
        with st.spinner("Processing files and calculating pricing logic..."):
            try:
                df_sku = read_full_file_standard(sku_file)
                
                # 1. Map Master Tracker Storage
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

                # 2. Cross-reference SKU Map
                price_map = {}
                sku_to_pim = {}
                for _, row in df_sku.iterrows():
                    norm_sku = normalize_sku(row[sku_sku])
                    pim = row[sku_pim]
                    if pim in tracker_map:
                        price_map[norm_sku] = tracker_map[pim]
                        sku_to_pim[norm_sku] = pim

                # Open up clean multi-sheet buffer stream
                output_buffer = io.BytesIO()
                with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
                    
                    # 3. Handle Shopee Processing Pipeline
                    if shopee_file and ("Shopee" in mode or "Both" in mode):
                        df_shopee = read_full_file_standard(shopee_file)
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

                    # 4. Handle Lazada Processing Pipeline
                    if lazada_file and ("Lazada" in mode or "Both" in mode):
                        df_lazada = read_full_file_standard(lazada_file)
                        for idx, row in df_lazada.iterrows():
                            norm_sku = normalize_sku(row[lazada_sku])
                            existing_price = row[lazada_price]
                            new_price = price_map.get(norm_sku, existing_price)
                            df_lazada.at[idx, lazada_price] = new_price
                        
                        df_lazada.to_excel(writer, sheet_name="Lazada_Upload", index=False)

                output_buffer.seek(0)
                
                st.success("🎉 Automation executed cleanly!")
                st.download_button(
                    label="📥 Download Processed Pricing Excel Workbook",
                    data=output_buffer,
                    file_name="Automated_Marketplace_Pricing.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            except Exception as e:
                st.error(f"An unexpected systematic failure occurred: {e}")
