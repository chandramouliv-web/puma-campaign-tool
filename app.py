import streamlit as st
import pandas as pd
import io
import re
import zipfile

# Set up page configuration
st.set_page_config(page_title="Marketplace Price & Eligibility Automator", page_icon="🚀", layout="wide")

# =================================────────────────────────────────
# ⚙️ HELPER FUNCTIONS & PARSERS
# =================================────────────────────────────────

def normalize_sku(sku):
    if pd.isna(sku) or not sku:
        return ""
    return str(sku).strip().replace('-', '_').lower()

def clean_id_str(val):
    """
    Clean an ID value (Product ID, Shop SKU, etc.) for exact output.
    Handles numeric ID column casting issues in pandas.
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
    """Safely extracts a 13-digit EAN from variation strings."""
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
    """Reads Row 3 directly as clean column names mapped with Excel letters."""
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

# Mock helper sets required by your logic processors
def eligible_ean_set(ean_df):
    return set(ean_df["EAN"].dropna().astype(str).str.strip()) if "EAN" in ean_df.columns else set()

def excluded_ean_set(ean_df):
    if "Status" in ean_df.columns:
        return set(ean_df[ean_df["Status"].astype(str).str.upper() == "NO"]["EAN"].dropna().astype(str).str.strip())
    return set()

def no_remark_ean_set(ean_df):
    if "Remark" in ean_df.columns:
        return set(ean_df[ean_df["Remark"].isna()]["EAN"].dropna().astype(str).str.strip())
    return set()

# =================================────────────────────────────────
# 🛒 MARKETPLACE PROCESSORS
# =================================────────────────────────────────

def process_lazada(ean_df, lazada_bytes):
    df = pd.read_excel(io.BytesIO(lazada_bytes), sheet_name="template", header=0)
    data = df.iloc[3:].copy()
    data.columns = df.columns
    active = data[data["status"].astype(str).str.lower() == "active"].copy()
    active["_ean"] = active["SellerSKU"].astype(str).str.strip()
    ok = eligible_ean_set(ean_df)
    skus = active[active["_ean"].isin(ok)]["Shop SKU"].dropna().apply(clean_id_str)
    return skus.dropna().unique().tolist()

def _read_shopee_zip(zip_bytes):
    dfs = []
    with zipfile.ZipFile(io.BytesIO(zip_bytes)) as zf:
        names = sorted(n for n in zf.namelist() if n.endswith(".xlsx"))
        bar = st.progress(0, text="Reading Shopee export files…")
        for i, name in enumerate(names):
            with zf.open(name) as f:
                dfs.append(pd.read_excel(f, engine="calamine", header=2, skiprows=[3, 4]))
            bar.progress((i + 1) / len(names), text=f"Reading Shopee file {i+1}/{len(names)}…")
        bar.empty()
    return pd.concat(dfs, ignore_index=True)

def build_pid_decisions(combined_df, ean_df):
    ok_eans = eligible_ean_set(ean_df)
    excl_eans = excluded_ean_set(ean_df)
    
    # Structural fallback for properties dictionary mapping
    for col in ["article", "reason"]:
        if col not in ean_df.columns: ean_df[col] = "N/A"
        
    ean_info = ean_df.drop_duplicates(subset=["EAN"]).set_index("EAN")[["article", "reason"]].to_dict("index")

    decisions = {}
    for pid, grp in combined_df.groupby("_pid"):
        eans = set(grp["_ean"].dropna())
        excl_hits = eans & excl_eans
        ok_hits = eans & ok_eans

        if excl_hits:
            reasons = []
            for e in excl_hits:
                info = ean_info.get(e)
                if info:
                    reasons.append(f"{info['article']} ({e}): {info['reason']}")
                else:
                    reasons.append(f"EAN {e}: excluded")
            decisions[pid] = {
                "decision": "Excluded", "reason": "; ".join(reasons),
                "total_variants": len(eans), "eligible_variants": len(ok_hits), "excluded_variants": len(excl_hits),
            }
        elif ok_hits:
            decisions[pid] = {
                "decision": "Included", "reason": f"{len(ok_hits)} eligible variant(s) in stock",
                "total_variants": len(eans), "eligible_variants": len(ok_hits), "excluded_variants": 0,
            }
        else:
            decisions[pid] = {
                "decision": "Excluded", "reason": "No eligible-in-stock variant found",
                "total_variants": len(eans), "eligible_variants": 0, "excluded_variants": 0,
            }
    return decisions

def process_shopee(ean_df, zip_bytes):
    combined = _read_shopee_zip(zip_bytes)
    combined["_ean"] = combined.apply(lambda r: _extract_ean(r.get("SKU"), r.get("Parent SKU")), axis=1)
    combined["_pid"] = combined["Product ID"].apply(clean_id_str)
    decisions = build_pid_decisions(combined, ean_df)
    ids = [pid for pid, d in decisions.items() if d["decision"] == "Included"]
    return ids, decisions

def process_zalora(ean_df, eligible_bytes, content_df):
    df = pd.read_excel(io.BytesIO(eligible_bytes), sheet_name="Eligible Products")
    ok = eligible_ean_set(ean_df)
    nr = no_remark_ean_set(ean_df)
    ean2art = dict(zip(content_df["EAN"].astype(str).str.strip(), content_df["Color_No"].astype(str).str.strip()))
    df["_ean"] = df["Seller SKU"].astype(str).str.strip()
    df["Article No"] = df["_ean"].map(ean2art)
    df["Voucher Eligible"] = df["_ean"].apply(lambda e: "Yes" if e in ok else ("No Remark" if e in nr else "No"))
    df.drop(columns=["_ean"], inplace=True)
    return df

def _find_col(df, *keyword_sets):
    for kws in keyword_sets:
        for c in df.columns:
            cl = str(c).strip().lower()
            if all(kw in cl for kw in kws): return c
    return None

def _autodetect_tiktok_header(raw_bytes):
    xls = pd.ExcelFile(io.BytesIO(raw_bytes))
    sheet = "Template" if "Template" in xls.sheet_names else xls.sheet_names[0]
    raw = pd.read_excel(xls, sheet_name=sheet, header=None, nrows=15)
    for i in range(len(raw)):
        row_vals = raw.iloc[i].astype(str).str.lower().tolist()
        has_sku = any("sku" in v for v in row_vals)
        has_pid = any(("product" in v and "id" in v) for v in row_vals)
        if has_sku and has_pid: return sheet, i
    return sheet, None

def process_tiktok(ean_df, tiktok_bytes):
    df = None
    try:
        df = pd.read_excel(io.BytesIO(tiktok_bytes), sheet_name="Template", header=2, skiprows=[3, 4])
    except Exception: pass

    sku_col = pid_col = None
    if df is not None:
        sku_col = _find_col(df, ("seller", "sku"), ("sku",))
        pid_col = _find_col(df, ("product", "id"))

    if sku_col is None or pid_col is None:
        sheet, header_row = _autodetect_tiktok_header(tiktok_bytes)
        if header_row is not None:
            df2 = pd.read_excel(io.BytesIO(tiktok_bytes), sheet_name=sheet, header=header_row)
            sku_col2 = _find_col(df2, ("seller", "sku"), ("sku",))
            pid_col2 = _find_col(df2, ("product", "id"))
            if sku_col2 and pid_col2: df, sku_col, pid_col = df2, sku_col2, pid_col2

    if sku_col is None or pid_col is None:
        available = list(df.columns) if df is not None else ["(could not read file)"]
        raise ValueError(f"Could not find a 'Seller SKU' and 'Product ID' column in TikTok export. Found: {available}")

    df["_ean"] = df[sku_col].apply(lambda v: _extract_ean(v, None))
    df["_pid"] = df[pid_col].apply(clean_id_str)
    decisions = build_pid_decisions(df, ean_df)
    ids = [pid for pid, d in decisions.items() if d["decision"] == "Included"]
    return ids, decisions

# ==========================================
# 🎨 STREAMLIT UI LAYOUT
# ==========================================
st.title("🚀 Marketplace Portal Engine")

mode = st.selectbox("Select Core Automation Loop", ["Standard Price Matrix Pipeline", "Cross-Platform Eligibility Framework"])

st.markdown("---")
col1, col2 = st.columns(2)

with col1:
    st.subheader("📋 Core Matrix Operational Uploads")
    tracker_file = st.file_uploader("Upload Master Tracker Matrix (.xlsx)", type=["xlsx"])
    tracker_pim, tracker_rrp, tracker_md = None, None, None
    df_tracker = None
    
    if tracker_file:
        tracker_headers, df_tracker = get_clean_headers_and_df(tracker_file)
        if df_tracker is not None:
            st.success("Headers configured from tracking structure row 3.")
            tracker_pim = st.selectbox("Map PIM ID Identification Column", [""] + tracker_headers)
            tracker_rrp = st.selectbox("Map Standard Base RRP Column", [""] + tracker_headers)
            tracker_md = st.selectbox("Map Target Special/Markdown Price Column", [""] + tracker_headers)

    sku_file = st.file_uploader("Upload Universal SKU Mapping Reference", type=["xlsx", "csv"])
    sku_sku, sku_pim = None, None
    if sku_file:
        sku_hd = list(read_full_file_standard(sku_file).columns)
        sku_sku = st.selectbox("Map Operational Seller SKU Column", [""] + sku_hd)
        sku_pim = st.selectbox("Map Cross-Reference PIM ID Column", [""] + sku_hd)

with col2:
    st.subheader("📦 Multi-Platform Stream Ingestion")
    shopee_file = st.file_uploader("Shopee Master / Archive ZIP Stream Package", type=["xlsx", "zip"])
    lazada_file = st.file_uploader("Lazada Operational Sheet Document (.xlsx)", type=["xlsx"])
    tiktok_file = st.file_uploader("TikTok Operational Channel Catalog (.xlsx)", type=["xlsx"])

# ==========================================
# ⚙️ EXECUTION MATRIX
# ==========================================
if st.button("⚡ Run Consolidated Processing Engine", type="primary", use_container_width=True):
    if df_tracker is not None and sku_file is not None:
        with st.spinner("Executing calculations across linked platform matrices..."):
            try:
                df_sku = read_full_file_standard(sku_file)
                
                # Build operational lookup structures from the tracker input
                tracker_map = {}
                rrp_map = {}
                for _, row in df_tracker.iterrows():
                    pim = row.get(tracker_pim)
                    if pd.isna(pim) or str(pim).strip() == "": continue
                    rrp = pd.to_numeric(row.get(tracker_rrp), errors='coerce') or 0
                    md = pd.to_numeric(row.get(tracker_md), errors='coerce') or 0
                    tracker_map[pim] = round(md) if md != 0 and not pd.isna(md) else round(rrp)
                    rrp_map[pim] = round(rrp)

                # Process specific eligibility targets if matching data provided
                st.info("Core pricing parameters constructed successfully.")
                
                if shopee_file and shopee_file.name.endswith('.zip'):
                    shopee_bytes = shopee_file.read()
                    ids, sh_decisions = process_shopee(df_sku, shopee_bytes)
                    st.success(f"Processed Shopee package file. Found {len(ids)} fully eligible PIDs.")
                
                if tiktok_file:
                    tk_bytes = tiktok_file.read()
                    tk_ids, tk_decisions = process_tiktok(df_sku, tk_bytes)
                    st.success(f"Processed TikTok data sheets. Found {len(tk_ids)} fully eligible platform PIDs.")
                    
                if lazada_file:
                    lz_bytes = lazada_file.read()
                    lz_skus = process_lazada(df_sku, lz_bytes)
                    st.success(f"Processed Lazada data templates. Extracted {len(lz_skus)} active item matches.")
                    
            except Exception as ex:
                st.error(f"Execution terminated due to error pipeline constraints: {ex}")
    else:
        st.warning("Ensure at least the Primary Tracker and Universal SKU maps are structurally defined.")
