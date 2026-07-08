import streamlit as st
import pandas as pd
import re
import io
import zipfile
from openpyxl import load_workbook

# ─────────────────────────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────────────────────────
st.set_page_config(page_title="PUMA Voucher & Promotional SKU Tool", page_icon="🏷️", layout="wide")

# ─────────────────────────────────────────────────────────────────
# CONSTANTS & CONFIGURATIONS
# ─────────────────────────────────────────────────────────────────

REGION_MARKETPLACES = {
    "PH": ["Lazada", "Shopee", "Zalora"],
    "MY": ["Lazada", "Shopee", "Zalora", "TikTok"],
    "SG": ["Lazada", "Shopee", "Zalora"],
}

REGION_CONFIG = {
    "PH": {
        "zecom_sheet": "PH", "zecom_read": "ph", "article_col": "PIM Article#",
        "threshold": 650, "currency": "PHP 650",
        "mp_flags": {"Lazada": "LAZADA", "Shopee": "SHOPEE", "Zalora": "ZALORA"},
        "default_excl": 71, "default_rrp": 32, "default_srp": 50, "default_launch": 0,
    },
    "MY": {
        "zecom_sheet": "MY", "zecom_read": "header3", "article_col": "Style#",
        "threshold": 39, "currency": "RM 39",
        "mp_flags": {"Lazada": "Lazada", "Shopee": "Shopee",
                     "Zalora": "Zalora MP", "TikTok": "TIKTOK"},
        "default_excl": 51, "default_rrp": 27, "default_srp": 49, "default_launch": 0,
    },
    "SG": {
        "zecom_sheet": "SG", "zecom_read": "header3", "article_col": "STYLE#",
        "threshold": 16, "currency": "SGD 16",
        "mp_flags": {"Lazada": "Lazada", "Shopee": "Shopee", "Zalora": "Zalora"},
        "default_excl": 52, "default_rrp": 26, "default_srp": 50, "default_launch": 0,
    },
}

PID_MARKETPLACES = {"Shopee", "TikTok"}

# ─────────────────────────────────────────────────────────────────
# DATA NORMALIZATION HELPERS
# ─────────────────────────────────────────────────────────────────

def normalize_sku(sku):
    """Brought from G-Script: Standardizes formatting strings for robust mapping keys."""
    if pd.isna(sku) or not str(sku).strip(): 
        return ""
    return str(sku).strip().replace("-", "_").lower()


def clean_id_str(val):
    if pd.isna(val):
        return None
    if isinstance(val, float):
        return str(int(val)) if val.is_integer() else str(val)
    s = str(val).strip()
    if re.match(r"^-?\d+\.0+$", s):
        s = s.split(".")[0]
    return s


def excel_col_letter(idx: int) -> str:
    idx += 1
    letters = ""
    while idx > 0:
        idx, rem = divmod(idx - 1, 26)
        letters = chr(65 + rem) + letters
    return letters

# ─────────────────────────────────────────────────────────────────
# FILE DECODING & VALIDATION
# ─────────────────────────────────────────────────────────────────

def validate_zecom_region(file_bytes: bytes, selected_region: str):
    try:
        wb = load_workbook(io.BytesIO(file_bytes), read_only=True)
        sheets = wb.sheetnames
        wb.close()
    except Exception as e:
        return False, f"Cannot read ZeCom file: {e}"
    
    # Backward support for both original Zecom tracking or G-Script structural sheets
    required_sheets = ["Tracker", "SKU_Map"] if "Tracker" in sheets else [selected_region]
    
    if "Tracker" in sheets:
        return True, "Structure: G-Script Format Detected"
        
    if selected_region == "PH":
        if "PH" not in sheets:
            extra = " (looks like MY/SG tracker)" if ("MY" in sheets or "SG" in sheets) else ""
            return False, f"⚠️ Wrong file — selected **PH** but 'PH' sheet not found{extra}."
    else:
        if selected_region not in sheets:
            extra = " (looks like PH tracker)" if "PH" in sheets else ""
            return False, f"⚠️ Wrong file — selected **{selected_region}** but sheet not found{extra}."
    return True, "OK"


@st.cache_data(show_spinner=False)
def read_zecom(file_bytes: bytes, region: str) -> pd.DataFrame:
    """Reads Master sheet structure - adapts seamlessly to direct or multi-tab files."""
    wb = load_workbook(io.BytesIO(file_bytes), read_only=True)
    sheets = wb.sheetnames
    wb.close()
    
    if "Tracker" in sheets:
        # Pull transactional pricing straight from your G-Script's structure
        df = pd.read_excel(io.BytesIO(file_bytes), sheet_name="Tracker")
        return df

    cfg = REGION_CONFIG[region]
    if cfg["zecom_read"] == "ph":
        df = pd.read_excel(io.BytesIO(file_bytes), sheet_name="PH", header=1)
        df.columns = df.iloc[0]
        df = df.iloc[1:].reset_index(drop=True)
    else:
        df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=cfg["zecom_sheet"], header=3)
    return df


@st.cache_data(show_spinner=False)
def parse_gscript_sku_map(file_bytes: bytes) -> tuple:
    """Parses and cross-maps the G-Script tracker data frames into operational dictionaries."""
    try:
        xl = pd.ExcelFile(io.BytesIO(file_bytes))
        if "Tracker" in xl.sheet_names and "SKU_Map" in xl.sheet_names:
            tracker_df = pd.read_excel(xl, sheet_name="Tracker")
            sku_df = pd.read_excel(xl, sheet_name="SKU_Map")
            
            tracker_map = {}
            rrp_map = {}
            
            for _, r in tracker_df.iterrows():
                pim = str(r.iloc[0]).strip()
                rrp = pd.to_numeric(r.iloc[1], errors="coerce") or 0
                md = pd.to_numeric(r.iloc[2], errors="coerce") or 0
                new_price = round(rrp) if md == 0 else round(md)
                tracker_map[pim] = new_price
                rrp_map[pim] = round(rrp)
                
            price_map = {}
            sku_to_pim = {}
            for _, r in sku_df.iterrows():
                raw_sku = r.iloc[0]
                pim = str(r.iloc[1]).strip()
                if pim in tracker_map:
                    norm_sku = normalize_sku(raw_sku)
                    price_map[norm_sku] = tracker_map[pim]
                    sku_to_pim[norm_sku] = pim
            return price_map, sku_to_pim, rrp_map
    except Exception:
        pass
    return None, None, None


# ─────────────────────────────────────────────────────────────────
# AUTO-DETECTION COLUMN HELPERS
# ─────────────────────────────────────────────────────────────────

_EXCL_RE = re.compile(r'open for all|exclude|vc only|vc max|vc -|shopee exclusive|platform vc', re.I)

def col_options(df):
    return [f"{excel_col_letter(i)}: {col}" for i, col in enumerate(df.columns)]

def _col_score_excl(series):
    vals = series.dropna().astype(str).str.strip()
    vals = vals[vals != ""]
    if vals.empty: return 0
    return int(vals.apply(lambda v: bool(_EXCL_RE.search(v))).sum())

def _col_by_name(df, keywords):
    for kw in keywords:
        for i, col in enumerate(df.columns):
            if kw.lower() in str(col).lower():
                return i
    return None

def guess_excl_idx(df, fallback):
    best_col, best_score = fallback, 0
    for i in range(len(df.columns)):
        s = _col_score_excl(df.iloc[:, i])
        if s > best_score:
            best_score, best_col = s, i
    return min(best_col, len(df.columns)-1)

def guess_rrp_idx(df, fallback):
    idx = _col_by_name(df, ["rrp"])
    return min(idx if idx is not None else fallback, len(df.columns)-1)

def guess_srp_idx(df, fallback):
    for kw in ["srp ao", "ec srp", "srp"]:
        for i, col in enumerate(df.columns):
            if kw in str(col).lower() and "rrp" not in str(col).lower():
                return min(i, len(df.columns)-1)
    return min(fallback, len(df.columns)-1)

def guess_launch_idx(df, fallback):
    for kw in ["launch date", "launch_date", "launchdate", "launch"]:
        for i, col in enumerate(df.columns):
            if kw in str(col).lower():
                return min(i, len(df.columns)-1)
    return min(fallback, len(df.columns)-1)

def sample_vals(df, col_idx, n=6):
    if col_idx >= len(df.columns): return "(no values)"
    vals = df.iloc[:, col_idx].dropna().unique()[:n]
    return ", ".join(str(v) for v in vals) if len(vals) else "(no values)"

def get_unique_remarks(df, col_idx):
    if col_idx >= len(df.columns): return []
    vals = df.iloc[:, col_idx].dropna().astype(str).str.strip()
    vals = vals[(vals != "") & (vals.str.lower() != "nan")]
    if vals.empty: return []
    return vals.value_counts().index.tolist()

# ─────────────────────────────────────────────────────────────────
# INVENTORY INGESTION
# ─────────────────────────────────────────────────────────────────

def _normalize_inv(df, ean_c, stock_c):
    ec = next((c for c in ean_c   if c in df.columns), None)
    sc = next((c for c in stock_c if c in df.columns), None)
    if ec is None or sc is None: return None
    out = df[[ec, sc]].copy()
    out.columns = ["EAN"] + ["Stock"]
    out["EAN"]   = out["EAN"].astype(str).str.strip()
    out["Stock"] = pd.to_numeric(out["Stock"], errors="coerce").fillna(0)
    return out[out["EAN"].str.match(r"^\d{13}$")]

def read_inventory(file_bytes, filename):
    ec = ["EAN", "Sku", "PROD_CODE", "SellerSku"]
    sc = ["Avail_Qty", "QtyAvailable", "QTY", "Quantity"]
    if filename.lower().endswith(".csv"):
        return _normalize_inv(pd.read_csv(io.BytesIO(file_bytes)), ec, sc)
    r = _normalize_inv(pd.read_excel(io.BytesIO(file_bytes)), ec, sc)
    return r if r is not None else _normalize_inv(pd.read_excel(io.BytesIO(file_bytes), header=4), ec, sc)

def parse_special_articles(text_input: str, file_bytes: bytes = None, filename: str = None) -> set:
    articles = set()
    if text_input and text_input.strip():
        for tok in re.split(r"[,\n\r\t]+", text_input.strip()):
            tok = tok.strip()
            if tok: articles.add(tok)
    if file_bytes and filename:
        try:
            fdf = pd.read_csv(io.BytesIO(file_bytes), header=None) if filename.lower().endswith(".csv") else pd.read_excel(io.BytesIO(file_bytes), header=None)
            first_col = fdf.iloc[:, 0].dropna().astype(str).str.strip()
            for v in first_col:
                if v and v.lower() not in ("article", "article no", "article number", "sku", "nan"):
                    articles.add(v)
        except Exception:
            pass
    return articles

# ─────────────────────────────────────────────────────────────────
# CALCULATION PIPELINES & LOGIC
# ─────────────────────────────────────────────────────────────────

def classify_launch_date(raw_val, today: pd.Timestamp):
    dt = pd.to_datetime(raw_val, errors="coerce")
    if pd.isna(dt): return False, "Blank", "No Launch Date Set"
    if dt.year < 2000: return False, dt.strftime("%d-%m-%Y"), "No Launch Date Set (default date)"
    disp = dt.strftime("%d-%m-%Y")
    if dt.normalize() > today: return False, disp, f"Future Launch ({disp})"
    return True, disp, ""


def classify_row(article, mp_status, status_ok, launch_ok, launch_reason, price_ok, remark,
                 eligible_remarks, include_no_remark, special_articles):
    if article in special_articles: return "ineligible", "Special Article Exclusion"
    if not status_ok:
        disp = mp_status if mp_status not in ("", "NAN", None) else "BLANK"
        return "ineligible", f"MP Status = {disp}"
    if not launch_ok: return "ineligible", launch_reason
    if not price_ok: return "ineligible", "Price below threshold (RRP/SRP)"
    if pd.isna(remark) or str(remark).strip() == "" or str(remark).strip().lower() == "nan":
        if include_no_remark: return "eligible", ""
        return "no_remark", "No remark (not included)"
    r = str(remark).strip()
    if r in eligible_remarks: return "eligible", ""
    return "ineligible", f'Remark not selected ("{r}")'


def process_zecom(zecom_df, region, marketplace, excl_idx, rrp_idx, srp_idx, launch_idx,
                  eligible_remarks: set, include_no_remark: bool,
                  special_articles: set, apply_launch_filter: bool = True) -> pd.DataFrame:
    cfg = REGION_CONFIG[region]
    df  = zecom_df.copy()

    # Fallbacks if columns are offset or missing in customized structures
    if excl_idx >= len(df.columns) or rrp_idx >= len(df.columns) or srp_idx >= len(df.columns):
        # Create minimal fallback mapping shell
        return pd.DataFrame(columns=["article", "mp_status", "rrp", "srp", "remark", "launch_date", "status", "reason"])

    mp_col = cfg["mp_flags"].get(marketplace)
    if mp_col and mp_col in df.columns:
        mp_status_disp = df[mp_col].fillna("").astype(str).str.strip().str.upper()
        status_ok = mp_status_disp == "YES"
    else:
        mp_status_disp = pd.Series(["N/A"] * len(df), index=df.index)
        status_ok = pd.Series([True] * len(df), index=df.index)

    threshold = cfg.get("threshold_overrides", {}).get(marketplace, cfg["threshold"])
    rrp = pd.to_numeric(df.iloc[:, rrp_idx], errors="coerce")
    srp = pd.to_numeric(df.iloc[:, srp_idx], errors="coerce")
    srp_ok   = srp.isna() | (srp == 0) | (srp >= threshold)
    price_ok = (rrp > threshold) & srp_ok

    today = pd.Timestamp.now().normalize()
    launch_ok_list, launch_disp_list, launch_reason_list = [], [], []
    
    launch_series = df.iloc[:, launch_idx] if launch_idx < len(df.columns) else pd.Series([None]*len(df))
    for raw_val in launch_series:
        if apply_launch_filter:
            ok, disp, reason = classify_launch_date(raw_val, today)
        else:
            _, disp, _ = classify_launch_date(raw_val, today)
            ok, reason = True, ""
        launch_ok_list.append(ok); launch_disp_list.append(disp); launch_reason_list.append(reason)

    remark_vals  = df.iloc[:, excl_idx]
    
    # Try dynamic resolution for structural column names
    art_col_name = cfg["article_col"] if cfg["article_col"] in df.columns else df.columns[0]
    article_vals = df[art_col_name].astype(str).str.strip()

    work = pd.DataFrame({
        "article":      article_vals.values,
        "mp_status":    mp_status_disp.values,
        "status_ok":    status_ok.values,
        "launch_date":  launch_disp_list,
        "launch_ok":    launch_ok_list,
        "launch_reason":launch_reason_list,
        "rrp":          rrp.values,
        "srp":          srp.values,
        "price_ok":     price_ok.values,
        "remark":       remark_vals.values,
    })

    work = work[work["article"].str.match(r"^[\w_\-]+$", na=False)]
    work = work[work["article"].str.lower() != "nan"]
    work = work.drop_duplicates(subset=["article"])

    statuses, reasons = [], []
    for row in work.itertuples(index=False):
        s, r = classify_row(row.article, row.mp_status, row.status_ok,
                            row.launch_ok, row.launch_reason, row.price_ok,
                            row.remark, eligible_remarks, include_no_remark, special_articles)
        statuses.append(s); reasons.append(r)
    work["status"] = statuses
    work["reason"] = reasons

    return work[["article", "mp_status", "rrp", "srp", "remark", "launch_date", "status", "reason"]]


def map_to_eans(article_df, content_df, inventory_df):
    merged = article_df.merge(content_df.rename(columns={"Color_No": "article"}), on="article", how="inner")
    merged["EAN"] = merged["EAN"].astype(str).str.strip()
    merged = merged.merge(inventory_df.rename(columns={"Stock": "stock_qty"}), on="EAN", how="left")
    merged["stock_qty"] = merged["stock_qty"].fillna(0)
    merged["has_stock"] = merged["stock_qty"] > 0
    return merged[["article", "EAN", "mp_status", "rrp", "srp", "remark", "launch_date", "status", "reason", "stock_qty", "has_stock"]]

def eligible_ean_set(df): return set(df[(df["status"] == "eligible") & df["has_stock"]]["EAN"])
def excluded_ean_set(df): return set(df[df["status"] == "ineligible"]["EAN"])
def no_remark_ean_set(df): return set(df[df["status"] == "no_remark"]["EAN"])

# ─────────────────────────────────────────────────────────────────
# MARKETPLACE ADAPTER ENGINES
# ─────────────────────────────────────────────────────────────────

def process_lazada(ean_df, lazada_bytes):
    df   = pd.read_excel(io.BytesIO(lazada_bytes), sheet_name="template", header=0)
    data = df.iloc[3:].copy(); data.columns = df.columns
    active = data[data["status"].astype(str).str.lower() == "active"].copy()
    active["_ean"] = active["SellerSKU"].astype(str).str.strip()
    ok = eligible_ean_set(ean_df)
    skus = active[active["_ean"].isin(ok)]["Shop SKU"].dropna().apply(clean_id_str)
    return skus.dropna().unique().tolist()


def _read_shopee_zip(zip_bytes):
    dfs = []
    with zipfile.ZipFile(io.BytesIO(zip_bytes)) as zf:
        names = sorted(n for n in zf.namelist() if n.endswith(".xlsx"))
        bar = st.progress(0, text="Reading Shopee export packages...")
        for i, name in enumerate(names):
            with zf.open(name) as f:
                dfs.append(pd.read_excel(f, engine="calamine", header=2, skiprows=[3, 4]))
            bar.progress((i + 1) / len(names), text=f"Parsing Shopee bundle {i+1}/{len(names)}...")
        bar.empty()
    return pd.concat(dfs, ignore_index=True)


def _extract_ean(sku_val, parent_val):
    for v in (sku_val, parent_val):
        if pd.notna(v):
            try:
                s = str(int(float(v)))
                if re.match(r"^\d{13}$", s): return s
            except (ValueError, TypeError):
                s = str(v).strip()
                if re.match(r"^\d{13}$", s): return s
    return None


def build_pid_decisions(combined_df, ean_df):
    ok_eans   = eligible_ean_set(ean_df)
    excl_eans = excluded_ean_set(ean_df)
    ean_info  = ean_df.drop_duplicates(subset=["EAN"]).set_index("EAN")[["article", "reason"]].to_dict("index")

    decisions = {}
    for pid, grp in combined_df.groupby("_pid"):
        eans = set(grp["_ean"].dropna())
        excl_hits = eans & excl_eans
        ok_hits   = eans & ok_eans

        if excl_hits:
            reasons = []
            for e in excl_hits:
                info = ean_info.get(e)
                reasons.append(f"{info['article']} ({e}): {info['reason']}" if info else f"EAN {e}: excluded")
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
    ok = eligible_ean_set(ean_df); nr = no_remark_ean_set(ean_df)
    ean2art = dict(zip(content_df["EAN"].astype(str).str.strip(), content_df["Color_No"].astype(str).str.strip()))
    df["_ean"]       = df["Seller SKU"].astype(str).str.strip()
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
    except Exception:
        pass

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
        raise ValueError(f"Could not find dynamic 'Seller SKU' and 'Product ID' columns. Columns: {available}")

    df["_ean"] = df[sku_col].apply(lambda v: _extract_ean(v, None))
    df["_pid"] = df[pid_col].apply(clean_id_str)
    decisions = build_pid_decisions(df, ean_df)
    ids = [pid for pid, d in decisions.items() if d["decision"] == "Included"]
    return ids, decisions

# ─────────────────────────────────────────────────────────────────
# ARTIFACT & ARCHIVE EXPORTERS
# ─────────────────────────────────────────────────────────────────

def _to_excel_multi(sheets: dict):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        for name, d in sheets.items():
            d.to_excel(w, index=False, sheet_name=name[:31])
    return buf.getvalue()

def make_lazada_output(ids): return _to_excel_multi({"Sheet1": pd.DataFrame({"SHOP SKU": ids})})
def make_shopee_output(ids): return _to_excel_multi({"Sheet1": pd.DataFrame({"Product ID": ids})})
def make_zalora_output(ann): return _to_excel_multi({"Eligible Products": ann})


def make_summary_excel(ean_df, region, marketplace, pct, voucher_type, pid_decisions=None, column_labels=None):
    detail = ean_df.copy()
    detail["status"] = detail["status"].map({"eligible": "Eligible", "ineligible": "Ineligible", "no_remark": "No Remark"})
    detail = detail.rename(columns={
        "article": "Article", "mp_status": "MP Status", "rrp": "RRP", "srp": "SRP",
        "remark": "Remark", "launch_date": "Launch Date", "status": "Status",
        "reason": "Exclusion Reason", "stock_qty": "Stock Qty", "has_stock": "In Stock",
    })
    detail.insert(0, "Region", region)
    detail.insert(1, "Marketplace", marketplace)
    detail.insert(2, "Voucher %", pct)
    detail.insert(3, "Voucher Type", voucher_type)
    cols = ["Region", "Marketplace", "Voucher %", "Voucher Type", "Article", "EAN",
            "MP Status", "RRP", "SRP", "Remark", "Launch Date", "Status", "Exclusion Reason", "Stock Qty", "In Stock"]
    detail = detail[[c for c in cols if c in detail.columns]]

    sheets = {"Article_EAN_Detail": detail}
    
    # Structural calculations
    stats_rows = [
        ("Region", region), ("Marketplace", marketplace),
        ("Voucher %", pct), ("Voucher Type", voucher_type),
        ("Total Rows Processed", len(detail)),
        ("Eligible Count", int((detail["Status"] == "Eligible").sum())),
        ("Ineligible Count", int((detail["Status"] == "Ineligible").sum())),
    ]

    if column_labels:
        for k, v in column_labels.items():
            stats_rows.append((f"Mapped Header Column ({k})", v))

    if pid_decisions:
        pid_rows = [{"Product ID": pid, "Decision": d["decision"], "Total Variants": d["total_variants"], "Eligible Variants": d["eligible_variants"], "Excluded Variants": d["excluded_variants"], "Reason": d["reason"]} for pid, d in pid_decisions.items()]
        sheets["Product_ID_Summary"] = pd.DataFrame(pid_rows)

    sheets["Summary_Stats"] = pd.DataFrame(stats_rows, columns=["Metric", "Value"])
    return _to_excel_multi(sheets)

# ─────────────────────────────────────────────────────────────────
# CORE STREAMLIT USER INTERFACE
# ─────────────────────────────────────────────────────────────────

def main():
    st.title("🏷️ PUMA Voucher & Promotional SKU Tool")
    st.caption("Generate marketplace-ready voucher tracking structures from sheet configurations.")

    # ── ① REGION & MARKETPLACE ───────────────────────────────────
    st.markdown("---")
    st.subheader("① Region & Marketplace")
    c1, c2 = st.columns(2)
    with c1: region = st.selectbox("Region", ["PH", "MY", "SG"])
    with c2: marketplace = st.selectbox("Marketplace", REGION_MARKETPLACES[region])

    # ── ② UPLOAD FILES ────────────────────────────────────────────
    st.markdown("---")
    st.subheader("② Upload Files")
    cf1, cf2 = st.columns(2)
    with cf1:
        st.markdown("**Core Workbook Configuration**")
        zecom_file = st.file_uploader("ZeCom Tracker or G-Script Master Book (.xlsx)", type=["xlsx"], key="zecom")
        content_file = st.file_uploader("Content Configuration Reference File (.xlsx)", type=["xlsx"], key="content")
        inv_file = st.file_uploader(f"Real-time Inventory Sheet (.xlsx, .csv) [{region}]", type=["xlsx", "csv"], key="inv")
    with cf2:
        st.markdown(f"**Target {marketplace} Vendor Export System**")
        if   marketplace == "Lazada": mp_file = st.file_uploader("Lazada Product Export Sheet (.xlsx)", type=["xlsx"], key="mp")
        elif marketplace == "Shopee": mp_file = st.file_uploader("Shopee Export Bundle Package (.zip)", type=["zip"],  key="mp")
        elif marketplace == "Zalora": mp_file = st.file_uploader("Zalora EligibleProducts Master Sheet (.xlsx)", type=["xlsx"], key="mp")
        elif marketplace == "TikTok": mp_file = st.file_uploader("TikTok Shop Center Inventory Export (.xlsx)", type=["xlsx"], key="mp")
        else: mp_file = None

    # ── ②b SPECIAL ARTICLE EXCLUSION ─────────────────────────────
    st.markdown("---")
    st.subheader("②b Special Article Hard Exclusions")
    sa1, sa2 = st.columns(2)
    with sa1: special_text = st.text_area("Input specific global exclusion articles (comma/newline separated)", placeholder="521351_01\n521352_02", height=100)
    with sa2: special_file = st.file_uploader("Or append via simple text file upload", type=["csv", "xlsx"], key="special_file")

    special_articles = parse_special_articles(special_text, special_file.getvalue() if special_file else None, special_file.name if special_file else None)
    if special_articles: st.info(f"🚫 {len(special_articles)} system rules forced to global ineligibility lists.")

    # ── ③ ZECOM COLUMNS + REMARKS + VOUCHER ──────────────────────
    excl_idx = rrp_idx = srp_idx = launch_idx = zecom_df = None
    apply_launch_filter = True
    voucher_configs = []
    voucher_type = "Regular VC"

    if zecom_file:
        file_bytes = zecom_file.getvalue()
        ok_region, err_msg = validate_zecom_region(file_bytes, region)
        if not ok_region:
            st.error(err_msg); st.stop()

        st.markdown("---")
        st.subheader("③ Structure Column Strategy Configuration")
        
        # Intercept G-Script layout profile mapping
        gscript_price_map, gscript_sku_to_pim, gscript_rrp_map = parse_gscript_sku_map(file_bytes)
        
        with st.spinner("Reading primary structural dataset matrices..."):
            zecom_df = read_zecom(file_bytes, region)

        if gscript_price_map:
            st.success("🤖 G-Script 'Tracker' + 'SKU_Map' pattern detected! Automatically resolving hierarchies.")
            excl_idx, rrp_idx, srp_idx, launch_idx = 2, 1, 2, 0 # structural assignments
        else:
            cfg = REGION_CONFIG[region]
            opts = col_options(zecom_df)
            
            d_excl = guess_excl_idx(zecom_df, cfg["default_excl"])
            d_rrp  = guess_rrp_idx(zecom_df, cfg["default_rrp"])
            d_srp  = guess_srp_idx(zecom_df, cfg["default_srp"])
            d_launch = guess_launch_idx(zecom_df, cfg["default_launch"])

            sc1, sc2, sc3, sc4 = st.columns(4)
            with sc1:
                excl_sel = st.selectbox("📋 Campaign Selection Mapping Target", opts, index=d_excl)
                excl_idx = opts.index(excl_sel)
                st.caption(f"Sample: `{sample_vals(zecom_df, excl_idx)}`")
            with sc2:
                rrp_sel = st.selectbox("💰 Target RRP Verification Header", opts, index=d_rrp)
                rrp_idx = opts.index(rrp_sel)
                st.caption(f"Sample: `{sample_vals(zecom_df, rrp_idx)}`")
            with sc3:
                srp_sel = st.selectbox("🏷️ Targeted SRP Promotion Variant Line", opts, index=d_srp)
                srp_idx = opts.index(srp_sel)
                st.caption(f"Sample: `{sample_vals(zecom_df, srp_idx)}`")
            with sc4:
                launch_sel = st.selectbox("📅 Internal System Launch Window", opts, index=d_launch)
                launch_idx = opts.index(launch_sel)
                st.caption(f"Sample: `{sample_vals(zecom_df, launch_idx)}`")

        apply_launch_filter = st.checkbox("🚫 Reject unlaunched products (recommended safely checking dates)", value=True)

        # ── ④ MULTI-VOUCHER CONFIGURATION GRID ──
        st.markdown("---")
        st.subheader("④ Distribution Run Strategy Matrix")
        unique_remarks = get_unique_remarks(zecom_df, excl_idx) if not gscript_price_map else ["Promotion Price Rule Active"]
        voucher_type = st.radio("Run Target Allocation Type Strategy Profile", ["Regular VC", "Bundle Discount"], horizontal=True)

        if "voucher_row_ids" not in st.session_state: st.session_state.voucher_row_ids = [0]
        if "voucher_row_counter" not in st.session_state: st.session_state.voucher_row_counter = 1

        def _render_voucher_row(rid, position):
            st.markdown(f"**Voucher Configuration #{position}**")
            rcol1, rcol2, rcol3 = st.columns([1, 3, 0.6])
            with rcol1:
                pct_key = f"vc_pct_{rid}"
                pct_raw = st.text_input("Discount % Value Allocation", value=st.session_state.get(pct_key, "10"), key=pct_key)
                pct_clean = pct_raw.strip().replace("%", "")
                pct_val = int(pct_clean) if pct_clean.isdigit() else None
            with rcol2:
                selected = st.multiselect("Allowed Eligibility Rules Contexts Dropdown Selection", options=unique_remarks, default=unique_remarks if gscript_price_map else [], key=f"vc_remarks_{rid}")
                include_nr = st.checkbox("Accept blank / unassigned fields directly as promotion ready", value=False, key=f"vc_nr_{rid}")
            with rcol3:
                st.markdown("&nbsp;")
                remove_clicked = st.button("🗑️", key=f"vc_rm_{rid}") if len(st.session_state.voucher_row_ids) > 1 else False
            return {"rid": rid, "pct": pct_val, "remarks": set(selected), "include_no_remark": include_nr, "remove": remove_clicked}

        voucher_configs = []
        to_remove = None
        for pos, rid in enumerate(list(st.session_state.voucher_row_ids), start=1):
            row = _render_voucher_row(rid, pos)
            voucher_configs.append(row)
            if row["remove"]: to_remove = rid

        if to_remove is not None:
            st.session_state.voucher_row_ids.remove(to_remove)
            st.rerun()

        if st.button("➕ Add Variant Parameter Pipeline Row"):
            st.session_state.voucher_row_ids.append(st.session_state.voucher_row_counter)
            st.session_state.voucher_row_counter += 1
            st.rerun()

    # ── ⑤ COMPILING PIPELINE TRIGGER ──
    st.markdown("---")
    st.subheader("⑤ Execute System Compilations")
    missing = []
    if not zecom_file: missing.append("ZeCom / Master Price Document Tracker Configuration")
    if not content_file: missing.append("EAN Content Reference Index Data")
    if not inv_file: missing.append("Real-Time Snapshot Warehousing Metrics File")
    if not mp_file: missing.append(f"{marketplace} Native Structural Export Sheet Reference Target")

    if missing:
        st.info(f"Validation Dependencies Awaiting Input Injection: **{', '.join(missing)}**")
    else:
        if st.button("🚀 Process & Generate Dynamic Market Target Distributions", type="primary"):
            _run(zecom_file, content_file, inv_file, mp_file, region, marketplace, excl_idx, rrp_idx, srp_idx, launch_idx, apply_launch_filter, special_articles, voucher_type, voucher_configs)

    render_results()


# ─────────────────────────────────────────────────────────────────
# DATA PROCESSING PIPELINE execution ENGINE
# ─────────────────────────────────────────────────────────────────

def _run(zecom_file, content_file, inv_file, mp_file,
         region, marketplace, excl_idx, rrp_idx, srp_idx, launch_idx, apply_launch_filter,
         special_articles, voucher_type, voucher_configs):

    with st.status("Initializing algorithmic mapping transforms...", expanded=True) as status:
        file_bytes = zecom_file.getvalue()
        g_price_map, g_sku_to_pim, g_rrp_map = parse_gscript_sku_map(file_bytes)
        
        zecom_df = read_zecom(file_bytes, region)
        column_labels = {
            "Campaign Exclusions": f"{zecom_df.columns[excl_idx]}" if excl_idx < len(zecom_df.columns) else "Structural Vector",
            "Mapped Base Reference RRP": f"{zecom_df.columns[rrp_idx]}" if rrp_idx < len(zecom_df.columns) else "Structural Vector"
        }

        content_df = pd.read_excel(io.BytesIO(content_file.getvalue()), sheet_name="content")[["Color_No", "EAN"]].dropna()
        content_df["EAN"] = content_df["EAN"].astype(str).str.strip()
        content_df["Color_No"] = content_df["Color_No"].astype(str).str.strip()

        inv_df = read_inventory(inv_file.getvalue(), inv_file.name)
        all_outputs = []

        for vi, row in enumerate(voucher_configs, start=1):
            pct = row["pct"]
            
            if g_price_map:
                # Execution if Apps Script Data Core Structure is found
                art = pd.DataFrame({
                    "article": list(g_price_map.keys()), "mp_status": "YES", "rrp": [g_rrp_map.get(p, 0) for p in g_sku_to_pim.values()],
                    "srp": list(g_price_map.values()), "remark": "Promotion Price Rule Active", "launch_date": "01-01-2025", "status": "eligible", "reason": ""
                })
            else:
                art = process_zecom(zecom_df, region, marketplace, excl_idx, rrp_idx, srp_idx, launch_idx, row["remarks"], row["include_no_remark"], special_articles, apply_launch_filter)

            ean_df = map_to_eans(art, content_df, inv_df)
            n_ok = len(eligible_ean_set(ean_df))

            result = None
            pid_decisions = None

            if n_ok > 0:
                if marketplace == "Lazada":
                    result = {"mp": "Lazada", "ids": process_lazada(ean_df, mp_file.getvalue())}
                elif marketplace == "Shopee":
                    ids, pid_decisions = process_shopee(ean_df, mp_file.getvalue())
                    result = {"mp": "Shopee", "ids": ids}
                elif marketplace == "Zalora":
                    ann = process_zalora(ean_df, mp_file.getvalue(), content_df)
                    result = {"mp": "Zalora", "ann": ann, "yes_count": (ann["Voucher Eligible"] == "Yes").sum()}
                elif marketplace == "TikTok":
                    ids, pid_decisions = process_tiktok(ean_df, mp_file.getvalue())
                    result = {"mp": "TikTok", "ids": ids}

            summary_bytes = make_summary_excel(ean_df, region, marketplace, pct, voucher_type, pid_decisions, column_labels)
            all_outputs.append({"pct": pct, "result": result, "summary_bytes": summary_bytes, "pid_decisions": pid_decisions})

        status.update(label="Compilation execution finalized successfully!", state="complete")

    st.session_state["last_run"] = {
        "all_outputs": all_outputs, "region": region, "marketplace": marketplace,
        "voucher_type": voucher_type, "generated_at": pd.Timestamp.now().strftime("%Y%m%d_%H%M%S")
    }


def render_results():
    last = st.session_state.get("last_run")
    if not last: return

    all_outputs = last["all_outputs"]
    region = last["region"]
    marketplace = last["marketplace"]
    voucher_type = last["voucher_type"]
    today = pd.Timestamp.now().strftime("%Y%m%d")
    vt_short = "Bundle" if voucher_type == "Bundle Discount" else "VC"

    st.markdown("---")
    hcol1, hcol2 = st.columns([5, 1])
    with hcol1:
        st.subheader("⑥ Export Structured Outputs")
    with hcol2:
        if st.button("🧹 Clear Output Matrices"):
            del st.session_state["last_run"]
            st.rerun()

    for out in all_outputs:
        pct = out["pct"]
        result = out["result"]
        summary_bytes = out["summary_bytes"]

        st.markdown(f"#### Runtime Vector: Matrix Target Parameter Allocation ({pct}% {voucher_type})")
        d1, d2 = st.columns(2)

        with d1:
            st.markdown("**Clean Direct Marketplace Output**")
            if result:
                mp = result["mp"]
                fname = f"{mp}_{region}_{pct}pct_{vt_short}_{today}.xlsx"
                if mp == "Zalora":
                    st.metric(f"{mp} Target Total Matrix", f"{result['yes_count']} Registered SKUs")
                    data = make_zalora_output(result["ann"])
                else:
                    label = "Shop SKUs" if mp == "Lazada" else "Product IDs"
                    st.metric(f"{mp} Target Total Match Engine", f"{len(result['ids'])} {label}")
                    data = make_lazada_output(result["ids"]) if mp == "Lazada" else make_shopee_output(result["ids"])
                
                st.download_button(f"⬇️ Export Consolidated Deployment File ({fname})", data=data, file_name=fname, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key=f"dl_main_{pct}")
            else:
                st.info("No distribution outputs triggered due to zero remaining viable variants.")

        with d2:
            st.markdown("**Unified System Validation Audit Trail Report**")
            summary_fname = f"QC_Verification_Summary_{marketplace}_{region}_{pct}pct_{vt_short}_{today}.xlsx"
            st.download_button(f"⬇️ Export Evaluation Log Package ({summary_fname})", data=summary_bytes, file_name=summary_fname, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key=f"dl_summary_{pct}")
        st.markdown("---")


if __name__ == "__main__":
    main()
