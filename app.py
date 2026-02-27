import io
import re
import numpy as np
import pandas as pd
import streamlit as st

st.set_page_config(page_title="MAWB Audit Analyzer", layout="wide")
st.title("MAWB Audit Analyzer (Billing-only)")
st.caption(
    "Upload Billing charges export + optional MAWB→ETA mapping file. "
    "Outputs Excel with Analysis Summary (embedded dashboard), Risk Summary, Overlap, "
    "and Destination added to ALL MAWB-containing tabs."
)

# =========================
# Config
# =========================
STRUCTURAL_CLIENTS = {"PROCARESX"}  # excluded from ALL detail pages/exports; only totals shown
TLMF_CODE = "TLMF"

# =========================
# Helpers
# =========================
def safe_numeric(s: pd.Series) -> pd.Series:
    return pd.to_numeric(s, errors="coerce").fillna(0)

def norm_colname(s: str) -> str:
    return re.sub(r"[\s_\-]+", "", str(s).strip().lower())

def find_first_col(df: pd.DataFrame, candidates: list[str]) -> str:
    mapping = {norm_colname(c): c for c in df.columns.astype(str)}
    for cand in candidates:
        key = norm_colname(cand)
        if key in mapping:
            return mapping[key]
    return ""

def find_sheet_with_required_cols(xls: pd.ExcelFile, required_candidates: dict) -> str:
    for sh in xls.sheet_names:
        try:
            tmp = pd.read_excel(xls, sheet_name=sh, nrows=60)
        except Exception:
            continue
        ok = True
        for _, cand_list in required_candidates.items():
            if not find_first_col(tmp, cand_list):
                ok = False
                break
        if ok:
            return sh
    return ""

def clean_eta_series(s: pd.Series) -> pd.Series:
    s = s.astype(str).fillna("").str.strip()
    s = s.str.replace(r"(?i)^\s*eta\s*[:\-]\s*", "", regex=True)
    s = s.str.replace(r"\s+", " ", regex=True)

    yyyymmdd = s.str.match(r"^\d{8}$")
    s2 = s.copy()
    if yyyymmdd.any():
        parsed = pd.to_datetime(s.loc[yyyymmdd], format="%Y%m%d", errors="coerce")
        s2.loc[yyyymmdd] = parsed.astype("datetime64[ns]").astype(str)

    dt1 = pd.to_datetime(s2, errors="coerce", infer_datetime_format=True)
    mask = dt1.isna() & s2.ne("")
    if mask.any():
        dt2 = pd.to_datetime(s2[mask], errors="coerce", dayfirst=True, infer_datetime_format=True)
        dt1.loc[mask] = dt2
    return dt1.dt.normalize()

def normalize_mawb(x: str) -> str:
    if x is None:
        return ""
    s = str(x).strip().upper()
    if not s or s in {"NAN", "NONE"}:
        return ""
    s_alnum = re.sub(r"[^0-9A-Z]", "", s)
    if s_alnum.isdigit() and len(s_alnum) == 11:
        return f"{s_alnum[:3]}-{s_alnum[3:]}"
    if s_alnum.isdigit() and len(s_alnum) == 12:
        s11 = s_alnum[-11:]
        if len(s11) == 11:
            return f"{s11[:3]}-{s11[3:]}"
        return s_alnum
    if "-" in s and len(s.split("-")[0]) == 3:
        return s
    return s_alnum or s

def parse_mawb_list(text: str) -> list[str]:
    if not text or not str(text).strip():
        return []
    tokens = re.split(r"[,\s]+", str(text).strip())
    tokens = [normalize_mawb(t) for t in tokens if str(t).strip()]
    tokens = [t for t in tokens if t]
    return sorted(set(tokens))

def to_date_only(df_in: pd.DataFrame, cols: list[str]) -> pd.DataFrame:
    df_out = df_in.copy()
    for c in cols:
        if c in df_out.columns:
            df_out[c] = pd.to_datetime(df_out[c], errors="coerce").dt.date
    return df_out

def ratio_or_nan(numer: pd.Series, denom: pd.Series) -> pd.Series:
    numer = pd.to_numeric(numer, errors="coerce")
    denom = pd.to_numeric(denom, errors="coerce").fillna(0)
    out = pd.Series(np.nan, index=denom.index, dtype="float64")
    mask = denom != 0
    out.loc[mask] = (numer.loc[mask] / denom.loc[mask]).astype("float64")
    return out

def add_profit_cols(df_in: pd.DataFrame, ar_col: str, ap_col: str,
                    profit_col="Profit", pm_col="Profit Margin %") -> pd.DataFrame:
    df = df_in.copy()
    df[profit_col] = df[ar_col] - df[ap_col]
    df[pm_col] = ratio_or_nan(df[profit_col], df[ar_col])
    return df

def first_non_empty(series: pd.Series) -> str:
    s = series.astype(str).replace(["nan", "None"], "").str.strip()
    s = s[s.ne("")]
    return s.iloc[0] if len(s) else ""

def format_pct(x):
    try:
        if x is None or pd.isna(x):
            return ""
        return f"{float(x) * 100:.2f}%"
    except Exception:
        return ""

# ✅ 核心：把 Destination 补到“任何包含 MAWB 的 tab”
def attach_destination_if_mawb(df_tab: pd.DataFrame, mawb_to_dest: pd.DataFrame) -> pd.DataFrame:
    if df_tab is None or df_tab.empty:
        return df_tab
    if "MAWB" not in df_tab.columns:
        return df_tab
    out = df_tab.copy()
    if "Destination" in out.columns:
        # 已经有了就不重复 merge
        return out
    out = out.merge(mawb_to_dest, on="MAWB", how="left")
    out["Destination"] = out["Destination"].fillna("")
    # 为了更直观，把 Destination 放到 MAWB 后面
    cols = list(out.columns)
    mawb_idx = cols.index("MAWB")
    cols.remove("Destination")
    cols.insert(mawb_idx + 1, "Destination")
    return out[cols]

# Excel formatting helpers
def excel_set_percent_col(ws, col_idx: int, workbook, width: int = 16):
    fmt = workbook.add_format({"num_format": "0.00%"})
    ws.set_column(col_idx, col_idx, width, fmt)

def excel_set_number_col(ws, col_idx: int, workbook, width: int = 16):
    fmt = workbook.add_format({"num_format": "#,##0.00"})
    ws.set_column(col_idx, col_idx, width, fmt)

# =========================
# Uploaders
# =========================
billing_file = st.file_uploader("Upload Billing Charges Excel (.xlsx)", type=["xlsx"], key="billing")
eta_file = st.file_uploader("Optional: Upload MAWB→ETA mapping Excel (.xlsx)", type=["xlsx"], key="eta_mapping")

st.divider()
st.subheader("Optional Filter: Keep only specified MAWBs")
mawb_text = st.text_area(
    "Paste MAWBs here (comma / space / newline separated). Supports 99934022122 → 999-34022122. Leave blank to keep all.",
    height=140,
    placeholder="Example:\n999-34022122\n99934022133\n999 34022144"
)

# =========================
# Column candidates
# =========================
BILLING_REQUIRED = {
    "MAWB": ["MAWB", "Mawb", "Master AWB", "MasterAWB"],
    "Cost Amount": ["Cost Amount", "Cost", "AP Amount", "Total Cost", "CostAmount"],
    "Sell Amount": ["Sell Amount", "Sell", "AR Amount", "Total Sell", "SellAmount"],
}
BILLING_OPTIONAL = {
    "Client": ["Client", "Customer", "Account", "Shipper", "Bill To", "Billed To"],
    "Charge Code": ["Charge Code", "ChargeCode", "Charge", "Code"],
    "Vendor": ["Vendor", "Carrier", "Supplier"],
    "Destination": [
        "Destination", "Dest", "DST", "To", "POD",
        "Dest Airport", "Destination Airport", "Airport", "To Airport"
    ],
}
ETA_REQUIRED = {
    "MAWB": ["MAWB", "Mawb", "Master AWB", "MasterAWB"],
    "ETA": ["ETA", "Eta", "Estimated Time of Arrival", "Arrival", "Arrival Date", "ETA Date"],
}

# =========================
# Main
# =========================
if not billing_file:
    st.info("Please upload a Billing Charges Excel file to start.")
    st.stop()

try:
    # ---- Read billing charges ----
    xls = pd.ExcelFile(billing_file)
    billing_sheet = find_sheet_with_required_cols(xls, BILLING_REQUIRED)
    if not billing_sheet:
        st.error("Could not find a sheet containing MAWB + Cost Amount + Sell Amount.")
        st.stop()

    raw_df = pd.read_excel(xls, sheet_name=billing_sheet)

    mawb_col = find_first_col(raw_df, BILLING_REQUIRED["MAWB"])
    cost_col = find_first_col(raw_df, BILLING_REQUIRED["Cost Amount"])
    sell_col = find_first_col(raw_df, BILLING_REQUIRED["Sell Amount"])
    client_col = find_first_col(raw_df, BILLING_OPTIONAL["Client"])
    charge_code_col = find_first_col(raw_df, BILLING_OPTIONAL["Charge Code"])
    vendor_col = find_first_col(raw_df, BILLING_OPTIONAL["Vendor"])
    dest_col = find_first_col(raw_df, BILLING_OPTIONAL["Destination"])

    if not (mawb_col and cost_col and sell_col):
        st.error("Required columns could not be detected.")
        st.stop()

    df_all = raw_df.copy()
    df_all["MAWB"] = df_all[mawb_col].apply(normalize_mawb)
    df_all["Cost Amount"] = safe_numeric(df_all[cost_col])
    df_all["Sell Amount"] = safe_numeric(df_all[sell_col])

    df_all["Client"] = df_all[client_col].astype(str).str.strip() if client_col else "UNKNOWN"
    df_all.loc[df_all["Client"].isin(["", "nan", "None"]), "Client"] = "UNKNOWN"

    df_all["Charge Code"] = df_all[charge_code_col].astype(str).str.strip() if charge_code_col else "UNKNOWN"
    df_all.loc[df_all["Charge Code"].isin(["", "nan", "None"]), "Charge Code"] = "UNKNOWN"

    df_all["Vendor"] = df_all[vendor_col].astype(str).str.strip() if vendor_col else "UNKNOWN"
    df_all.loc[df_all["Vendor"].isin(["", "nan", "None"]), "Vendor"] = "UNKNOWN"

    if dest_col:
        df_all["Destination"] = df_all[dest_col].astype(str).str.strip()
        df_all.loc[df_all["Destination"].isin(["", "nan", "None"]), "Destination"] = ""
    else:
        df_all["Destination"] = ""

    df_all = df_all[df_all["MAWB"].ne("")].copy()

    # line-level profit/margin
    df_all = add_profit_cols(df_all, "Sell Amount", "Cost Amount")

    # ---- MAWB filter ----
    mawb_keep = parse_mawb_list(mawb_text)
    if mawb_keep:
        before_rows = len(df_all)
        before_mawb = df_all["MAWB"].nunique()
        df_all = df_all[df_all["MAWB"].isin(mawb_keep)].copy()
        after_rows = len(df_all)
        after_mawb = df_all["MAWB"].nunique()
        found_set = set(df_all["MAWB"].unique())
        mawb_not_found = sorted(set(mawb_keep) - found_set)
        st.info(f"MAWB filter applied: rows {before_rows} → {after_rows}, unique MAWB {before_mawb} → {after_mawb}.")
    else:
        mawb_not_found = []

    # ---- ETA mapping ----
    eta_map = None
    eta_parse_note = None
    if eta_file:
        xls2 = pd.ExcelFile(eta_file)
        map_sheet = find_sheet_with_required_cols(xls2, ETA_REQUIRED)
        if not map_sheet:
            st.warning("ETA mapping file uploaded, but could not find MAWB + ETA columns.")
        else:
            mdf0 = pd.read_excel(xls2, sheet_name=map_sheet)
            m_mawb = find_first_col(mdf0, ETA_REQUIRED["MAWB"])
            m_eta = find_first_col(mdf0, ETA_REQUIRED["ETA"])
            if m_mawb and m_eta:
                mdf = mdf0[[m_mawb, m_eta]].copy()
                mdf.columns = ["MAWB", "ETA"]
                mdf["MAWB"] = mdf["MAWB"].apply(normalize_mawb)
                mdf["ETA"] = clean_eta_series(mdf["ETA"])
                bad = int(mdf["ETA"].isna().sum())
                tot = int(len(mdf))
                if tot and bad:
                    eta_parse_note = f"ETA parsing note: {bad}/{tot} ETA values could not be parsed."
                eta_map = mdf.dropna(subset=["MAWB"]).groupby("MAWB", as_index=False)["ETA"].max()

    if eta_map is not None and not eta_map.empty:
        df_all = df_all.merge(eta_map, on="MAWB", how="left")
    else:
        df_all["ETA"] = pd.NaT
    df_all["ETA"] = pd.to_datetime(df_all["ETA"], errors="coerce").dt.normalize()

    # =========================
    # Structural client totals (PROCARESX only totals)
    # =========================
    structural_df = df_all[df_all["Client"].isin(STRUCTURAL_CLIENTS)].copy()
    structural_totals = None
    if not structural_df.empty:
        s_ap = float(structural_df["Cost Amount"].sum())
        s_ar = float(structural_df["Sell Amount"].sum())
        s_profit = s_ar - s_ap
        s_margin = (s_profit / s_ar) if s_ar else np.nan
        structural_totals = pd.DataFrame([{
            "Client": "PROCARESX",
            "MAWB Count": int(structural_df["MAWB"].nunique()),
            "AP": s_ap,
            "AR": s_ar,
            "Profit": s_profit,
            "Profit Margin %": s_margin,
            "Note": "Structural allocation; excluded from ALL detail pages/exports"
        }])

    # Auditable dataset
    df = df_all[~df_all["Client"].isin(STRUCTURAL_CLIENTS)].copy()

    # =========================
    # MAWB -> Destination mapping (for ALL MAWB tabs, including Raw_Billing_Enriched)
    # =========================
    mawb_to_dest = (
        df_all.groupby("MAWB", as_index=False)
              .agg(Destination=("Destination", first_non_empty))
    )

    # =========================
    # MAWB summary (auditable)
    # =========================
    mawb_summary = (
        df.groupby("MAWB", as_index=False)
          .agg(
              Client=("Client", "first"),
              Total_AP=("Cost Amount", "sum"),
              Total_AR=("Sell Amount", "sum"),
              Line_Count=("MAWB", "size"),
              ETA=("ETA", "max")
          )
    )
    mawb_summary = add_profit_cols(mawb_summary, "Total_AR", "Total_AP")
    mawb_summary["ETA Month"] = mawb_summary["ETA"].dt.to_period("M").astype(str).replace("NaT", "")

    def exception_type(r):
        if r["Total_AP"] == 0 and r["Total_AR"] == 0:
            return "Cost=Sell=0"
        if r["Total_AR"] == 0 and r["Total_AP"] > 0:
            return "Revenue=0"
        if r["Total_AP"] == 0 and r["Total_AR"] > 0:
            return "Cost=0"
        pm = r["Profit Margin %"]
        if not pd.isna(pm):
            if pm < 0.30:
                return "Margin<30%"
            if pm > 0.80:
                return "Margin>80%"
        return ""

    mawb_summary["Exception_Type"] = mawb_summary.apply(exception_type, axis=1)

    # MAWB tabs
    mawb_all = mawb_summary.copy()
    mawb_all["Classification"] = np.where(
        (mawb_all["Total_AP"] > 0) & (mawb_all["Total_AR"] > 0) &
        (~mawb_all["Profit Margin %"].isna()) &
        (mawb_all["Profit Margin %"] >= 0.30) & (mawb_all["Profit Margin %"] <= 0.80),
        "Closed", "Open"
    )

    exceptions = mawb_all[mawb_all["Classification"].eq("Open")].copy()
    margin_anomalies = mawb_all[mawb_all["Exception_Type"].isin(["Margin<30%", "Margin>80%"])].copy()
    negative_profit_mawb = mawb_all[mawb_all["Profit"] < 0].copy()
    both_zero = mawb_all[(mawb_all["Total_AR"] == 0) & (mawb_all["Total_AP"] == 0)].copy()
    sell_zero_only = mawb_all[(mawb_all["Total_AR"] == 0) & (mawb_all["Total_AP"] > 0)].copy()
    cost_zero_only = mawb_all[(mawb_all["Total_AP"] == 0) & (mawb_all["Total_AR"] > 0)].copy()

    # =========================
    # TLMF integrated view with AR allocation
    # =========================
    df2 = df.copy()
    df2["Is_TLMF"] = df2["Charge Code"].astype(str).str.upper().eq(TLMF_CODE)

    tlmf_sell_by_mawb = df2[df2["Is_TLMF"]].groupby("MAWB")["Sell Amount"].sum().rename("TLMF_AR_Total")
    tlmf_cost_by_mawb_vendor = (
        df2[df2["Is_TLMF"]]
        .groupby(["MAWB", "Vendor"])["Cost Amount"].sum()
        .rename("TLMF_AP_Vendor")
        .reset_index()
    )
    tlmf_cost_total_by_mawb = tlmf_cost_by_mawb_vendor.groupby("MAWB")["TLMF_AP_Vendor"].sum().rename("TLMF_AP_Total")

    tlmf_vendor_cnt = tlmf_cost_by_mawb_vendor.groupby("MAWB")["Vendor"].nunique().rename("TLMF_Vendor_Cnt")
    tlmf_cost_by_mawb_vendor = tlmf_cost_by_mawb_vendor.merge(
        tlmf_cost_total_by_mawb.reset_index(), on="MAWB", how="left"
    ).merge(
        tlmf_sell_by_mawb.reset_index(), on="MAWB", how="left"
    ).merge(
        tlmf_vendor_cnt.reset_index(), on="MAWB", how="left"
    )

    tlmf_cost_by_mawb_vendor["TLMF_AR_Total"] = tlmf_cost_by_mawb_vendor["TLMF_AR_Total"].fillna(0.0)
    tlmf_cost_by_mawb_vendor["TLMF_AP_Total"] = tlmf_cost_by_mawb_vendor["TLMF_AP_Total"].fillna(0.0)
    tlmf_cost_by_mawb_vendor["TLMF_Vendor_Cnt"] = tlmf_cost_by_mawb_vendor["TLMF_Vendor_Cnt"].fillna(1)

    def alloc_ar(row):
        ar_total = float(row["TLMF_AR_Total"])
        ap_total = float(row["TLMF_AP_Total"])
        ap_vendor = float(row["TLMF_AP_Vendor"])
        if ar_total == 0:
            return 0.0
        if ap_total > 0:
            return ar_total * (ap_vendor / ap_total)
        # fallback equal split
        n = max(int(row["TLMF_Vendor_Cnt"]), 1)
        return ar_total / n

    tlmf_cost_by_mawb_vendor["TLMF_AR_Allocated"] = tlmf_cost_by_mawb_vendor.apply(alloc_ar, axis=1)
    alloc_map = tlmf_cost_by_mawb_vendor[["MAWB", "Vendor", "TLMF_AR_Allocated"]].copy()

    base = (
        df2.groupby(["MAWB", "Client", "Vendor", "Charge Code"], as_index=False)
           .agg(AP=("Cost Amount", "sum"), AR=("Sell Amount", "sum"))
    )
    base["Is_TLMF"] = base["Charge Code"].astype(str).str.upper().eq(TLMF_CODE)
    base["AR_Adj"] = np.where(base["Is_TLMF"], 0.0, base["AR"].astype(float))
    base = base.merge(alloc_map, on=["MAWB", "Vendor"], how="left")
    base["TLMF_AR_Allocated"] = base["TLMF_AR_Allocated"].fillna(0.0)
    base.loc[base["Is_TLMF"], "AR_Adj"] = base.loc[base["Is_TLMF"], "TLMF_AR_Allocated"]

    integrated = base.rename(columns={"AP": "AP", "AR_Adj": "AR"}).copy()
    integrated = add_profit_cols(integrated, "AR", "AP")

    vendor_by_code = (
        integrated.groupby(["Vendor", "Charge Code"], as_index=False)
        .agg(AP=("AP", "sum"), AR=("AR", "sum"))
    )
    vendor_by_code = add_profit_cols(vendor_by_code, "AR", "AP")

    client_by_code = (
        integrated.groupby(["Client", "Charge Code"], as_index=False)
        .agg(AP=("AP", "sum"), AR=("AR", "sum"))
    )
    client_by_code = add_profit_cols(client_by_code, "AR", "AP")

    # Profit<0 list to homepage
    vendor_neg = vendor_by_code[vendor_by_code["Profit"] < 0].copy()
    vendor_neg.insert(0, "Type", "Vendor")
    vendor_neg.rename(columns={"Vendor": "Name"}, inplace=True)
    vendor_neg = vendor_neg[["Type", "Name", "Charge Code", "AP", "AR", "Profit", "Profit Margin %"]].sort_values("Profit")

    client_neg = client_by_code[client_by_code["Profit"] < 0].copy()
    client_neg.insert(0, "Type", "Client")
    client_neg.rename(columns={"Client": "Name"}, inplace=True)
    client_neg = client_neg[["Type", "Name", "Charge Code", "AP", "AR", "Profit", "Profit Margin %"]].sort_values("Profit")

    integrated_negative_on_home = pd.concat([vendor_neg, client_neg], ignore_index=True)
    integrated_negative_on_home = integrated_negative_on_home.sort_values(["Type", "Profit"])

    # =========================
    # Risk model (baseline = current file normal range 30%-80% by client)
    # =========================
    flags = mawb_all[["MAWB", "Client", "Total_AP", "Total_AR", "Profit", "Profit Margin %"]].copy()

    normal = flags[
        (~flags["Profit Margin %"].isna()) &
        (flags["Profit Margin %"] >= 0.30) & (flags["Profit Margin %"] <= 0.80) &
        (flags["Total_AR"] > 0)
    ].copy()

    overall_baseline = (normal["Profit"].sum() / normal["Total_AR"].sum()) if normal["Total_AR"].sum() else np.nan
    client_baseline = (
        normal.groupby("Client", as_index=False)
              .agg(N_Profit=("Profit", "sum"), N_AR=("Total_AR", "sum"))
    )
    client_baseline["Baseline_Margin"] = np.where(client_baseline["N_AR"] > 0,
                                                  client_baseline["N_Profit"] / client_baseline["N_AR"],
                                                  np.nan)
    flags = flags.merge(client_baseline[["Client", "Baseline_Margin"]], on="Client", how="left")
    flags["Baseline_Margin"] = flags["Baseline_Margin"].fillna(overall_baseline)

    flags["Expected_Profit"] = flags["Total_AR"] * flags["Baseline_Margin"]
    flags["Excess_Profit"] = np.maximum(0.0, flags["Profit"] - flags["Expected_Profit"])
    flags["Shortfall_Profit"] = np.maximum(0.0, flags["Expected_Profit"] - flags["Profit"])

    flags["LowMarginFlag"] = (flags["Profit Margin %"] < 0.30) & (~flags["Profit Margin %"].isna())
    flags["HighMarginFlag"] = (flags["Profit Margin %"] > 0.80) & (~flags["Profit Margin %"].isna())
    flags["NegativeProfitFlag"] = flags["Profit"] < 0

    # TLMF structural flag (simple MAWB-level): Has TLMF and (TLMF multi-vendor OR any vendor-by-code TLMF profit<0)
    has_tlmf = df2.groupby("MAWB")["Is_TLMF"].any().rename("Has_TLMF")
    tlmf_vendor_cnt2 = df2[df2["Is_TLMF"]].groupby("MAWB")["Vendor"].nunique().rename("TLMF_Vendor_Cnt2")
    tlmf_mawb_vendor = integrated[integrated["Charge Code"].astype(str).str.upper().eq(TLMF_CODE)].copy()
    tlmf_min_vendor_profit = tlmf_mawb_vendor.groupby("MAWB")["Profit"].min().rename("Min_TLMF_Vendor_Profit")

    flags = flags.merge(has_tlmf.reset_index(), on="MAWB", how="left")
    flags = flags.merge(tlmf_vendor_cnt2.reset_index(), on="MAWB", how="left")
    flags = flags.merge(tlmf_min_vendor_profit.reset_index(), on="MAWB", how="left")
    flags["Has_TLMF"] = flags["Has_TLMF"].fillna(False)
    flags["TLMF_Vendor_Cnt2"] = flags["TLMF_Vendor_Cnt2"].fillna(0)
    # Min_TLMF_Vendor_Profit can be NaN
    flags["TLMF_Structural_Flag"] = (
        flags["Has_TLMF"] &
        ((flags["TLMF_Vendor_Cnt2"] > 1) | (flags["Min_TLMF_Vendor_Profit"] < 0))
    )

    # Overlap (disclosure only)
    overlap = flags[flags["HighMarginFlag"] & flags["TLMF_Structural_Flag"]].copy()

    # Mutual overstatement buckets
    tlmf_bucket = flags[flags["TLMF_Structural_Flag"]].copy()
    high_only_bucket = flags[flags["HighMarginFlag"] & (~flags["TLMF_Structural_Flag"])].copy()

    low_bucket = flags[flags["LowMarginFlag"]].copy()
    neg_bucket = flags[flags["NegativeProfitFlag"]].copy()

    estimated_overstatement = float(tlmf_bucket["Excess_Profit"].sum() + high_only_bucket["Excess_Profit"].sum())
    estimated_underpricing = float(low_bucket["Shortfall_Profit"].sum() + np.maximum(0.0, -neg_bucket["Profit"]).sum())

    def risk_row(name, dfm):
        ap = float(dfm["Total_AP"].sum()) if len(dfm) else 0.0
        ar = float(dfm["Total_AR"].sum()) if len(dfm) else 0.0
        prof = float(dfm["Profit"].sum()) if len(dfm) else 0.0
        pm = (prof / ar) if ar else np.nan
        return {
            "Risk Type": name,
            "MAWB Count": int(dfm["MAWB"].nunique()) if len(dfm) else 0,
            "Total AP": ap,
            "Total AR": ar,
            "Total Profit": prof,
            "Profit Margin %": pm
        }

    risk_summary = pd.DataFrame([
        risk_row("Low Margin Risk (PM<30%)", flags[flags["LowMarginFlag"]]),
        risk_row("High Margin Risk (PM>80%, non-TLMF)", flags[flags["HighMarginFlag"] & (~flags["TLMF_Structural_Flag"])]),
        risk_row("TLMF Structural Risk", flags[flags["TLMF_Structural_Flag"]]),
        risk_row("Negative Profit", flags[flags["NegativeProfitFlag"]]),
    ])

    # =========================
    # ✅ UI homepage
    # =========================
    if eta_parse_note:
        st.info(eta_parse_note)

    st.subheader("Analysis Summary (Audit Control Dashboard)")

    if structural_totals is not None:
        tmp = structural_totals.copy()
        tmp["Profit Margin %"] = tmp["Profit Margin %"].apply(format_pct)
        st.markdown("**PROCARESX (Structural; excluded from detail)**")
        st.dataframe(tmp, use_container_width=True)

    tmp_rs = risk_summary.copy()
    tmp_rs["Profit Margin %"] = tmp_rs["Profit Margin %"].apply(format_pct)
    st.markdown("**Risk Summary (Auditable)**")
    st.dataframe(tmp_rs, use_container_width=True)

    st.markdown(f"**Estimated Overstatement:** {estimated_overstatement:,.2f}")
    st.markdown(f"**Estimated Underpricing:** {estimated_underpricing:,.2f}")

    st.markdown("**Overlap (High Margin ∩ TLMF Structural) — disclosure only**")
    st.write(f"Overlap MAWB Count: {int(overlap['MAWB'].nunique())} | Overlap Total Profit: {float(overlap['Profit'].sum()):,.2f}")

    st.markdown("**Integrated View Negative Profit (Vendor/Client, Profit<0)**")
    tmp_in = integrated_negative_on_home.copy()
    tmp_in["Profit Margin %"] = tmp_in["Profit Margin %"].apply(format_pct)
    st.dataframe(tmp_in, use_container_width=True, height=360)

    # =========================
    # Export (Excel)
    # =========================
    output = io.BytesIO()

    # Raw_Billing_Enriched (auditable only) — ✅ includes Destination already
    raw_billing_enriched = df.copy()

    # Also create MAWB->Destination mapping and attach to every MAWB tab
    mawb_all_out = attach_destination_if_mawb(mawb_all.rename(columns={"Total_AP": "Total_Cost", "Total_AR": "Total_Sell"}), mawb_to_dest)
    exceptions_out = attach_destination_if_mawb(exceptions.rename(columns={"Total_AP": "Total_Cost", "Total_AR": "Total_Sell"}), mawb_to_dest)
    margin_anomalies_out = attach_destination_if_mawb(margin_anomalies.rename(columns={"Total_AP": "Total_Cost", "Total_AR": "Total_Sell"}), mawb_to_dest)
    negative_profit_out = attach_destination_if_mawb(negative_profit_mawb.rename(columns={"Total_AP": "Total_Cost", "Total_AR": "Total_Sell"}), mawb_to_dest)
    both_zero_out = attach_destination_if_mawb(both_zero.rename(columns={"Total_AP": "Total_Cost", "Total_AR": "Total_Sell"}), mawb_to_dest)
    sell_zero_only_out = attach_destination_if_mawb(sell_zero_only.rename(columns={"Total_AP": "Total_Cost", "Total_AR": "Total_Sell"}), mawb_to_dest)
    cost_zero_only_out = attach_destination_if_mawb(cost_zero_only.rename(columns={"Total_AP": "Total_Cost", "Total_AR": "Total_Sell"}), mawb_to_dest)

    # ✅ Raw_Billing_Enriched 也强制补 Destination（防止原始文件 Destination 列为空或被识别不到）
    raw_billing_enriched_out = attach_destination_if_mawb(raw_billing_enriched, mawb_to_dest)

    # Flags table (optional) includes MAWB -> add destination too
    flags_out = attach_destination_if_mawb(flags, mawb_to_dest)

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        workbook = writer.book
        header_fmt = workbook.add_format({"bold": True, "font_size": 14})
        subheader_fmt = workbook.add_format({"bold": True, "font_size": 12})
        bold_fmt = workbook.add_format({"bold": True})
        percent_fmt = workbook.add_format({"num_format": "0.00%"})
        number_fmt = workbook.add_format({"num_format": "#,##0.00"})

        # =========================
        # ✅ Analysis Summary (第一页：真正嵌入内容，而不是只有超链接)
        # =========================
        ws = workbook.add_worksheet("Analysis Summary")
        writer.sheets["Analysis Summary"] = ws
        ws.write(0, 0, "Analysis Summary (Audit Control Dashboard)", header_fmt)

        row = 2

        # PROCARESX totals
        if structural_totals is not None:
            ws.write(row, 0, "PROCARESX (Structural; excluded from ALL detail)", subheader_fmt)
            structural_totals.to_excel(writer, index=False, sheet_name="Analysis Summary", startrow=row + 1, startcol=0)
            # format percent col if exists
            if "Profit Margin %" in structural_totals.columns:
                pmc = list(structural_totals.columns).index("Profit Margin %")
                excel_set_percent_col(ws, pmc, workbook)
            row = row + 3 + len(structural_totals)

        # Risk Summary table
        ws.write(row, 0, "Risk Summary (Auditable)", subheader_fmt)
        risk_summary.to_excel(writer, index=False, sheet_name="Analysis Summary", startrow=row + 1, startcol=0)

        # format percent in risk summary
        if "Profit Margin %" in risk_summary.columns:
            pmc = list(risk_summary.columns).index("Profit Margin %")
            excel_set_percent_col(ws, pmc, workbook)
        row = row + 3 + len(risk_summary)

        # Estimates + overlap
        ws.write(row, 0, "Estimated Overstatement", bold_fmt)
        ws.write_number(row, 1, float(estimated_overstatement), number_fmt)
        ws.write(row + 1, 0, "Estimated Underpricing", bold_fmt)
        ws.write_number(row + 1, 1, float(estimated_underpricing), number_fmt)

        ws.write(row + 3, 0, "Overlap (High Margin ∩ TLMF Structural) — disclosure only", bold_fmt)
        ws.write(row + 4, 0, "Overlap MAWB Count")
        ws.write_number(row + 4, 1, int(overlap["MAWB"].nunique()), number_fmt)
        ws.write(row + 5, 0, "Overlap Total Profit")
        ws.write_number(row + 5, 1, float(overlap["Profit"].sum()), number_fmt)

        row = row + 7

        # Integrated negative profit list
        ws.write(row, 0, "Integrated View – Negative Profit Exposure (Profit<0)", subheader_fmt)
        integrated_negative_on_home.to_excel(writer, index=False, sheet_name="Analysis Summary", startrow=row + 1, startcol=0)
        # format percent col if exists
        if "Profit Margin %" in integrated_negative_on_home.columns:
            pmc = list(integrated_negative_on_home.columns).index("Profit Margin %")
            excel_set_percent_col(ws, pmc, workbook)

        row = row + 3 + len(integrated_negative_on_home)

        # Links section at bottom
        ws.write(row, 0, "Links to detail tabs:", bold_fmt)
        row += 1
        tab_links = [
            ("MAWB Summary", "MAWB_Summary"),
            ("Exceptions", "Exceptions"),
            ("Margin Anomalies", "Margin_Anomalies"),
            ("Negative Profit (MAWB)", "Negative_Profit"),
            ("Both Zero", "Both_Zero"),
            ("Sell Zero Only", "Sell_Zero_Only"),
            ("Cost Zero Only", "Cost_Zero_Only"),
            ("Raw Billing Enriched", "Raw_Billing_Enriched"),
            ("Risk Flags (MAWB)", "Risk_Flags_MAWB"),
        ]
        for text, sheet in tab_links:
            ws.write_url(row, 0, f"internal:'{sheet}'!A1", string=text)
            row += 1

        # =========================
        # Detail tabs (Destination already attached for MAWB tabs + raw)
        # =========================
        mawb_all_out = to_date_only(mawb_all_out, ["ETA"])
        exceptions_out = to_date_only(exceptions_out, ["ETA"])
        margin_anomalies_out = to_date_only(margin_anomalies_out, ["ETA"])
        negative_profit_out = to_date_only(negative_profit_out, ["ETA"])
        both_zero_out = to_date_only(both_zero_out, ["ETA"])
        sell_zero_only_out = to_date_only(sell_zero_only_out, ["ETA"])
        cost_zero_only_out = to_date_only(cost_zero_only_out, ["ETA"])
        raw_billing_enriched_out = to_date_only(raw_billing_enriched_out, ["ETA"])
        flags_out = to_date_only(flags_out, [])

        mawb_all_out.to_excel(writer, index=False, sheet_name="MAWB_Summary")
        exceptions_out.to_excel(writer, index=False, sheet_name="Exceptions")
        margin_anomalies_out.to_excel(writer, index=False, sheet_name="Margin_Anomalies")
        negative_profit_out.to_excel(writer, index=False, sheet_name="Negative_Profit")
        both_zero_out.to_excel(writer, index=False, sheet_name="Both_Zero")
        sell_zero_only_out.to_excel(writer, index=False, sheet_name="Sell_Zero_Only")
        cost_zero_only_out.to_excel(writer, index=False, sheet_name="Cost_Zero_Only")
        raw_billing_enriched_out.to_excel(writer, index=False, sheet_name="Raw_Billing_Enriched")
        flags_out.to_excel(writer, index=False, sheet_name="Risk_Flags_MAWB")

        # percent formatting on all sheets that have Profit Margin %
        for sh, dfx in [
            ("MAWB_Summary", mawb_all_out),
            ("Exceptions", exceptions_out),
            ("Margin_Anomalies", margin_anomalies_out),
            ("Negative_Profit", negative_profit_out),
            ("Both_Zero", both_zero_out),
            ("Sell_Zero_Only", sell_zero_only_out),
            ("Cost_Zero_Only", cost_zero_only_out),
            ("Raw_Billing_Enriched", raw_billing_enriched_out),
            ("Risk_Flags_MAWB", flags_out),
        ]:
            ws2 = writer.sheets[sh]
            if "Profit Margin %" in dfx.columns:
                pmc = list(dfx.columns).index("Profit Margin %")
                excel_set_percent_col(ws2, pmc, workbook)
            # 数字列宽友好
            for coln in ["Sell Amount", "Cost Amount", "Profit", "Total_Sell", "Total_Cost", "Total_AR", "Total_AP", "Total Profit"]:
                if coln in dfx.columns:
                    ci = list(dfx.columns).index(coln)
                    excel_set_number_col(ws2, ci, workbook)

    st.download_button(
        "Download Report Excel (Analysis Summary embedded + Destination on all MAWB tabs)",
        data=output.getvalue(),
        file_name="MAWB_Audit_Report_FINAL.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

except Exception as e:
    st.exception(e)
