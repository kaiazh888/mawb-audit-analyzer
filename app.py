import io
import re
import numpy as np
import pandas as pd
import streamlit as st

st.set_page_config(page_title="MAWB Audit Analyzer", layout="wide")
st.title("MAWB Audit Analyzer (Billing-only) — Enhanced")
st.caption(
    "Upload Billing charges export + optional MAWB→ETA mapping file. "
    "Supports MAWB filter box, profit margin analysis, zero buckets, outliers, negative profit, "
    "and Charge Code / Vendor summaries. Enhanced with: "
    "Vendor→Primary Charge Code (audit by Cost) + TLMF AR allocation."
)

# ---------------- Helpers ----------------
def safe_numeric(s: pd.Series) -> pd.Series:
    return pd.to_numeric(s, errors="coerce").fillna(0.0)

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

    # YYYYMMDD
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

def pct(numer: pd.Series, denom: pd.Series) -> pd.Series:
    return (numer / denom).where(denom != 0, 0)

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

def format_pct_str(x):
    try:
        return f"{float(x) * 100:.2f}%"
    except Exception:
        return ""

def make_kpi_vertical(kpi_dict: dict, pct_keys: set[str]) -> pd.DataFrame:
    rows = []
    for k, v in kpi_dict.items():
        rows.append({"Metric": k, "Value": format_pct_str(v) if k in pct_keys else v})
    return pd.DataFrame(rows)

# Excel formatting helpers
def excel_set_percent_col(ws, col_idx: int, workbook, width: int = 16):
    fmt = workbook.add_format({"num_format": "0.00%"})
    ws.set_column(col_idx, col_idx, width, fmt)

def excel_set_currency_col(ws, col_idx: int, workbook, width: int = 16):
    fmt = workbook.add_format({"num_format": "#,##0.00"})
    ws.set_column(col_idx, col_idx, width, fmt)

# ---- Charge Code → Category mapping (for vendor mapping view; extend if needed) ----
bucket_defs = {
    "DTRF": "Trucking / Transfer",
    "TLMF": "Delivery Cartage / Last-mile",
    "TISC": "ISC / Customs",
    "THAWB": "Brokerage / HAWB Clearance",
    "TABD": "Airline Breakdown / Handling",
    "DATXF": "Airline Transfer AP",
    "DDOC": "Documentation",
    "DSTOR": "Storage / Warehouse",
    "LABEL": "Label / Packaging",
}
def map_bucket(code: str) -> str:
    c = (code or "").strip().upper()
    if c in bucket_defs:
        return bucket_defs[c]
    if c.startswith("D"):
        return "Operational (D*) - Other"
    if c.startswith("T"):
        return "Operational (T*) - Other"
    return "Other / Unclassified"

def confidence_from_share(share: float) -> str:
    if share >= 0.80:
        return "High"
    if share >= 0.60:
        return "Medium"
    return "Low"

# ---------------- Uploaders ----------------
billing_file = st.file_uploader("Upload Billing Charges Excel (.xlsx)", type=["xlsx"], key="billing")
eta_file = st.file_uploader("Optional: Upload MAWB→ETA mapping Excel (.xlsx)", type=["xlsx"], key="eta_mapping")

st.divider()
st.subheader("Optional Filter: Keep only specified MAWBs")
mawb_text = st.text_area(
    "Paste MAWBs here (comma / space / newline separated). Supports 99934022122 → 999-34022122. Leave blank to keep all.",
    height=140,
    placeholder="Example:\n999-34022122\n99934022133\n999 34022144"
)

st.divider()
st.subheader("Enhanced: TLMF AR Allocation (for AR集中 / AP分散)")
enable_tlmf_alloc = st.checkbox("Enable TLMF vendor-level AR allocation (recommended)", value=True)
alloc_method = st.radio(
    "Allocation method",
    ["By Cost Share (recommended)", "By Line Count (fallback)"],
    horizontal=True,
    disabled=not enable_tlmf_alloc
)
treat_blank_vendor_as_ar_line = st.checkbox("TLMF: treat blank Vendor rows as AR consolidated line", value=True)
low_cost_threshold = st.number_input("TLMF: low-cost threshold (<=) treated as AR line", min_value=0.0, value=1.0, step=1.0)
sell_positive_threshold = st.number_input("TLMF: sell positive threshold (>) for AR line", min_value=0.0, value=0.0, step=10.0)
vendor_margin_low = st.number_input("TLMF vendor anomaly: low margin (<)", min_value=-10.0, max_value=1.0, value=0.30, step=0.05, format="%.2f")
vendor_margin_high = st.number_input("TLMF vendor anomaly: high margin (>)", min_value=0.0, max_value=10.0, value=0.80, step=0.05, format="%.2f")

# ---------------- Config: column candidates ----------------
BILLING_REQUIRED = {
    "MAWB": ["MAWB", "Mawb", "Master AWB", "MasterAWB"],
    "Cost Amount": ["Cost Amount", "Cost", "AP Amount", "Total Cost", "CostAmount"],
    "Sell Amount": ["Sell Amount", "Sell", "AR Amount", "Total Sell", "SellAmount"],
}
BILLING_OPTIONAL = {
    "Client": ["Client", "Customer", "Account", "Shipper", "Bill To", "Billed To"],
    "Charge Code": ["Charge Code", "ChargeCode", "Charge", "Code"],
    "Vendor": ["Vendor", "Carrier", "Supplier"],
}
ETA_REQUIRED = {
    "MAWB": ["MAWB", "Mawb", "Master AWB", "MasterAWB"],
    "ETA": ["ETA", "Eta", "Estimated Time of Arrival", "Arrival", "Arrival Date", "ETA Date"],
}

# ---------------- Main ----------------
if not billing_file:
    st.info("Please upload a Billing Charges Excel file to start.")
    st.stop()

try:
    # ✅ split margin labels
    MARGIN_LOW_LABEL = "Margin<30%"
    MARGIN_HIGH_LABEL = "Margin>80%"

    # ---- Read billing charges ----
    xls = pd.ExcelFile(billing_file)
    billing_sheet = find_sheet_with_required_cols(xls, BILLING_REQUIRED)
    if not billing_sheet:
        st.error(
            "Could not find a sheet in the Billing file containing required fields:\n"
            "- MAWB\n- Cost Amount\n- Sell Amount\n\n"
            "Tip: check your headers in the export."
        )
        st.stop()

    raw_df = pd.read_excel(xls, sheet_name=billing_sheet)

    mawb_col = find_first_col(raw_df, BILLING_REQUIRED["MAWB"])
    cost_col = find_first_col(raw_df, BILLING_REQUIRED["Cost Amount"])
    sell_col = find_first_col(raw_df, BILLING_REQUIRED["Sell Amount"])
    client_col = find_first_col(raw_df, BILLING_OPTIONAL["Client"])
    charge_code_col = find_first_col(raw_df, BILLING_OPTIONAL["Charge Code"])
    vendor_col = find_first_col(raw_df, BILLING_OPTIONAL["Vendor"])

    if not (mawb_col and cost_col and sell_col):
        st.error("Billing sheet found but required columns could not be detected after scanning.")
        st.stop()

    # Normalize billing
    df = raw_df.copy()
    df["MAWB"] = df[mawb_col].apply(normalize_mawb)
    df["Cost Amount"] = safe_numeric(df[cost_col])
    df["Sell Amount"] = safe_numeric(df[sell_col])

    df["Client"] = df[client_col].astype(str).str.strip() if client_col else "UNKNOWN"
    df.loc[df["Client"].isin(["", "nan", "None"]), "Client"] = "UNKNOWN"

    df["Charge Code"] = df[charge_code_col].astype(str).str.strip().str.upper() if charge_code_col else "UNKNOWN"
    df.loc[df["Charge Code"].isin(["", "nan", "None"]), "Charge Code"] = "UNKNOWN"

    df["Vendor"] = df[vendor_col].astype(str).str.strip().str.upper() if vendor_col else "UNKNOWN"
    df.loc[df["Vendor"].isin(["", "nan", "None"]), "Vendor"] = "UNKNOWN"

    df = df[df["MAWB"].ne("")].copy()

    # ---- Optional MAWB filter ----
    mawb_keep = parse_mawb_list(mawb_text)
    if mawb_keep:
        before_rows = len(df)
        before_mawb = df["MAWB"].nunique()

        df = df[df["MAWB"].isin(mawb_keep)].copy()

        after_rows = len(df)
        after_mawb = df["MAWB"].nunique()

        found_set = set(df["MAWB"].unique())
        mawb_not_found = sorted(set(mawb_keep) - found_set)

        st.info(f"MAWB filter applied: rows {before_rows} → {after_rows}, unique MAWB {before_mawb} → {after_mawb}.")
    else:
        mawb_not_found = []

    # ---- Read ETA mapping (optional) ----
    eta_map = None
    eta_parse_note = None

    if eta_file:
        xls2 = pd.ExcelFile(eta_file)
        map_sheet = find_sheet_with_required_cols(xls2, ETA_REQUIRED)

        if not map_sheet:
            st.warning("ETA mapping file uploaded, but could not find MAWB + ETA columns in any sheet.")
        else:
            mdf0 = pd.read_excel(xls2, sheet_name=map_sheet)
            m_mawb = find_first_col(mdf0, ETA_REQUIRED["MAWB"])
            m_eta = find_first_col(mdf0, ETA_REQUIRED["ETA"])

            if not (m_mawb and m_eta):
                st.warning("ETA mapping sheet found, but MAWB/ETA columns could not be detected.")
            else:
                mdf = mdf0[[m_mawb, m_eta]].copy()
                mdf.columns = ["MAWB", "ETA"]
                mdf["MAWB"] = mdf["MAWB"].apply(normalize_mawb)
                mdf["ETA"] = clean_eta_series(mdf["ETA"])

                bad_eta_rows = int(mdf["ETA"].isna().sum())
                total_rows = int(len(mdf))
                if total_rows > 0 and bad_eta_rows > 0:
                    eta_parse_note = f"ETA parsing note: {bad_eta_rows} / {total_rows} ETA values could not be parsed and were left blank."

                eta_map = (
                    mdf.dropna(subset=["MAWB"])
                       .groupby("MAWB", as_index=False)["ETA"]
                       .max()
                )

    # ---- Merge ETA into billing ----
    if eta_map is not None and not eta_map.empty:
        df = df.merge(eta_map, on="MAWB", how="left")
    else:
        df["ETA"] = pd.NaT

    df["ETA"] = pd.to_datetime(df["ETA"], errors="coerce").dt.normalize()

    # ---- MAWB summary ----
    summary = (
        df.groupby("MAWB", as_index=False)
          .agg(
              Client=("Client", "first"),
              Total_Cost=("Cost Amount", "sum"),
              Total_Sell=("Sell Amount", "sum"),
              Line_Count=("MAWB", "size"),
              ETA=("ETA", "max")
          )
    )
    summary["ETA Month"] = summary["ETA"].dt.to_period("M").astype(str).replace("NaT", "")

    summary["Profit"] = summary["Total_Sell"] - summary["Total_Cost"]
    summary["Profit Margin %"] = pct(summary["Profit"], summary["Total_Sell"])

    # ✅ Classification rule stays MAWB-level (keep your original logic), but margin split for reporting
    def is_closed(r):
        if not (r["Total_Cost"] > 0 and r["Total_Sell"] > 0):
            return "Open"
        pm = r["Profit Margin %"]
        if (pm < 0.30) or (pm > 0.80):
            return "Open"
        return "Closed"

    summary["Classification"] = summary.apply(is_closed, axis=1)

    def exception_type(r):
        if r["Total_Cost"] == 0 and r["Total_Sell"] == 0:
            return "Cost=Sell=0"
        if r["Total_Sell"] == 0:
            return "Revenue=0"
        if r["Total_Cost"] == 0:
            return "Cost=0"
        pm = r["Profit Margin %"]
        if pm != 0 and pm < 0.30:
            return MARGIN_LOW_LABEL
        if pm != 0 and pm > 0.80:
            return MARGIN_HIGH_LABEL
        return ""

    summary["Exception_Type"] = summary.apply(exception_type, axis=1)
    exceptions = summary[summary["Classification"].eq("Open")].copy()

    # ---- Client Summary ----
    client_summary = (
        df.groupby("Client", as_index=False)
          .agg(
              Total_Cost=("Cost Amount", "sum"),
              Total_Sell=("Sell Amount", "sum"),
              Line_Count=("Client", "size"),
              MAWB_Count=("MAWB", pd.Series.nunique),
              Latest_ETA=("ETA", "max"),
          )
    )
    client_summary["Profit"] = client_summary["Total_Sell"] - client_summary["Total_Cost"]
    client_summary["Profit Margin %"] = pct(client_summary["Profit"], client_summary["Total_Sell"])
    client_summary = client_summary.sort_values("Profit", ascending=False)

    # ---- Margin Outliers / Negative Profit ----
    margin_outliers = summary[
        ((summary["Profit Margin %"] < 0.30) | (summary["Profit Margin %"] > 0.80)) &
        (summary["Profit Margin %"] != 0)
    ].copy().sort_values("Profit Margin %")

    negative_profit = summary[summary["Profit"] < 0].copy().sort_values("Profit")

    # ---- Zero buckets (Profit/Margin) ----
    zero_margin = summary[summary["Profit Margin %"] == 0].copy().sort_values(["Total_Sell", "Total_Cost"], ascending=False)
    zero_profit = summary[summary["Profit"] == 0].copy().sort_values(["Total_Sell", "Total_Cost"], ascending=False)

    both_zero = summary[(summary["Total_Sell"] == 0) & (summary["Total_Cost"] == 0)].copy().sort_values("MAWB")
    sell_zero_only = summary[(summary["Total_Sell"] == 0) & (summary["Total_Cost"] > 0)].copy().sort_values("Total_Cost", ascending=False)
    cost_zero_only = summary[(summary["Total_Cost"] == 0) & (summary["Total_Sell"] > 0)].copy().sort_values("Total_Sell", ascending=False)

    # ---- Charge Code Summary ----
    chargecode_summary = (
        df.groupby("Charge Code", as_index=False)
          .agg(
              Total_Cost=("Cost Amount", "sum"),
              Total_Sell=("Sell Amount", "sum"),
              Line_Count=("Charge Code", "size"),
              MAWB_Count=("MAWB", pd.Series.nunique),
          )
    )
    chargecode_summary["Profit"] = chargecode_summary["Total_Sell"] - chargecode_summary["Total_Cost"]
    chargecode_summary["Profit Margin %"] = pct(chargecode_summary["Profit"], chargecode_summary["Total_Sell"])
    chargecode_summary = chargecode_summary.sort_values("Profit", ascending=False)

    # Charge code exception counts (MAWB-level flags)
    mawb_flags = summary[["MAWB", "Exception_Type"]].copy()
    mawb_charge = df[["MAWB", "Charge Code"]].drop_duplicates()
    cc_exc = mawb_charge.merge(mawb_flags, on="MAWB", how="left")
    chargecode_exceptions = (
        cc_exc.pivot_table(
            index="Charge Code",
            columns="Exception_Type",
            values="MAWB",
            aggfunc=pd.Series.nunique,
            fill_value=0
        ).reset_index()
    )
    chargecode_summary = chargecode_summary.merge(chargecode_exceptions, on="Charge Code", how="left").fillna(0)

    # ---- Vendor Summary ----
    vendor_summary = (
        df.groupby("Vendor", as_index=False)
          .agg(
              Total_Cost=("Cost Amount", "sum"),
              Total_Sell=("Sell Amount", "sum"),
              Line_Count=("Vendor", "size"),
              MAWB_Count=("MAWB", pd.Series.nunique),
          )
    )
    vendor_summary["Profit"] = vendor_summary["Total_Sell"] - vendor_summary["Total_Cost"]
    vendor_summary["Profit Margin %"] = pct(vendor_summary["Profit"], vendor_summary["Total_Sell"])
    vendor_summary = vendor_summary.sort_values("Profit", ascending=False)

    mawb_vendor = df[["MAWB", "Vendor"]].drop_duplicates()
    v_exc = mawb_vendor.merge(mawb_flags, on="MAWB", how="left")
    vendor_exceptions = (
        v_exc.pivot_table(
            index="Vendor",
            columns="Exception_Type",
            values="MAWB",
            aggfunc=pd.Series.nunique,
            fill_value=0
        ).reset_index()
    )
    vendor_summary = vendor_summary.merge(vendor_exceptions, on="Vendor", how="left").fillna(0)

    # ---- Charge Code Profit <= 0 by MAWB ----
    cc_mawb = (
        df.groupby(["MAWB", "Charge Code"], as_index=False)
          .agg(
              Client=("Client", "first"),
              Vendor=("Vendor", "first"),
              Total_Cost=("Cost Amount", "sum"),
              Total_Sell=("Sell Amount", "sum"),
              ETA=("ETA", "max"),
          )
    )
    cc_mawb["Profit"] = cc_mawb["Total_Sell"] - cc_mawb["Total_Cost"]
    cc_mawb["Profit Margin %"] = pct(cc_mawb["Profit"], cc_mawb["Total_Sell"])
    cc_mawb["ETA Month"] = pd.to_datetime(cc_mawb["ETA"], errors="coerce").dt.to_period("M").astype(str).replace("NaT", "")

    chargecode_profit_le0_mawb = cc_mawb[cc_mawb["Profit"] <= 0].copy().sort_values(
        ["Profit", "Total_Sell"], ascending=[True, False]
    )

    # ---------------- Enhanced Module 1: Vendor→Primary Charge Code (Audit by Cost) ----------------
    df_v = df[(df["Vendor"].ne("UNKNOWN")) & (df["Vendor"].ne("")) & (df["Charge Code"].ne("UNKNOWN"))].copy()
    vendor_primary = pd.DataFrame()
    vendor_top5 = pd.DataFrame()
    vendor_mixed_risk = pd.DataFrame()

    if not df_v.empty:
        vc = (
            df_v.groupby(["Vendor", "Charge Code"], as_index=False)
               .agg(
                   Line_Count=("MAWB", "size"),
                   MAWB_Count=("MAWB", pd.Series.nunique),
                   Total_Cost=("Cost Amount", "sum"),
                   Total_Sell=("Sell Amount", "sum"),
               )
        )
        vt = (
            vc.groupby("Vendor", as_index=False)
              .agg(
                  Vendor_Total_Cost=("Total_Cost", "sum"),
                  Vendor_Total_Sell=("Total_Sell", "sum"),
                  Total_Lines=("Line_Count", "sum"),
                  Total_MAWBs=("MAWB_Count", "sum"),
              )
        )

        vc = vc.merge(vt[["Vendor", "Vendor_Total_Cost"]], on="Vendor", how="left")
        vc["Cost_Share"] = np.where(vc["Vendor_Total_Cost"] > 0, vc["Total_Cost"] / vc["Vendor_Total_Cost"], 0.0)
        vc["Category"] = vc["Charge Code"].apply(map_bucket)
        vc = vc.sort_values(["Vendor", "Total_Cost"], ascending=[True, False])
        vc["Rank"] = vc.groupby("Vendor")["Total_Cost"].rank(method="first", ascending=False)

        top1 = vc[vc["Rank"] == 1].copy()
        top2 = vc[vc["Rank"] == 2].copy()

        vendor_primary = (
            vt.merge(
                top1.rename(columns={
                    "Charge Code": "Primary_Charge_Code",
                    "Category": "Primary_Category",
                    "Total_Cost": "Primary_Cost",
                    "Total_Sell": "Primary_Sell",
                    "Cost_Share": "Primary_Cost_Share",
                })[["Vendor", "Primary_Charge_Code", "Primary_Category", "Primary_Cost", "Primary_Sell", "Primary_Cost_Share"]],
                on="Vendor",
                how="left",
            )
            .merge(
                top2.rename(columns={
                    "Charge Code": "Secondary_Charge_Code",
                    "Category": "Secondary_Category",
                    "Total_Cost": "Secondary_Cost",
                    "Cost_Share": "Secondary_Cost_Share",
                })[["Vendor", "Secondary_Charge_Code", "Secondary_Category", "Secondary_Cost", "Secondary_Cost_Share"]],
                on="Vendor",
                how="left",
            )
        )
        vendor_primary["Confidence"] = vendor_primary["Primary_Cost_Share"].apply(confidence_from_share)
        vendor_primary["Secondary_vs_Primary_Cost_Ratio"] = np.where(
            vendor_primary["Primary_Cost"].fillna(0) > 0,
            vendor_primary["Secondary_Cost"].fillna(0) / vendor_primary["Primary_Cost"].fillna(0),
            0.0,
        )
        vendor_mixed_risk = vendor_primary[
            (vendor_primary["Confidence"] == "Low") | (vendor_primary["Secondary_vs_Primary_Cost_Ratio"] >= 0.50)
        ].copy()

        vendor_top5 = vc.groupby("Vendor").head(5).copy()

    # ---------------- Enhanced Module 2: TLMF AR Allocation (Vendor-level) ----------------
    tlmf_all_rows = pd.DataFrame()
    tlmf_ar_lines = pd.DataFrame()
    tlmf_vendor_cost_lines = pd.DataFrame()
    tlmf_mawb_totals = pd.DataFrame()
    tlmf_alloc = pd.DataFrame()
    tlmf_anomalies = pd.DataFrame()

    if enable_tlmf_alloc:
        tlmf_all_rows = df[df["Charge Code"].eq("TLMF")].copy()
        if not tlmf_all_rows.empty:
            is_blank_vendor = tlmf_all_rows["Vendor"].isin(["", "UNKNOWN"])
            is_low_cost_high_sell = (tlmf_all_rows["Cost Amount"] <= float(low_cost_threshold)) & (
                tlmf_all_rows["Sell Amount"] > float(sell_positive_threshold)
            )
            ar_line_mask = is_low_cost_high_sell | (treat_blank_vendor_as_ar_line & is_blank_vendor)

            tlmf_ar_lines = tlmf_all_rows[ar_line_mask].copy()
            tlmf_vendor_cost_lines = tlmf_all_rows[~ar_line_mask].copy()

            mawb_sell = tlmf_all_rows.groupby("MAWB", as_index=False)["Sell Amount"].sum().rename(columns={"Sell Amount": "TLMF_Total_Sell"})
            mawb_cost = tlmf_all_rows.groupby("MAWB", as_index=False)["Cost Amount"].sum().rename(columns={"Cost Amount": "TLMF_Total_Cost"})
            tlmf_mawb_totals = mawb_sell.merge(mawb_cost, on="MAWB", how="left")

            vendor_cost = (
                tlmf_vendor_cost_lines[~tlmf_vendor_cost_lines["Vendor"].isin(["", "UNKNOWN"])]
                .groupby(["MAWB", "Vendor"], as_index=False)
                .agg(
                    Vendor_TLMF_Cost=("Cost Amount", "sum"),
                    Vendor_Lines=("MAWB", "size"),
                )
            )

            if not vendor_cost.empty:
                tlmf_alloc = vendor_cost.merge(tlmf_mawb_totals, on="MAWB", how="left")
                tlmf_alloc["TLMF_Total_Sell"] = tlmf_alloc["TLMF_Total_Sell"].fillna(0.0)
                tlmf_alloc["TLMF_Total_Cost"] = tlmf_alloc["TLMF_Total_Cost"].fillna(0.0)

                if alloc_method.startswith("By Cost"):
                    tlmf_alloc["Alloc_Share"] = np.where(
                        tlmf_alloc["TLMF_Total_Cost"] > 0,
                        tlmf_alloc["Vendor_TLMF_Cost"] / tlmf_alloc["TLMF_Total_Cost"],
                        0.0,
                    )
                else:
                    total_lines = tlmf_alloc.groupby("MAWB")["Vendor_Lines"].transform("sum")
                    tlmf_alloc["Alloc_Share"] = np.where(total_lines > 0, tlmf_alloc["Vendor_Lines"] / total_lines, 0.0)

                tlmf_alloc["Vendor_AR_Allocated"] = tlmf_alloc["TLMF_Total_Sell"] * tlmf_alloc["Alloc_Share"]
                tlmf_alloc["Vendor_Profit_Est"] = tlmf_alloc["Vendor_AR_Allocated"] - tlmf_alloc["Vendor_TLMF_Cost"]
                tlmf_alloc["Vendor_Margin_Est"] = np.where(
                    tlmf_alloc["Vendor_AR_Allocated"] > 0,
                    tlmf_alloc["Vendor_Profit_Est"] / tlmf_alloc["Vendor_AR_Allocated"],
                    0.0,
                )

                # anomalies
                mawb_cost_no_ar = tlmf_mawb_totals[(tlmf_mawb_totals["TLMF_Total_Cost"] > 0) & (tlmf_mawb_totals["TLMF_Total_Sell"] == 0)].copy()
                mawb_cost_no_ar["Anomaly"] = "TLMF: Cost>0 but Total Sell=0 (AR missing/posted elsewhere)"

                mawb_ar_no_cost = tlmf_mawb_totals[(tlmf_mawb_totals["TLMF_Total_Sell"] > 0) & (tlmf_mawb_totals["TLMF_Total_Cost"] == 0)].copy()
                mawb_ar_no_cost["Anomaly"] = "TLMF: Sell>0 but Total Cost=0 (AP missing/not accrued)"

                vendor_anom = tlmf_alloc.copy()
                vendor_anom["Anomaly"] = ""
                vendor_anom.loc[vendor_anom["Vendor_Profit_Est"] < 0, "Anomaly"] = "TLMF Vendor: Negative profit after allocation"
                vendor_anom.loc[vendor_anom["Vendor_Margin_Est"] < float(vendor_margin_low), "Anomaly"] = f"TLMF Vendor: Margin<{vendor_margin_low:.2f}"
                vendor_anom.loc[vendor_anom["Vendor_Margin_Est"] > float(vendor_margin_high), "Anomaly"] = f"TLMF Vendor: Margin>{vendor_margin_high:.2f}"
                vendor_anom = vendor_anom[vendor_anom["Anomaly"].ne("")].copy()

                tlmf_anomalies = pd.concat(
                    [
                        mawb_cost_no_ar[["MAWB", "TLMF_Total_Sell", "TLMF_Total_Cost", "Anomaly"]],
                        mawb_ar_no_cost[["MAWB", "TLMF_Total_Sell", "TLMF_Total_Cost", "Anomaly"]],
                        vendor_anom[["MAWB", "Vendor", "Vendor_TLMF_Cost", "Vendor_AR_Allocated", "Vendor_Profit_Est", "Vendor_Margin_Est", "Anomaly"]],
                    ],
                    ignore_index=True
                )

    # ---- KPI / Summary numbers ----
    total_mawb = len(summary)
    closed_cnt = int((summary["Classification"] == "Closed").sum())
    open_cnt = total_mawb - closed_cnt

    total_sell_sum = float(summary["Total_Sell"].sum())
    total_profit_sum = float(summary["Profit"].sum())
    overall_pm = (total_profit_sum / total_sell_sum) if total_sell_sum else 0

    neg_profit_cnt = int((summary["Profit"] < 0).sum())
    neg_profit_amt = float(summary.loc[summary["Profit"] < 0, "Profit"].sum())
    neg_profit_ratio = (neg_profit_cnt / total_mawb) if total_mawb else 0

    eta_filled_ratio = float((summary["ETA"].notna().sum() / total_mawb)) if total_mawb else 0

    kpi_dict = {
        "Total MAWB": total_mawb,
        "Closed Count": closed_cnt,
        "Closed %": (closed_cnt / total_mawb) if total_mawb else 0,
        "Open Count": open_cnt,
        "Revenue=0 Count": int((summary["Exception_Type"] == "Revenue=0").sum()),
        "Cost=0 Count": int((summary["Exception_Type"] == "Cost=0").sum()),
        "Cost=Sell=0 Count": int((summary["Exception_Type"] == "Cost=Sell=0").sum()),
        f"{MARGIN_LOW_LABEL} Count": int((summary["Exception_Type"] == MARGIN_LOW_LABEL).sum()),
        f"{MARGIN_HIGH_LABEL} Count": int((summary["Exception_Type"] == MARGIN_HIGH_LABEL).sum()),
        "Total Cost": float(summary["Total_Cost"].sum()),
        "Total Sell": total_sell_sum,
        "Total Profit": total_profit_sum,
        "Overall Profit Margin %": overall_pm,
        "ETA Filled %": eta_filled_ratio,
    }
    KPI_PCT_KEYS = {"Closed %", "Overall Profit Margin %", "ETA Filled %"}
    kpi_vertical = make_kpi_vertical(kpi_dict, KPI_PCT_KEYS)

    neg_summary = pd.DataFrame([
        {"Metric": "Profit < 0 Count", "Value": neg_profit_cnt},
        {"Metric": "Profit < 0 Total Amount", "Value": neg_profit_amt},
        {"Metric": "Profit < 0 % of MAWBs", "Value": format_pct_str(neg_profit_ratio)},
    ])

    # ---------------- UI ----------------
    if eta_parse_note:
        st.info(eta_parse_note)

    if mawb_keep:
        st.subheader("MAWB Not Found (in uploaded Billing file)")
        st.dataframe(pd.DataFrame({"MAWB": mawb_not_found}), use_container_width=True)

    st.subheader("Analysis Summary (KPI)")
    st.dataframe(kpi_vertical, use_container_width=True)

    st.subheader("Summary: Profit < 0 (Count / Amount / Ratio)")
    st.dataframe(neg_summary, use_container_width=True)

    def display_df(df_in, date_cols=None):
        out = df_in.copy()
        if date_cols:
            out = to_date_only(out, date_cols)
        for c in ["Profit Margin %", "Closed %", "ETA Filled %", "Overall Profit Margin %", "Vendor_Margin_Est", "Alloc_Share"]:
            if c in out.columns:
                out[c] = out[c].apply(format_pct_str)
        return out

    # ---- Your original tabs (kept) ----
    st.subheader("Exceptions (Open items)")
    st.dataframe(display_df(exceptions, date_cols=["ETA"]), use_container_width=True)

    st.subheader("MAWB Summary (All)")
    st.dataframe(display_df(summary, date_cols=["ETA"]), use_container_width=True)

    st.subheader("Client Profit Summary")
    st.dataframe(display_df(client_summary, date_cols=["Latest_ETA"]), use_container_width=True)

    st.subheader(f"Profit Margin Outliers ({MARGIN_LOW_LABEL} / {MARGIN_HIGH_LABEL}, PM!=0)")
    st.dataframe(display_df(margin_outliers, date_cols=["ETA"]), use_container_width=True)

    st.subheader("Negative Profit (Profit < 0)")
    st.dataframe(display_df(negative_profit, date_cols=["ETA"]), use_container_width=True)

    st.subheader("Zero Margin (Profit Margin % = 0)")
    st.dataframe(display_df(zero_margin, date_cols=["ETA"]), use_container_width=True)

    st.subheader("Zero Profit (Profit = 0)")
    st.dataframe(display_df(zero_profit, date_cols=["ETA"]), use_container_width=True)

    st.subheader("Cost=Sell=0 (Both Zero)")
    st.dataframe(display_df(both_zero, date_cols=["ETA"]), use_container_width=True)

    st.subheader("Sell=0 ONLY (Total_Sell=0 and Total_Cost>0)")
    st.dataframe(display_df(sell_zero_only, date_cols=["ETA"]), use_container_width=True)

    st.subheader("Cost=0 ONLY (Total_Cost=0 and Total_Sell>0)")
    st.dataframe(display_df(cost_zero_only, date_cols=["ETA"]), use_container_width=True)

    st.subheader("Charge Code Summary")
    st.dataframe(display_df(chargecode_summary), use_container_width=True)

    st.subheader("Vendor Summary")
    st.dataframe(display_df(vendor_summary), use_container_width=True)

    st.subheader("Charge Code Profit <= 0 (by MAWB)")
    st.dataframe(display_df(chargecode_profit_le0_mawb, date_cols=["ETA"]), use_container_width=True)

    # ---- New enhanced tabs ----
    st.divider()
    st.subheader("Enhanced: Vendor → Primary Charge Code (Audit by Cost)")
    if vendor_primary.empty:
        st.info("No vendor+chargecode rows available to build vendor→primary mapping.")
    else:
        vp_show = vendor_primary.copy()
        vp_show["Primary_Cost_Share"] = vp_show["Primary_Cost_Share"].apply(format_pct_str)
        vp_show["Secondary_Cost_Share"] = vp_show["Secondary_Cost_Share"].apply(format_pct_str)
        st.dataframe(vp_show.sort_values(["Confidence", "Vendor_Total_Cost"], ascending=[True, False]), use_container_width=True)

    st.subheader("Enhanced: Vendor Top5 Code Distribution (Evidence)")
    if vendor_top5.empty:
        st.info("No top5 distribution.")
    else:
        vt_show = vendor_top5.copy()
        vt_show["Cost_Share"] = vt_show["Cost_Share"].apply(format_pct_str)
        st.dataframe(vt_show, use_container_width=True)

    st.subheader("Enhanced: Mixed Vendor Risk List (Needs Review)")
    if vendor_mixed_risk.empty:
        st.info("No mixed vendors under current logic.")
    else:
        mv_show = vendor_mixed_risk.copy()
        mv_show["Primary_Cost_Share"] = mv_show["Primary_Cost_Share"].apply(format_pct_str)
        mv_show["Secondary_Cost_Share"] = mv_show["Secondary_Cost_Share"].apply(format_pct_str)
        st.dataframe(mv_show.sort_values("Vendor_Total_Cost", ascending=False), use_container_width=True)

    if enable_tlmf_alloc:
        st.subheader("Enhanced: TLMF AR Consolidated Lines (Detected)")
        if tlmf_ar_lines.empty:
            st.info("No TLMF AR lines detected (or no TLMF rows).")
        else:
            st.dataframe(tlmf_ar_lines[["MAWB", "Vendor", "Cost Amount", "Sell Amount"]], use_container_width=True)

        st.subheader("Enhanced: TLMF Vendor Allocation (Allocated AR → Vendor)")
        if tlmf_alloc.empty:
            st.info("No TLMF vendor allocation computed (no vendor cost lines).")
        else:
            st.dataframe(display_df(tlmf_alloc), use_container_width=True)

        st.subheader("Enhanced: TLMF Anomalies (MAWB + Vendor)")
        if tlmf_anomalies.empty:
            st.info("No TLMF anomalies detected.")
        else:
            st.dataframe(display_df(tlmf_anomalies), use_container_width=True)

    # ---------------- Export ----------------
    output = io.BytesIO()

    summary_x = to_date_only(summary, ["ETA"])
    margin_outliers_x = to_date_only(margin_outliers, ["ETA"])
    negative_profit_x = to_date_only(negative_profit, ["ETA"])
    zero_margin_x = to_date_only(zero_margin, ["ETA"])
    zero_profit_x = to_date_only(zero_profit, ["ETA"])
    exceptions_x = to_date_only(exceptions, ["ETA"])
    client_summary_x = to_date_only(client_summary, ["Latest_ETA"])
    df_x = to_date_only(df, ["ETA"])
    both_zero_x = to_date_only(both_zero, ["ETA"])
    sell_zero_only_x = to_date_only(sell_zero_only, ["ETA"])
    cost_zero_only_x = to_date_only(cost_zero_only, ["ETA"])
    chargecode_profit_le0_mawb_x = to_date_only(chargecode_profit_le0_mawb, ["ETA"])

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        workbook = writer.book
        header_fmt = workbook.add_format({"bold": True, "font_size": 14})
        subheader_fmt = workbook.add_format({"bold": True, "font_size": 12})
        bold_fmt = workbook.add_format({"bold": True})
        percent_fmt = workbook.add_format({"num_format": "0.00%"})
        number_fmt = workbook.add_format({"num_format": "#,##0.00"})

        # Analysis Summary sheet
        ws = workbook.add_worksheet("Analysis Summary")
        writer.sheets["Analysis Summary"] = ws
        ws.write(0, 0, "Analysis Summary", header_fmt)

        link_start_row = 2
        ws.write(link_start_row, 0, "This page provides an overview. Click detail links below:", bold_fmt)

        tab_links = [
            ("Open exceptions overview + detail", "Exceptions"),
            ("MAWB level summary + detail", "MAWB_Summary"),
            ("Client margin summary + detail", "Client_Summary"),
            ("Margin anomalies + detail", "Margin_Outliers"),
            ("Negative profit MAWBs + detail", "Negative_Profit"),
            ("Zero margin tickets + detail", "Zero_Margin"),
            ("Zero profit tickets + detail", "Zero_Profit"),
            ("Cost=Sell=0 tickets + detail", "Both_Zero"),
            ("Sell=0 only tickets + detail", "Sell_Zero_Only"),
            ("Cost=0 only tickets + detail", "Cost_Zero_Only"),
            ("Charge code summary + detail", "ChargeCode_Summary"),
            ("Vendor summary + detail", "Vendor_Summary"),
            ("ChargeCode Profit<=0 by MAWB + detail", "ChargeCode_ProfitLE0_MAWB"),
            ("Raw enriched billing + detail", "Raw_Billing_Enriched"),
            ("Enhanced: Vendor Primary Code + detail", "Vendor_PrimaryCode"),
            ("Enhanced: Vendor Top5 + detail", "Vendor_Top5"),
            ("Enhanced: Mixed Vendor Risk + detail", "Vendor_MixedRisk"),
        ]
        if enable_tlmf_alloc:
            tab_links += [
                ("Enhanced: TLMF All Rows + detail", "TLMF_All"),
                ("Enhanced: TLMF AR Lines + detail", "TLMF_AR_Lines"),
                ("Enhanced: TLMF Vendor Alloc + detail", "TLMF_Allocated"),
                ("Enhanced: TLMF Anomalies + detail", "TLMF_Anomalies"),
                ("Enhanced: TLMF MAWB Totals + detail", "TLMF_MAWB_Totals"),
            ]
        if mawb_keep:
            tab_links.insert(0, ("MAWB not found from filter + detail", "MAWB_Not_Found"))

        r = link_start_row + 1
        for text, sheet_name in tab_links:
            ws.write_url(r, 0, f"internal:'{sheet_name}'!A1", string=text)
            r += 1

        # KPI vertical
        kpi_row = r + 1
        ws.write(kpi_row, 0, "KPI (two-column)", subheader_fmt)
        ws.write(kpi_row + 1, 0, "Metric", bold_fmt)
        ws.write(kpi_row + 1, 1, "Value", bold_fmt)
        kpi_write_row = kpi_row + 2
        for i, (k, v) in enumerate(kpi_dict.items()):
            ws.write(kpi_write_row + i, 0, k)
            if k in {"Closed %", "Overall Profit Margin %", "ETA Filled %"}:
                ws.write_number(kpi_write_row + i, 1, float(v), percent_fmt)
            else:
                try:
                    ws.write_number(kpi_write_row + i, 1, float(v), number_fmt)
                except Exception:
                    ws.write(kpi_write_row + i, 1, str(v))

        # Write original sheets
        exceptions_x.to_excel(writer, index=False, sheet_name="Exceptions")
        summary_x.to_excel(writer, index=False, sheet_name="MAWB_Summary")
        client_summary_x.to_excel(writer, index=False, sheet_name="Client_Summary")
        margin_outliers_x.to_excel(writer, index=False, sheet_name="Margin_Outliers")
        negative_profit_x.to_excel(writer, index=False, sheet_name="Negative_Profit")
        zero_margin_x.to_excel(writer, index=False, sheet_name="Zero_Margin")
        zero_profit_x.to_excel(writer, index=False, sheet_name="Zero_Profit")
        both_zero_x.to_excel(writer, index=False, sheet_name="Both_Zero")
        sell_zero_only_x.to_excel(writer, index=False, sheet_name="Sell_Zero_Only")
        cost_zero_only_x.to_excel(writer, index=False, sheet_name="Cost_Zero_Only")
        chargecode_summary.to_excel(writer, index=False, sheet_name="ChargeCode_Summary")
        vendor_summary.to_excel(writer, index=False, sheet_name="Vendor_Summary")
        chargecode_profit_le0_mawb_x.to_excel(writer, index=False, sheet_name="ChargeCode_ProfitLE0_MAWB")
        df_x.to_excel(writer, index=False, sheet_name="Raw_Billing_Enriched")

        if mawb_keep:
            pd.DataFrame({"MAWB": mawb_not_found}).to_excel(writer, index=False, sheet_name="MAWB_Not_Found")

        # Enhanced sheets
        vendor_primary.to_excel(writer, index=False, sheet_name="Vendor_PrimaryCode")
        vendor_top5.to_excel(writer, index=False, sheet_name="Vendor_Top5")
        vendor_mixed_risk.to_excel(writer, index=False, sheet_name="Vendor_MixedRisk")

        if enable_tlmf_alloc:
            tlmf_all_rows.to_excel(writer, index=False, sheet_name="TLMF_All")
            tlmf_ar_lines.to_excel(writer, index=False, sheet_name="TLMF_AR_Lines")
            tlmf_mawb_totals.to_excel(writer, index=False, sheet_name="TLMF_MAWB_Totals")
            tlmf_alloc.to_excel(writer, index=False, sheet_name="TLMF_Allocated")
            tlmf_anomalies.to_excel(writer, index=False, sheet_name="TLMF_Anomalies")

        # percent formatting across sheets
        percent_sheets = {
            "Exceptions": exceptions_x,
            "MAWB_Summary": summary_x,
            "Client_Summary": client_summary_x,
            "Margin_Outliers": margin_outliers_x,
            "Negative_Profit": negative_profit_x,
            "Zero_Margin": zero_margin_x,
            "Zero_Profit": zero_profit_x,
            "Both_Zero": both_zero_x,
            "Sell_Zero_Only": sell_zero_only_x,
            "Cost_Zero_Only": cost_zero_only_x,
            "ChargeCode_Summary": chargecode_summary,
            "Vendor_Summary": vendor_summary,
            "ChargeCode_ProfitLE0_MAWB": chargecode_profit_le0_mawb_x,
            "Vendor_PrimaryCode": vendor_primary,
            "TLMF_Allocated": tlmf_alloc if enable_tlmf_alloc else pd.DataFrame(),
        }
        for sh, dfx in percent_sheets.items():
            if sh in writer.sheets and not dfx.empty:
                ws2 = writer.sheets[sh]
                for colname in ["Profit Margin %", "Primary_Cost_Share", "Secondary_Cost_Share", "Alloc_Share", "Vendor_Margin_Est"]:
                    if colname in dfx.columns:
                        idx = list(dfx.columns).index(colname)
                        excel_set_percent_col(ws2, idx, workbook)

    st.download_button(
        "Download Report Excel",
        data=output.getvalue(),
        file_name="MAWB_Audit_Report_Enhanced.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

except Exception as e:
    st.exception(e)
