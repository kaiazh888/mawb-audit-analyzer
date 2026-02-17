import io
import re
import pandas as pd
import streamlit as st

st.set_page_config(page_title="MAWB Audit Analyzer", layout="wide")
st.title("MAWB Audit Analyzer (Billing-only)")
st.caption(
    "Upload Billing charges export + optional MAWB→ETA mapping file. "
    "Supports MAWB filter box, profit margin analysis, zero buckets, outliers, negative profit, "
    "and Charge Code / Vendor summaries."
)

# =========================
# Config (Audit Control)
# =========================
STRUCTURAL_CLIENTS = {"PROCARESX"}  # excluded from ALL detail pages/exports; only totals shown

# ---------------- Helpers ----------------
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
    """Robust ETA parser (text) + normalize to DATE (no time)."""
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

def normalize_mawb(x: str) -> str:
    if x is None:
        return ""
    s = str(x).strip().upper()
    if not s or s in {"NAN", "NONE"}:
        return ""

    s_alnum = re.sub(r"[^0-9A-Z]", "", s)

    # digits 11 => 3+8
    if s_alnum.isdigit() and len(s_alnum) == 11:
        return f"{s_alnum[:3]}-{s_alnum[3:]}"
    # digits 12 => last 11
    if s_alnum.isdigit() and len(s_alnum) == 12:
        s11 = s_alnum[-11:]
        if len(s11) == 11:
            return f"{s11[:3]}-{s11[3:]}"
        return s_alnum

    # keep existing hyphen format if already 3-xxxx
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

def format_pct_str_or_blank(x):
    """x is ratio (0..1). Return blank for NaN/NA."""
    try:
        if x is None or x is pd.NA or pd.isna(x):
            return ""
        return f"{float(x) * 100:.2f}%"
    except Exception:
        return ""

def ratio_or_nan(numer: pd.Series, denom: pd.Series) -> pd.Series:
    """Return numer/denom, but NaN when denom==0 (so margin displays blank)."""
    denom0 = denom.fillna(0)
    out = pd.Series([pd.NA] * len(denom0), index=denom0.index, dtype="float64")
    mask = denom0 != 0
    out.loc[mask] = (numer.loc[mask] / denom0.loc[mask]).astype("float64")
    return out

# ---- Robust NA handling to prevent float(pd.NA) crashes ----
def is_na(x) -> bool:
    try:
        return x is pd.NA or pd.isna(x)
    except Exception:
        return x is None

def to_float_or_none(x):
    """Convert to float if possible; return None for NA/blank."""
    if is_na(x):
        return None
    try:
        return float(x)
    except Exception:
        return None

def excel_write_number_or_blank(ws, r, c, x, num_fmt=None, blank=""):
    """Write numeric if valid, else blank."""
    v = to_float_or_none(x)
    if v is None:
        ws.write(r, c, blank)
    else:
        if num_fmt is None:
            ws.write_number(r, c, v)
        else:
            ws.write_number(r, c, v, num_fmt)

# Excel formatting helpers
def excel_set_percent_col(ws, col_idx: int, workbook, width: int = 16):
    fmt = workbook.add_format({"num_format": "0.00%"})
    ws.set_column(col_idx, col_idx, width, fmt)

# ---------------- Uploaders ----------------
billing_file = st.file_uploader("Upload Billing Charges Excel (.xlsx)", type=["xlsx"], key="billing")
eta_file = st.file_uploader(
    "Optional: Upload MAWB→ETA mapping Excel (.xlsx)",
    type=["xlsx"],
    key="eta_mapping"
)

st.divider()
st.subheader("Optional Filter: Keep only specified MAWBs")
mawb_text = st.text_area(
    "Paste MAWBs here (comma / space / newline separated). Supports 99934022122 → 999-34022122. Leave blank to keep all.",
    height=140,
    placeholder="Example:\n999-34022122\n99934022133\n999 34022144"
)

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

    # ---- Normalize billing ----
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

    df_all = df_all[df_all["MAWB"].ne("")].copy()

    # Add line-level Profit/Margin (always alongside Sell/Cost)
    df_all["Profit"] = df_all["Sell Amount"] - df_all["Cost Amount"]
    df_all["Profit Margin %"] = ratio_or_nan(df_all["Profit"], df_all["Sell Amount"])

    # ---- Optional MAWB filter ----
    mawb_keep = parse_mawb_list(mawb_text)
    if mawb_keep:
        before_rows = len(df_all)
        before_mawb = df_all["MAWB"].nunique()

        df_all = df_all[df_all["MAWB"].isin(mawb_keep)].copy()

        after_rows = len(df_all)
        after_mawb = df_all["MAWB"].nunique()

        found_set = set(df_all["MAWB"].unique())
        mawb_not_found = sorted(set(mawb_keep) - found_set)

        st.info(
            f"MAWB filter applied: rows {before_rows} → {after_rows}, "
            f"unique MAWB {before_mawb} → {after_mawb}."
        )
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
                    eta_parse_note = (
                        f"ETA parsing note: {bad_eta_rows} / {total_rows} ETA values could not be parsed and were left blank."
                    )

                eta_map = (
                    mdf.dropna(subset=["MAWB"])
                       .groupby("MAWB", as_index=False)["ETA"]
                       .max()
                )

    # ---- Merge ETA into billing ----
    if eta_map is not None and not eta_map.empty:
        df_all = df_all.merge(eta_map, on="MAWB", how="left")
    else:
        df_all["ETA"] = pd.NaT
    df_all["ETA"] = pd.to_datetime(df_all["ETA"], errors="coerce").dt.normalize()

    # =========================
    # Structural client totals (PROCARESX) — ONLY totals shown
    # =========================
    structural_df = df_all[df_all["Client"].isin(STRUCTURAL_CLIENTS)].copy()
    structural_totals = None
    if not structural_df.empty:
        s_cost = float(structural_df["Cost Amount"].sum())
        s_sell = float(structural_df["Sell Amount"].sum())
        s_profit = s_sell - s_cost
        s_margin = (s_profit / s_sell) if s_sell else None  # show blank when AR=0
        structural_totals = {
            "Client": "PROCARESX",
            "MAWB Count": int(structural_df["MAWB"].nunique()),
            "Total AP": s_cost,
            "Total AR": s_sell,
            "Total Profit": s_profit,
            "Profit Margin %": s_margin,
            "Note": "Structural allocation; excluded from ALL audit detail pages/exports"
        }

    # =========================
    # Auditable dataset (exclude structural clients everywhere)
    # =========================
    df = df_all[~df_all["Client"].isin(STRUCTURAL_CLIENTS)].copy()

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
    summary["Profit Margin %"] = ratio_or_nan(summary["Profit"], summary["Total_Sell"])

    # Classification & Exception Types (auditable only)
    def classification(r):
        if not (r["Total_Cost"] > 0 and r["Total_Sell"] > 0):
            return "Open"
        pm = r["Profit Margin %"]
        if pd.isna(pm):
            return "Open"
        if (pm < 0.30) or (pm > 0.80):
            return "Open"
        return "Closed"

    def exception_type(r):
        if r["Total_Cost"] == 0 and r["Total_Sell"] == 0:
            return "Cost=Sell=0"
        if r["Total_Sell"] == 0 and r["Total_Cost"] > 0:
            return "Revenue=0"
        if r["Total_Cost"] == 0 and r["Total_Sell"] > 0:
            return "Cost=0"
        pm = r["Profit Margin %"]
        if not pd.isna(pm):
            if pm < 0.30:
                return "Margin<30%"
            if pm > 0.80:
                return "Margin>80%"
        return ""

    summary["Classification"] = summary.apply(classification, axis=1)
    summary["Exception_Type"] = summary.apply(exception_type, axis=1)

    exceptions = summary[summary["Classification"].eq("Open")].copy()

    # ---- Client Summary (auditable clients only) ----
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
    client_summary["Profit Margin %"] = ratio_or_nan(client_summary["Profit"], client_summary["Total_Sell"])
    client_summary = client_summary.sort_values("Profit", ascending=False)

    # ---- Margin Anomalies / Negative Profit ----
    margin_anomalies = summary[
        ((summary["Profit Margin %"] < 0.30) | (summary["Profit Margin %"] > 0.80)) &
        (~summary["Profit Margin %"].isna())
    ].copy().sort_values("Profit Margin %")

    negative_profit = summary[summary["Profit"] < 0].copy().sort_values("Profit")

    # ---- Zero buckets ----
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
    chargecode_summary["Profit Margin %"] = ratio_or_nan(chargecode_summary["Profit"], chargecode_summary["Total_Sell"])
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
    vendor_summary["Profit Margin %"] = ratio_or_nan(vendor_summary["Profit"], vendor_summary["Total_Sell"])
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
    cc_mawb["Profit Margin %"] = ratio_or_nan(cc_mawb["Profit"], cc_mawb["Total_Sell"])
    cc_mawb["ETA Month"] = pd.to_datetime(cc_mawb["ETA"], errors="coerce").dt.to_period("M").astype(str).replace("NaT", "")

    chargecode_profit_le0_mawb = cc_mawb[cc_mawb["Profit"] <= 0].copy().sort_values(
        ["Profit", "Total_Sell"], ascending=[True, False]
    )

    # =========================
    # KPI (Auditable only) + Structural Totals block
    # =========================
    total_mawb = int(len(summary))
    closed_cnt = int((summary["Classification"] == "Closed").sum())
    open_cnt = total_mawb - closed_cnt

    total_sell_sum = float(summary["Total_Sell"].sum())
    total_cost_sum = float(summary["Total_Cost"].sum())
    total_profit_sum = float(summary["Profit"].sum())
    overall_pm = (total_profit_sum / total_sell_sum) if total_sell_sum else None  # blank if AR=0

    neg_profit_cnt = int((summary["Profit"] < 0).sum())
    neg_profit_amt = float(summary.loc[summary["Profit"] < 0, "Profit"].sum())
    neg_profit_ratio = (neg_profit_cnt / total_mawb) if total_mawb else 0

    eta_filled_ratio = float((summary["ETA"].notna().sum() / total_mawb)) if total_mawb else 0

    kpi_dict = {
        "Auditable Total MAWB": total_mawb,
        "Auditable Closed Count": closed_cnt,
        "Auditable Closed %": (closed_cnt / total_mawb) if total_mawb else 0,
        "Auditable Open Count": open_cnt,
        "Revenue=0 Count": int((summary["Exception_Type"] == "Revenue=0").sum()),
        "Cost=0 Count": int((summary["Exception_Type"] == "Cost=0").sum()),
        "Cost=Sell=0 Count": int((summary["Exception_Type"] == "Cost=Sell=0").sum()),
        "Margin<30% Count": int((summary["Exception_Type"] == "Margin<30%").sum()),
        "Margin>80% Count": int((summary["Exception_Type"] == "Margin>80%").sum()),
        "Auditable Total AP": total_cost_sum,
        "Auditable Total AR": total_sell_sum,
        "Auditable Total Profit": total_profit_sum,
        "Auditable Overall Profit Margin %": overall_pm,
        "ETA Filled %": eta_filled_ratio,
    }
    KPI_PCT_KEYS = {"Auditable Closed %", "Auditable Overall Profit Margin %", "ETA Filled %"}

    # ---------------- UI ----------------
    if eta_parse_note:
        st.info(eta_parse_note)

    if mawb_keep:
        st.subheader("MAWB Not Found (in uploaded Billing file)")
        st.dataframe(pd.DataFrame({"MAWB": mawb_not_found}), use_container_width=True)

    # Structural totals block (PROCARESX only totals)
    if structural_totals:
        st.subheader("Structural Client Totals (Excluded from ALL Audit Details)")
        tmp = pd.DataFrame([structural_totals])
        tmp_display = tmp.copy()
        tmp_display["Profit Margin %"] = tmp_display["Profit Margin %"].apply(format_pct_str_or_blank)
        st.dataframe(tmp_display, use_container_width=True)

    st.subheader("Analysis Summary (KPI) — Auditable Clients Only")
    kpi_rows = []
    for k, v in kpi_dict.items():
        if k in KPI_PCT_KEYS:
            kpi_rows.append({"Metric": k, "Value": format_pct_str_or_blank(v)})
        else:
            kpi_rows.append({"Metric": k, "Value": "" if is_na(v) else v})
    st.dataframe(pd.DataFrame(kpi_rows), use_container_width=True)

    st.subheader("Summary: Profit < 0 (Auditable MAWBs)")
    st.dataframe(pd.DataFrame([
        {"Metric": "Profit < 0 Count", "Value": neg_profit_cnt},
        {"Metric": "Profit < 0 Total Amount", "Value": neg_profit_amt},
        {"Metric": "Profit < 0 % of MAWBs", "Value": format_pct_str_or_blank(neg_profit_ratio)},
    ]), use_container_width=True)

    def display_df(df_in, date_cols=None):
        out = df_in.copy()
        if date_cols:
            out = to_date_only(out, date_cols)
        if "Profit Margin %" in out.columns:
            out["Profit Margin %"] = out["Profit Margin %"].apply(format_pct_str_or_blank)
        if "ETA Filled %" in out.columns:
            out["ETA Filled %"] = out["ETA Filled %"].apply(format_pct_str_or_blank)
        if "Auditable Closed %" in out.columns:
            out["Auditable Closed %"] = out["Auditable Closed %"].apply(format_pct_str_or_blank)
        if "Auditable Overall Profit Margin %" in out.columns:
            out["Auditable Overall Profit Margin %"] = out["Auditable Overall Profit Margin %"].apply(format_pct_str_or_blank)
        return out

    st.subheader("Exceptions (Open items) — Auditable Only")
    st.dataframe(display_df(exceptions, date_cols=["ETA"]), use_container_width=True)

    st.subheader("MAWB Summary (All) — Auditable Only")
    st.dataframe(display_df(summary, date_cols=["ETA"]), use_container_width=True)

    st.subheader("Client Profit Summary — Auditable Only")
    st.dataframe(display_df(client_summary, date_cols=["Latest_ETA"]), use_container_width=True)

    st.subheader("Margin Anomalies (PM<30% or >80%) — Auditable Only")
    st.dataframe(display_df(margin_anomalies, date_cols=["ETA"]), use_container_width=True)

    st.subheader("Negative Profit (Profit < 0) — Auditable Only")
    st.dataframe(display_df(negative_profit, date_cols=["ETA"]), use_container_width=True)

    st.subheader("Cost=Sell=0 (Both Zero) — Auditable Only")
    st.dataframe(display_df(both_zero, date_cols=["ETA"]), use_container_width=True)

    st.subheader("Sell=0 ONLY (Total_Sell=0 and Total_Cost>0) — Auditable Only")
    st.dataframe(display_df(sell_zero_only, date_cols=["ETA"]), use_container_width=True)

    st.subheader("Cost=0 ONLY (Total_Cost=0 and Total_Sell>0) — Auditable Only")
    st.dataframe(display_df(cost_zero_only, date_cols=["ETA"]), use_container_width=True)

    st.subheader("Charge Code Summary — Auditable Only")
    st.dataframe(display_df(chargecode_summary), use_container_width=True)

    st.subheader("Vendor Summary — Auditable Only")
    st.dataframe(display_df(vendor_summary), use_container_width=True)

    st.subheader("Charge Code Profit <= 0 (by MAWB) — Auditable Only")
    st.dataframe(display_df(chargecode_profit_le0_mawb, date_cols=["ETA"]), use_container_width=True)

    # ---------------- Export ----------------
    output = io.BytesIO()

    # For export: convert ETA columns to date
    summary_x = to_date_only(summary, ["ETA"])
    margin_anomalies_x = to_date_only(margin_anomalies, ["ETA"])
    negative_profit_x = to_date_only(negative_profit, ["ETA"])
    exceptions_x = to_date_only(exceptions, ["ETA"])
    client_summary_x = to_date_only(client_summary, ["Latest_ETA"])
    both_zero_x = to_date_only(both_zero, ["ETA"])
    sell_zero_only_x = to_date_only(sell_zero_only, ["ETA"])
    cost_zero_only_x = to_date_only(cost_zero_only, ["ETA"])
    chargecode_profit_le0_mawb_x = to_date_only(chargecode_profit_le0_mawb, ["ETA"])
    df_x = to_date_only(df, ["ETA"])  # Raw enriched (auditable only)

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        workbook = writer.book
        header_fmt = workbook.add_format({"bold": True, "font_size": 14})
        subheader_fmt = workbook.add_format({"bold": True, "font_size": 12})
        bold_fmt = workbook.add_format({"bold": True})
        percent_fmt = workbook.add_format({"num_format": "0.00%"})
        number_fmt = workbook.add_format({"num_format": "#,##0.00"})

        # ✅ Summary sheet
        ws = workbook.add_worksheet("Analysis Summary")
        writer.sheets["Analysis Summary"] = ws
        ws.write(0, 0, "Analysis Summary (Auditable Only)", header_fmt)

        row = 2
        # Structural totals (if any)
        if structural_totals:
            ws.write(row, 0, "Structural Client Totals (Excluded from ALL Audit Details)", subheader_fmt)
            row += 1
            headers = ["Client", "MAWB Count", "Total AP", "Total AR", "Total Profit", "Profit Margin %", "Note"]
            for j, h in enumerate(headers):
                ws.write(row, j, h, bold_fmt)
            row += 1

            ws.write(row, 0, structural_totals["Client"])
            excel_write_number_or_blank(ws, row, 1, structural_totals["MAWB Count"], number_fmt)
            excel_write_number_or_blank(ws, row, 2, structural_totals["Total AP"], number_fmt)
            excel_write_number_or_blank(ws, row, 3, structural_totals["Total AR"], number_fmt)
            excel_write_number_or_blank(ws, row, 4, structural_totals["Total Profit"], number_fmt)
            excel_write_number_or_blank(ws, row, 5, structural_totals["Profit Margin %"], percent_fmt)
            ws.write(row, 6, structural_totals["Note"])
            row += 3

        # Links
        ws.write(row, 0, "Click detail links below:", bold_fmt)
        row += 1
        tab_links = [
            ("Exceptions (Open) — detail", "Exceptions"),
            ("MAWB Summary — detail", "MAWB_Summary"),
            ("Client Summary — detail", "Client_Summary"),
            ("Margin Anomalies — detail", "Margin_Anomalies"),
            ("Negative Profit — detail", "Negative_Profit"),
            ("Cost=Sell=0 — detail", "Both_Zero"),
            ("Sell=0 only — detail", "Sell_Zero_Only"),
            ("Cost=0 only — detail", "Cost_Zero_Only"),
            ("Charge code summary — detail", "ChargeCode_Summary"),
            ("Vendor summary — detail", "Vendor_Summary"),
            ("ChargeCode Profit<=0 by MAWB — detail", "ChargeCode_ProfitLE0_MAWB"),
            ("Raw enriched billing (auditable only) — detail", "Raw_Billing_Enriched"),
        ]
        if mawb_keep:
            tab_links.insert(0, ("MAWB not found from filter — detail", "MAWB_Not_Found"))

        for text, sheet_name in tab_links:
            ws.write_url(row, 0, f"internal:'{sheet_name}'!A1", string=text)
            row += 1

        # KPI table
        row += 1
        ws.write(row, 0, "KPI — Auditable Only", subheader_fmt)
        row += 1
        ws.write(row, 0, "Metric", bold_fmt)
        ws.write(row, 1, "Value", bold_fmt)
        row += 1

        for k, v in kpi_dict.items():
            ws.write(row, 0, k)
            if k in KPI_PCT_KEYS:
                excel_write_number_or_blank(ws, row, 1, v, percent_fmt)
            else:
                excel_write_number_or_blank(ws, row, 1, v, number_fmt)
            row += 1

        # ---- Other sheets (ALL auditable only; PROCARESX excluded) ----
        exceptions_x.to_excel(writer, index=False, sheet_name="Exceptions")
        summary_x.to_excel(writer, index=False, sheet_name="MAWB_Summary")
        client_summary_x.to_excel(writer, index=False, sheet_name="Client_Summary")

        margin_anomalies_x.to_excel(writer, index=False, sheet_name="Margin_Anomalies")
        negative_profit_x.to_excel(writer, index=False, sheet_name="Negative_Profit")

        both_zero_x.to_excel(writer, index=False, sheet_name="Both_Zero")
        sell_zero_only_x.to_excel(writer, index=False, sheet_name="Sell_Zero_Only")
        cost_zero_only_x.to_excel(writer, index=False, sheet_name="Cost_Zero_Only")

        chargecode_summary.to_excel(writer, index=False, sheet_name="ChargeCode_Summary")
        vendor_summary.to_excel(writer, index=False, sheet_name="Vendor_Summary")
        chargecode_profit_le0_mawb_x.to_excel(writer, index=False, sheet_name="ChargeCode_ProfitLE0_MAWB")

        if mawb_keep:
            pd.DataFrame({"MAWB": mawb_not_found}).to_excel(writer, index=False, sheet_name="MAWB_Not_Found")

        df_x.to_excel(writer, index=False, sheet_name="Raw_Billing_Enriched")

        # ---- Format % columns to show % symbol in ALL sheets ----
        percent_sheets = {
            "Exceptions": exceptions_x,
            "MAWB_Summary": summary_x,
            "Client_Summary": client_summary_x,
            "Margin_Anomalies": margin_anomalies_x,
            "Negative_Profit": negative_profit_x,
            "Both_Zero": both_zero_x,
            "Sell_Zero_Only": sell_zero_only_x,
            "Cost_Zero_Only": cost_zero_only_x,
            "ChargeCode_Summary": chargecode_summary,
            "Vendor_Summary": vendor_summary,
            "ChargeCode_ProfitLE0_MAWB": chargecode_profit_le0_mawb_x,
            "Raw_Billing_Enriched": df_x,
        }

        for sh, dfx in percent_sheets.items():
            if sh in writer.sheets and "Profit Margin %" in dfx.columns:
                ws2 = writer.sheets[sh]
                pm_col = list(dfx.columns).index("Profit Margin %")
                excel_set_percent_col(ws2, pm_col, workbook)

    st.download_button(
        "Download Report Excel (Auditable Only + Structural Totals)",
        data=output.getvalue(),
        file_name="MAWB_Audit_Report_AuditableOnly_PLUS_StructuralTotals.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

except Exception as e:
    st.exception(e)
