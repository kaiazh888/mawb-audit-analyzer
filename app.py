import io
import re
import numpy as np
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

def format_pct_str_or_blank(x):
    try:
        if x is None or pd.isna(x):
            return ""
        return f"{float(x) * 100:.2f}%"
    except Exception:
        return ""

def ratio_or_nan(numer: pd.Series, denom: pd.Series) -> pd.Series:
    numer = pd.to_numeric(numer, errors="coerce")
    denom = pd.to_numeric(denom, errors="coerce").fillna(0)
    out = pd.Series(np.nan, index=denom.index, dtype="float64")
    mask = denom != 0
    out.loc[mask] = (numer.loc[mask] / denom.loc[mask]).astype("float64")
    return out

def is_na(x) -> bool:
    try:
        return x is pd.NA or pd.isna(x)
    except Exception:
        return x is None

def to_float_or_none(x):
    if is_na(x):
        return None
    try:
        return float(x)
    except Exception:
        return None

def excel_write_number_or_blank(ws, r, c, x, num_fmt=None, blank=""):
    v = to_float_or_none(x)
    if v is None:
        ws.write(r, c, blank)
    else:
        if num_fmt is None:
            ws.write_number(r, c, v)
        else:
            ws.write_number(r, c, v, num_fmt)

def excel_set_percent_col(ws, col_idx: int, workbook, width: int = 16):
    fmt = workbook.add_format({"num_format": "0.00%"})
    ws.set_column(col_idx, col_idx, width, fmt)

# =========================
# Integrated View Builders
# =========================
def build_integrated_view(df: pd.DataFrame, group_key: str) -> pd.DataFrame:
    """
    Create integrated view:
    - TOTAL row per group (Vendor/Client)
    - breakdown rows per (group, Charge Code)
    Always include AP/AR/Profit/Margin + counts.
    """
    if df.empty:
        return pd.DataFrame(columns=[group_key, "Charge Code", "AP", "AR", "Profit", "Profit Margin %", "MAWB Count", "Line Count"])

    base = df.copy()
    base["AP"] = base["Cost Amount"]
    base["AR"] = base["Sell Amount"]
    base["Profit"] = base["AR"] - base["AP"]

    # breakdown by (group, charge code)
    cc = (
        base.groupby([group_key, "Charge Code"], as_index=False)
            .agg(
                AP=("AP", "sum"),
                AR=("AR", "sum"),
                Profit=("Profit", "sum"),
                Line_Count=("MAWB", "size"),
                MAWB_Count=("MAWB", pd.Series.nunique),
            )
    )
    cc["Profit Margin %"] = ratio_or_nan(cc["Profit"], cc["AR"])
    cc = cc.rename(columns={"Line_Count": "Line Count", "MAWB_Count": "MAWB Count"})

    # total per group
    tot = (
        base.groupby(group_key, as_index=False)
            .agg(
                AP=("AP", "sum"),
                AR=("AR", "sum"),
                Profit=("Profit", "sum"),
                Line_Count=("MAWB", "size"),
                MAWB_Count=("MAWB", pd.Series.nunique),
            )
    )
    tot["Charge Code"] = "(TOTAL)"
    tot["Profit Margin %"] = ratio_or_nan(tot["Profit"], tot["AR"])
    tot = tot.rename(columns={"Line_Count": "Line Count", "MAWB_Count": "MAWB Count"})

    out = pd.concat([tot[[group_key, "Charge Code", "AP", "AR", "Profit", "Profit Margin %", "MAWB Count", "Line Count"]],
                     cc[[group_key, "Charge Code", "AP", "AR", "Profit", "Profit Margin %", "MAWB Count", "Line Count"]]],
                    ignore_index=True)

    # sort: totals first within each group
    out["_is_total"] = (out["Charge Code"] == "(TOTAL)").astype(int)
    out = out.sort_values([group_key, "_is_total", "Profit"], ascending=[True, False, False]).drop(columns=["_is_total"])
    return out

def build_audit_dashboard(summary: pd.DataFrame, exceptions: pd.DataFrame, df_lines: pd.DataFrame) -> pd.DataFrame:
    """
    Audit Control Dashboard table: KPI + exception distribution + amount by exception type.
    """
    total_mawb = int(len(summary))
    closed_cnt = int((summary["Classification"] == "Closed").sum())
    open_cnt = total_mawb - closed_cnt

    total_ar = float(summary["Total_Sell"].sum())
    total_ap = float(summary["Total_Cost"].sum())
    total_profit = float(summary["Profit"].sum())
    overall_pm = (total_profit / total_ar) if total_ar else np.nan

    # exception distribution (count + amounts)
    exc = summary.copy()
    exc["Exception_Type"] = exc["Exception_Type"].fillna("")
    exc_dist = (
        exc.groupby("Exception_Type", as_index=False)
           .agg(
               MAWB_Count=("MAWB", "count"),
               AP=("Total_Cost", "sum"),
               AR=("Total_Sell", "sum"),
               Profit=("Profit", "sum"),
           )
    )
    exc_dist["Profit Margin %"] = ratio_or_nan(exc_dist["Profit"], exc_dist["AR"])
    exc_dist = exc_dist.sort_values(["MAWB_Count", "Profit"], ascending=[False, False])

    # headline KPI block as rows
    kpi_rows = [
        ("Auditable Total MAWB", total_mawb),
        ("Auditable Closed Count", closed_cnt),
        ("Auditable Open Count", open_cnt),
        ("Auditable Total AP", total_ap),
        ("Auditable Total AR", total_ar),
        ("Auditable Total Profit", total_profit),
        ("Auditable Overall Profit Margin %", overall_pm),
        ("Open (Exceptions) MAWB Count", int(len(exceptions))),
    ]
    kpi_df = pd.DataFrame(kpi_rows, columns=["Metric", "Value"])

    # flatten exception dist with labels
    exc_dist_out = exc_dist.rename(columns={"Exception_Type": "Exception Type"})
    # Add a section marker row (for Excel readability)
    section = pd.DataFrame([{"Metric": "----- Exception Distribution (by MAWB) -----", "Value": ""}])
    # Convert exc_dist to KPI-like rows (compact)
    exc_rows = []
    for _, r in exc_dist_out.iterrows():
        label = str(r["Exception Type"]) if str(r["Exception Type"]).strip() else "(No Exception)"
        exc_rows.append({"Metric": f"{label} | MAWB Count", "Value": r["MAWB_Count"]})
        exc_rows.append({"Metric": f"{label} | AP", "Value": r["AP"]})
        exc_rows.append({"Metric": f"{label} | AR", "Value": r["AR"]})
        exc_rows.append({"Metric": f"{label} | Profit", "Value": r["Profit"]})
        exc_rows.append({"Metric": f"{label} | Profit Margin %", "Value": r["Profit Margin %"]})
    exc_kpi_df = pd.DataFrame(exc_rows)

    return pd.concat([kpi_df, section, exc_kpi_df], ignore_index=True)

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

    # line-level profit/margin (required)
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
    client_summary["Profit Margin %"] = ratio_or_nan(client_summary["Profit"], client_summary["Total_Sell"])
    client_summary = client_summary.sort_values("Profit", ascending=False)

    # ---- Other existing analysis pages ----
    margin_anomalies = summary[
        ((summary["Profit Margin %"] < 0.30) | (summary["Profit Margin %"] > 0.80)) &
        (~summary["Profit Margin %"].isna())
    ].copy().sort_values("Profit Margin %")

    negative_profit = summary[summary["Profit"] < 0].copy().sort_values("Profit")

    both_zero = summary[(summary["Total_Sell"] == 0) & (summary["Total_Cost"] == 0)].copy().sort_values("MAWB")
    sell_zero_only = summary[(summary["Total_Sell"] == 0) & (summary["Total_Cost"] > 0)].copy().sort_values("Total_Cost", ascending=False)
    cost_zero_only = summary[(summary["Total_Cost"] == 0) & (summary["Total_Sell"] > 0)].copy().sort_values("Total_Sell", ascending=False)

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
    # NEW: Integrated Views + Audit Dashboard
    # =========================
    vendor_integrated_view = build_integrated_view(df, "Vendor")
    client_integrated_view = build_integrated_view(df, "Client")
    audit_dashboard = build_audit_dashboard(summary, exceptions, df)

    # ---------------- UI ----------------
    if eta_parse_note:
        st.info(eta_parse_note)

    if mawb_keep:
        st.subheader("MAWB Not Found (in uploaded Billing file)")
        st.dataframe(pd.DataFrame({"MAWB": mawb_not_found}), use_container_width=True)

    if structural_totals is not None:
        st.subheader("Structural Client Totals (Excluded from ALL Detail Pages/Exports)")
        tmp = structural_totals.copy()
        tmp["Profit Margin %"] = tmp["Profit Margin %"].apply(format_pct_str_or_blank)
        st.dataframe(tmp, use_container_width=True)

    st.subheader("Audit Control Dashboard (NEW)")
    dash_disp = audit_dashboard.copy()
    dash_disp["Value"] = dash_disp.apply(
        lambda r: format_pct_str_or_blank(r["Value"]) if str(r["Metric"]).endswith("%") else ("" if is_na(r["Value"]) else r["Value"]),
        axis=1
    )
    st.dataframe(dash_disp, use_container_width=True)

    st.subheader("Vendor Integrated View (NEW)")
    viv = vendor_integrated_view.copy()
    viv["Profit Margin %"] = viv["Profit Margin %"].apply(format_pct_str_or_blank)
    st.dataframe(viv, use_container_width=True)

    st.subheader("Client Integrated View (NEW)")
    civ = client_integrated_view.copy()
    civ["Profit Margin %"] = civ["Profit Margin %"].apply(format_pct_str_or_blank)
    st.dataframe(civ, use_container_width=True)

    def display_df(df_in, date_cols=None):
        out = df_in.copy()
        if date_cols:
            out = to_date_only(out, date_cols)
        if "Profit Margin %" in out.columns:
            out["Profit Margin %"] = out["Profit Margin %"].apply(format_pct_str_or_blank)
        return out

    st.subheader("Exceptions (Open items) — Auditable Only")
    st.dataframe(display_df(exceptions, date_cols=["ETA"]), use_container_width=True)

    st.subheader("MAWB Summary (All) — Auditable Only")
    st.dataframe(display_df(summary, date_cols=["ETA"]), use_container_width=True)

    st.subheader("Client Summary — Auditable Only")
    st.dataframe(display_df(client_summary, date_cols=["Latest_ETA"]), use_container_width=True)

    st.subheader("Margin Anomalies — Auditable Only")
    st.dataframe(display_df(margin_anomalies, date_cols=["ETA"]), use_container_width=True)

    st.subheader("Negative Profit — Auditable Only")
    st.dataframe(display_df(negative_profit, date_cols=["ETA"]), use_container_width=True)

    st.subheader("Cost=Sell=0 — Auditable Only")
    st.dataframe(display_df(both_zero, date_cols=["ETA"]), use_container_width=True)

    st.subheader("Sell=0 ONLY — Auditable Only")
    st.dataframe(display_df(sell_zero_only, date_cols=["ETA"]), use_container_width=True)

    st.subheader("Cost=0 ONLY — Auditable Only")
    st.dataframe(display_df(cost_zero_only, date_cols=["ETA"]), use_container_width=True)

    st.subheader("Charge Code Summary — Auditable Only")
    st.dataframe(display_df(chargecode_summary), use_container_width=True)

    st.subheader("Vendor Summary — Auditable Only")
    st.dataframe(display_df(vendor_summary), use_container_width=True)

    st.subheader("Charge Code Profit <= 0 (by MAWB) — Auditable Only")
    st.dataframe(display_df(chargecode_profit_le0_mawb, date_cols=["ETA"]), use_container_width=True)

    # ---------------- Export ----------------
    output = io.BytesIO()

    summary_x = to_date_only(summary, ["ETA"])
    margin_anomalies_x = to_date_only(margin_anomalies, ["ETA"])
    negative_profit_x = to_date_only(negative_profit, ["ETA"])
    exceptions_x = to_date_only(exceptions, ["ETA"])
    client_summary_x = to_date_only(client_summary, ["Latest_ETA"])
    both_zero_x = to_date_only(both_zero, ["ETA"])
    sell_zero_only_x = to_date_only(sell_zero_only, ["ETA"])
    cost_zero_only_x = to_date_only(cost_zero_only, ["ETA"])
    chargecode_profit_le0_mawb_x = to_date_only(chargecode_profit_le0_mawb, ["ETA"])
    df_x = to_date_only(df, ["ETA"])

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        workbook = writer.book
        header_fmt = workbook.add_format({"bold": True, "font_size": 14})
        subheader_fmt = workbook.add_format({"bold": True, "font_size": 12})
        bold_fmt = workbook.add_format({"bold": True})
        percent_fmt = workbook.add_format({"num_format": "0.00%"})
        number_fmt = workbook.add_format({"num_format": "#,##0.00"})

        # ===== Analysis Summary sheet with links =====
        ws = workbook.add_worksheet("Analysis Summary")
        writer.sheets["Analysis Summary"] = ws
        ws.write(0, 0, "Analysis Summary (Auditable Only)", header_fmt)

        row = 2

        # Structural totals (PROCARESX)
        if structural_totals is not None:
            ws.write(row, 0, "Structural Client Totals (Excluded from ALL Detail Pages/Exports)", subheader_fmt)
            row += 1
            structural_totals.to_excel(writer, index=False, sheet_name="Analysis Summary", startrow=row, startcol=0)
            # format pm col
            pm_col = list(structural_totals.columns).index("Profit Margin %")
            excel_set_percent_col(ws, pm_col, workbook)
            row += len(structural_totals) + 3

        ws.write(row, 0, "Click detail links below:", bold_fmt)
        row += 1

        tab_links = [
            ("Audit Control Dashboard (NEW)", "Audit Dashboard"),
            ("Vendor Integrated View (NEW)", "Vendor Integrated View"),
            ("Client Integrated View (NEW)", "Client Integrated View"),
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

        # ===== NEW sheets =====
        audit_dashboard.to_excel(writer, index=False, sheet_name="Audit Dashboard")
        vendor_integrated_view.to_excel(writer, index=False, sheet_name="Vendor Integrated View")
        client_integrated_view.to_excel(writer, index=False, sheet_name="Client Integrated View")

        # ===== Existing sheets =====
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
        df_x.to_excel(writer, index=False, sheet_name="Raw_Billing_Enriched")

        if mawb_keep:
            pd.DataFrame({"MAWB": mawb_not_found}).to_excel(writer, index=False, sheet_name="MAWB_Not_Found")

        # ===== Percent formatting across sheets =====
        percent_targets = [
            ("Vendor Integrated View", vendor_integrated_view),
            ("Client Integrated View", client_integrated_view),
            ("MAWB_Summary", summary_x),
            ("Client_Summary", client_summary_x),
            ("Margin_Anomalies", margin_anomalies_x),
            ("Negative_Profit", negative_profit_x),
            ("Exceptions", exceptions_x),
            ("Both_Zero", both_zero_x),
            ("Sell_Zero_Only", sell_zero_only_x),
            ("Cost_Zero_Only", cost_zero_only_x),
            ("ChargeCode_Summary", chargecode_summary),
            ("Vendor_Summary", vendor_summary),
            ("ChargeCode_ProfitLE0_MAWB", chargecode_profit_le0_mawb_x),
            ("Raw_Billing_Enriched", df_x),
        ]

        for sh, dfx in percent_targets:
            if sh in writer.sheets and "Profit Margin %" in dfx.columns:
                ws2 = writer.sheets[sh]
                pm_col = list(dfx.columns).index("Profit Margin %")
                excel_set_percent_col(ws2, pm_col, workbook)

    st.download_button(
        "Download Report Excel (Includes NEW Dashboards)",
        data=output.getvalue(),
        file_name="MAWB_Audit_Report_WITH_DASHBOARDS.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

except Exception as e:
    st.exception(e)
