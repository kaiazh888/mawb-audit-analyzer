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
    """
    Robust ETA parser (text) + normalize to DATE (no time).
    """
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

    # ✅ keep DATE only (strip time)
    dt1 = dt1.dt.normalize()
    return dt1

def pct(numer: pd.Series, denom: pd.Series) -> pd.Series:
    return (numer / denom).where(denom != 0, 0)

def normalize_mawb(x: str) -> str:
    """
    Normalize MAWB to common format:
    - trims spaces
    - removes non-alphanumeric (keeps digits/letters)
    - if looks like 3-digit prefix + 8-digit serial (e.g., 99934022122), convert to 999-34022122
    """
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
    """
    For export/UI display: convert datetime column(s) to pure date object (YYYY-MM-DD in Excel),
    avoiding 00:00:00 display.
    """
    df_out = df_in.copy()
    for c in cols:
        if c in df_out.columns:
            df_out[c] = pd.to_datetime(df_out[c], errors="coerce").dt.date
    return df_out

# ---------------- Uploaders ----------------
billing_file = st.file_uploader("Upload Billing Charges Excel (.xlsx)", type=["xlsx"], key="billing")
eta_file = st.file_uploader(
    "Optional: Upload MAWB→ETA mapping Excel (.xlsx) (e.g., ARAP-MAWB-ChargeType-Overview...)",
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

    # Normalize billing
    df = raw_df.copy()
    df["MAWB"] = df[mawb_col].apply(normalize_mawb)
    df["Cost Amount"] = safe_numeric(df[cost_col])
    df["Sell Amount"] = safe_numeric(df[sell_col])

    df["Client"] = df[client_col].astype(str).str.strip() if client_col else "UNKNOWN"
    df.loc[df["Client"].isin(["", "nan", "None"]), "Client"] = "UNKNOWN"

    df["Charge Code"] = df[charge_code_col].astype(str).str.strip() if charge_code_col else "UNKNOWN"
    df.loc[df["Charge Code"].isin(["", "nan", "None"]), "Charge Code"] = "UNKNOWN"

    df["Vendor"] = df[vendor_col].astype(str).str.strip() if vendor_col else "UNKNOWN"
    df.loc[df["Vendor"].isin(["", "nan", "None"]), "Vendor"] = "UNKNOWN"

    # Drop blank MAWB
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

                # ✅ parse + normalize to DATE only
                mdf["ETA"] = clean_eta_series(mdf["ETA"])

                bad_eta_rows = int(mdf["ETA"].isna().sum())
                total_rows = int(len(mdf))
                if total_rows > 0 and bad_eta_rows > 0:
                    eta_parse_note = (
                        f"ETA parsing note: {bad_eta_rows} / {total_rows} ETA values could not be parsed and were left blank."
                    )

                # Same MAWB multiple rows: latest ETA (max) (date-only already)
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

    # ✅ Ensure ETA in df is DATE-only as well (even if came with time)
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

    # Profit & margin
    summary["Profit"] = summary["Total_Sell"] - summary["Total_Cost"]
    summary["Profit Margin %"] = pct(summary["Profit"], summary["Total_Sell"])

    # ✅ Classification / Exceptions (PM must be in 30%~80% to be Closed)
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

        # ✅ New: Margin out of range (exclude pm=0 to avoid Sell=0 impact)
        pm = r["Profit Margin %"]
        if (pm != 0) and ((pm < 0.30) or (pm > 0.80)):
            return "Margin_Out_of_Range"

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

    # ---- Zero buckets (Sell/Cost) ----
    sell_zero = summary[summary["Total_Sell"] == 0].copy().sort_values("Total_Cost", ascending=False)
    cost_zero = summary[summary["Total_Cost"] == 0].copy().sort_values("Total_Sell", ascending=False)
    both_zero = summary[(summary["Total_Sell"] == 0) & (summary["Total_Cost"] == 0)].copy().sort_values("MAWB")

    # ---- Mutually exclusive zero buckets ----
    sell_zero_only = summary[(summary["Total_Sell"] == 0) & (summary["Total_Cost"] > 0)].copy().sort_values("Total_Cost", ascending=False)
    cost_zero_only = summary[(summary["Total_Cost"] == 0) & (summary["Total_Sell"] > 0)].copy().sort_values("Total_Sell", ascending=False)

    # ---- D: Charge Code Summary ----
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

    # ---- D: Vendor Summary ----
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

    # ✅ NEW: Charge Code Profit <= 0 by MAWB (tab)
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

    chargecode_profit_le0_mawb = cc_mawb[cc_mawb["Profit"] <= 0].copy()
    chargecode_profit_le0_mawb = chargecode_profit_le0_mawb.sort_values(
        ["Profit", "Total_Sell"], ascending=[True, False]
    )

    # ---- KPI ----
    total_mawb = len(summary)
    closed_cnt = int((summary["Classification"] == "Closed").sum())
    open_cnt = total_mawb - closed_cnt
    total_sell_sum = float(summary["Total_Sell"].sum())
    total_profit_sum = float(summary["Profit"].sum())

    kpi = pd.DataFrame([{
        "Total MAWB": total_mawb,
        "Closed Count": closed_cnt,
        "Closed %": (closed_cnt / total_mawb) if total_mawb else 0,
        "Open Count": open_cnt,
        "Revenue=0 Count": int((summary["Exception_Type"] == "Revenue=0").sum()),
        "Cost=0 Count": int((summary["Exception_Type"] == "Cost=0").sum()),
        "Cost=Sell=0 Count": int((summary["Exception_Type"] == "Cost=Sell=0").sum()),
        "Margin_Out_of_Range Count": int((summary["Exception_Type"] == "Margin_Out_of_Range").sum()),
        "Total Cost": float(summary["Total_Cost"].sum()),
        "Total Sell": total_sell_sum,
        "Total Profit": total_profit_sum,
        "Overall Profit Margin %": (total_profit_sum / total_sell_sum) if total_sell_sum else 0,
        "ETA Filled %": float((summary["ETA"].notna().sum() / total_mawb)) if total_mawb else 0,
    }])

    # ---------------- UI ----------------
    if eta_parse_note:
        st.info(eta_parse_note)

    if mawb_keep:
        st.subheader("MAWB Not Found (in uploaded Billing file)")
        st.dataframe(pd.DataFrame({"MAWB": mawb_not_found}), use_container_width=True)

    c1, c2 = st.columns([1, 1])
    with c1:
        st.subheader("KPI Summary")
        st.dataframe(kpi, use_container_width=True)
    with c2:
        st.subheader("Exceptions (Open items)")
        st.dataframe(exceptions, use_container_width=True)

    st.subheader("MAWB Summary (All)")
    st.dataframe(to_date_only(summary, ["ETA"]), use_container_width=True)

    st.subheader("Client Profit Summary")
    st.dataframe(to_date_only(client_summary, ["Latest_ETA"]), use_container_width=True)

    st.subheader("Profit Margin Outliers (PM<30% or PM>80%, PM!=0)")
    st.dataframe(to_date_only(margin_outliers, ["ETA"]), use_container_width=True)

    st.subheader("Negative Profit (Profit < 0)")
    st.dataframe(to_date_only(negative_profit, ["ETA"]), use_container_width=True)

    st.subheader("Zero Margin (Profit Margin % = 0)")
    st.dataframe(to_date_only(zero_margin, ["ETA"]), use_container_width=True)

    st.subheader("Zero Profit (Profit = 0)")
    st.dataframe(to_date_only(zero_profit, ["ETA"]), use_container_width=True)

    st.subheader("Sell = 0 (Total_Sell = 0)")
    st.dataframe(to_date_only(sell_zero, ["ETA"]), use_container_width=True)

    st.subheader("Cost = 0 (Total_Cost = 0)")
    st.dataframe(to_date_only(cost_zero, ["ETA"]), use_container_width=True)

    st.subheader("Both Zero (Total_Sell=0 and Total_Cost=0)")
    st.dataframe(to_date_only(both_zero, ["ETA"]), use_container_width=True)

    st.subheader("Sell=0 ONLY (Total_Sell=0 and Total_Cost>0)")
    st.dataframe(to_date_only(sell_zero_only, ["ETA"]), use_container_width=True)

    st.subheader("Cost=0 ONLY (Total_Cost=0 and Total_Sell>0)")
    st.dataframe(to_date_only(cost_zero_only, ["ETA"]), use_container_width=True)

    st.subheader("Charge Code Summary (D)")
    st.dataframe(chargecode_summary, use_container_width=True)

    st.subheader("Vendor Summary (D)")
    st.dataframe(vendor_summary, use_container_width=True)

    # ✅ NEW TAB UI
    st.subheader("Charge Code Profit <= 0 (by MAWB)")
    st.dataframe(to_date_only(chargecode_profit_le0_mawb, ["ETA"]), use_container_width=True)

    # ---------------- Export ----------------
    output = io.BytesIO()

    # For export: convert ETA columns to DATE (no time display)
    summary_x = to_date_only(summary, ["ETA"])
    margin_outliers_x = to_date_only(margin_outliers, ["ETA"])
    negative_profit_x = to_date_only(negative_profit, ["ETA"])
    zero_margin_x = to_date_only(zero_margin, ["ETA"])
    zero_profit_x = to_date_only(zero_profit, ["ETA"])
    sell_zero_x = to_date_only(sell_zero, ["ETA"])
    cost_zero_x = to_date_only(cost_zero, ["ETA"])
    both_zero_x = to_date_only(both_zero, ["ETA"])
    sell_zero_only_x = to_date_only(sell_zero_only, ["ETA"])
    cost_zero_only_x = to_date_only(cost_zero_only, ["ETA"])
    exceptions_x = to_date_only(exceptions, ["ETA"])
    client_summary_x = to_date_only(client_summary, ["Latest_ETA"])
    df_x = to_date_only(df, ["ETA"])
    chargecode_profit_le0_mawb_x = to_date_only(chargecode_profit_le0_mawb, ["ETA"])

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        kpi.to_excel(writer, index=False, sheet_name="KPI")
        summary_x.to_excel(writer, index=False, sheet_name="MAWB_Summary")
        client_summary_x.to_excel(writer, index=False, sheet_name="Client_Summary")
        exceptions_x.to_excel(writer, index=False, sheet_name="Exceptions")

        margin_outliers_x.to_excel(writer, index=False, sheet_name="Margin_Outliers")
        negative_profit_x.to_excel(writer, index=False, sheet_name="Negative_Profit")

        zero_margin_x.to_excel(writer, index=False, sheet_name="Zero_Margin")
        zero_profit_x.to_excel(writer, index=False, sheet_name="Zero_Profit")
        sell_zero_x.to_excel(writer, index=False, sheet_name="Sell_Zero")
        cost_zero_x.to_excel(writer, index=False, sheet_name="Cost_Zero")
        both_zero_x.to_excel(writer, index=False, sheet_name="Both_Zero")
        sell_zero_only_x.to_excel(writer, index=False, sheet_name="Sell_Zero_Only")
        cost_zero_only_x.to_excel(writer, index=False, sheet_name="Cost_Zero_Only")

        chargecode_summary.to_excel(writer, index=False, sheet_name="ChargeCode_Summary")
        vendor_summary.to_excel(writer, index=False, sheet_name="Vendor_Summary")

        # ✅ NEW SHEET
        chargecode_profit_le0_mawb_x.to_excel(writer, index=False, sheet_name="ChargeCode_ProfitLE0_MAWB")

        if mawb_keep:
            pd.DataFrame({"MAWB": mawb_not_found}).to_excel(writer, index=False, sheet_name="MAWB_Not_Found")

        df_x.to_excel(writer, index=False, sheet_name="Raw_Billing_Enriched")

    st.download_button(
        "Download Report Excel",
        data=output.getvalue(),
        file_name="MAWB_Audit_Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

except Exception as e:
    st.exception(e)
