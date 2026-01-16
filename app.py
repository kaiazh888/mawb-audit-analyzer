import io
import re
import pandas as pd
import streamlit as st

st.set_page_config(page_title="MAWB Audit Analyzer", layout="wide")
st.title("MAWB Audit Analyzer (Billing-only)")
st.caption(
    "Upload Billing charges export + (optional) MAWB→ETA mapping file. "
    "Generates KPI, MAWB profit/margin, Client profit/margin, exceptions, and an exportable report."
)

# ---------------- Helpers ----------------
def safe_numeric(s: pd.Series) -> pd.Series:
    return pd.to_numeric(s, errors="coerce").fillna(0)

def norm_colname(s: str) -> str:
    """Normalize column name for matching (lowercase, remove spaces/underscores/dashes)."""
    return re.sub(r"[\s_\-]+", "", str(s).strip().lower())

def find_first_col(df: pd.DataFrame, candidates: list[str]) -> str:
    """
    Return the actual column name in df that matches any candidate (case/space-insensitive).
    """
    mapping = {norm_colname(c): c for c in df.columns.astype(str)}
    for cand in candidates:
        key = norm_colname(cand)
        if key in mapping:
            return mapping[key]
    return ""

def find_sheet_with_required_cols(xls: pd.ExcelFile, required_candidates: dict) -> str:
    """
    Scan sheets and return first sheet that contains ALL required columns (by candidates).
    required_candidates example:
      {"MAWB": [...], "Cost Amount": [...], "Sell Amount": [...]}
    """
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
    Robust ETA parser for messy text:
    - strips prefixes like 'ETA:' or extra spaces
    - supports YYYYMMDD
    - supports common separators: / - .
    - supports strings with time parts
    Returns datetime64[ns] with NaT for unparseable values.
    """
    s = s.astype(str).fillna("").str.strip()

    # remove obvious prefixes like "ETA:" / "ETA -"
    s = s.str.replace(r"(?i)^\s*eta\s*[:\-]\s*", "", regex=True)

    # normalize multiple spaces
    s = s.str.replace(r"\s+", " ", regex=True)

    # handle YYYYMMDD (8 digits)
    yyyymmdd = s.str.match(r"^\d{8}$")
    s2 = s.copy()
    if yyyymmdd.any():
        parsed_yyyymmdd = pd.to_datetime(s.loc[yyyymmdd], format="%Y%m%d", errors="coerce")
        s2.loc[yyyymmdd] = parsed_yyyymmdd.astype("datetime64[ns]").astype(str)

    # First pass: general parse
    dt1 = pd.to_datetime(s2, errors="coerce", infer_datetime_format=True)

    # Second pass: try dayfirst=True where first failed (dd-mm-yyyy)
    mask = dt1.isna() & s2.ne("")
    if mask.any():
        dt2 = pd.to_datetime(s2[mask], errors="coerce", dayfirst=True, infer_datetime_format=True)
        dt1.loc[mask] = dt2

    return dt1

def pct(numer: pd.Series, denom: pd.Series) -> pd.Series:
    """Safe percent: numer/denom, denom=0 -> 0"""
    return (numer / denom).where(denom != 0, 0)

# ---------------- Uploaders ----------------
billing_file = st.file_uploader("Upload Billing Charges Excel (.xlsx)", type=["xlsx"], key="billing")
eta_file = st.file_uploader(
    "Optional: Upload MAWB→ETA mapping Excel (.xlsx) (e.g., ARAP-MAWB-ChargeType-Overview...)",
    type=["xlsx"],
    key="eta_mapping"
)

# ---------------- Config: column candidates ----------------
BILLING_REQUIRED = {
    "MAWB": ["MAWB", "Mawb", "Master AWB", "MasterAWB"],
    "Cost Amount": ["Cost Amount", "Cost", "AP Amount", "Total Cost", "CostAmount"],
    "Sell Amount": ["Sell Amount", "Sell", "AR Amount", "Total Sell", "SellAmount"],
}

# Client is optional but strongly preferred
BILLING_OPTIONAL = {
    "Client": ["Client", "Customer", "Account", "Shipper", "Bill To", "Billed To"],
}

ETA_REQUIRED = {
    "MAWB": ["MAWB", "Mawb", "Master AWB", "MasterAWB"],
    "ETA": ["ETA", "Eta", "Estimated Time of Arrival", "Arrival", "Arrival Date", "ETA Date"],
}

# ---------------- Main ----------------
if billing_file:
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

        if not (mawb_col and cost_col and sell_col):
            st.error("Billing sheet found but required columns could not be detected after scanning.")
            st.stop()

        # Normalize billing
        df = raw_df.copy()
        df["MAWB"] = df[mawb_col].astype(str).str.strip()
        df["Cost Amount"] = safe_numeric(df[cost_col])
        df["Sell Amount"] = safe_numeric(df[sell_col])

        if client_col:
            df["Client"] = df[client_col].astype(str).str.strip()
            df.loc[df["Client"].isin(["", "nan", "None"]), "Client"] = "UNKNOWN"
        else:
            df["Client"] = "UNKNOWN"

        # Drop blank MAWB
        df = df[df["MAWB"].ne("") & df["MAWB"].ne("nan")]

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
                    mdf["MAWB"] = mdf["MAWB"].astype(str).str.strip()

                    # robust parse for TEXT ETA
                    mdf["ETA"] = clean_eta_series(mdf["ETA"])

                    bad_eta_rows = int(mdf["ETA"].isna().sum())
                    total_rows = int(len(mdf))
                    if total_rows > 0 and bad_eta_rows > 0:
                        eta_parse_note = (
                            f"ETA parsing note: {bad_eta_rows} / {total_rows} ETA values could not be parsed and were left blank."
                        )

                    # Same MAWB multiple rows: take latest ETA (max)
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

        # ---- MAWB summary ----
        summary = (
            df.groupby("MAWB", as_index=False)
              .agg(
                  Client=("Client", "first"),      # keep one client label; if multi-client per MAWB, it will take first
                  Total_Cost=("Cost Amount", "sum"),
                  Total_Sell=("Sell Amount", "sum"),
                  Line_Count=("MAWB", "size"),
                  ETA=("ETA", "max")               # latest ETA per MAWB
              )
        )

        summary["ETA Month"] = summary["ETA"].dt.to_period("M").astype(str).replace("NaT", "")

        # Profit & Profit Margin
        summary["Profit"] = summary["Total_Sell"] - summary["Total_Cost"]
        summary["Profit Margin %"] = pct(summary["Profit"], summary["Total_Sell"])

        # Classification / Exceptions
        summary["Classification"] = summary.apply(
            lambda r: "Closed" if (r["Total_Cost"] > 0 and r["Total_Sell"] > 0) else "Open",
            axis=1
        )

        def exception_type(r):
            if r["Total_Cost"] == 0 and r["Total_Sell"] == 0:
                return "Cost=Sell=0"
            if r["Total_Sell"] == 0:
                return "Revenue=0"
            if r["Total_Cost"] == 0:
                return "Cost=0"
            return ""

        summary["Exception_Type"] = summary.apply(exception_type, axis=1)
        exceptions = summary[summary["Classification"].eq("Open")].copy()

        # ---- Client Summary (Profit / Profit Margin by Client) ----
        client_summary = (
            df.groupby("Client", as_index=False)
              .agg(
                  Total_Cost=("Cost Amount", "sum"),
                  Total_Sell=("Sell Amount", "sum"),
                  Profit=("Sell Amount", "sum"),   # temp, will overwrite next
                  Line_Count=("Client", "size"),
                  MAWB_Count=("MAWB", pd.Series.nunique),
                  Latest_ETA=("ETA", "max"),
              )
        )
        client_summary["Profit"] = client_summary["Total_Sell"] - client_summary["Total_Cost"]
        client_summary["Profit Margin %"] = pct(client_summary["Profit"], client_summary["Total_Sell"])
        client_summary = client_summary.sort_values("Profit", ascending=False)

        # ---- KPI ----
        total_mawb = len(summary)
        closed_cnt = int((summary["Classification"] == "Closed").sum())
        open_cnt = total_mawb - closed_cnt

        kpi = pd.DataFrame([{
            "Total MAWB": total_mawb,
            "Closed Count": closed_cnt,
            "Closed %": (closed_cnt / total_mawb) if total_mawb else 0,
            "Open Count": open_cnt,
            "Revenue=0 Count": int((summary["Exception_Type"] == "Revenue=0").sum()),
            "Cost=0 Count": int((summary["Exception_Type"] == "Cost=0").sum()),
            "Cost=Sell=0 Count": int((summary["Exception_Type"] == "Cost=Sell=0").sum()),
            "Total Cost": float(summary["Total_Cost"].sum()),
            "Total Sell": float(summary["Total_Sell"].sum()),
            "Total Profit": float(summary["Profit"].sum()),
            "Overall Profit Margin %": float((summary["Profit"].sum() / summary["Total_Sell"].sum()) if summary["Total_Sell"].sum() else 0),
            "ETA Filled %": float((summary["ETA"].notna().sum() / total_mawb)) if total_mawb else 0,
        }])

        # ---- UI ----
        if eta_parse_note:
            st.info(eta_parse_note)

        c1, c2 = st.columns([1, 1])
        with c1:
            st.subheader("KPI Summary")
            st.dataframe(kpi, use_container_width=True)
        with c2:
            st.subheader("Exceptions (Open items)")
            st.dataframe(exceptions, use_container_width=True)

        st.subheader("MAWB Summary (All)")
        st.dataframe(summary, use_container_width=True)

        st.subheader("Client Profit Summary")
        if (client_col is None) or (client_col == ""):
            st.warning("Client column not found in the billing export; showing Client as UNKNOWN.")
        st.dataframe(client_summary, use_container_width=True)

        # ---- Export report to Excel ----
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            kpi.to_excel(writer, index=False, sheet_name="KPI")
            summary.to_excel(writer, index=False, sheet_name="MAWB_Summary")
            client_summary.to_excel(writer, index=False, sheet_name="Client_Summary")
            exceptions.to_excel(writer, index=False, sheet_name="Exceptions")
            df.to_excel(writer, index=False, sheet_name="Raw_Billing_Enriched")

        st.download_button(
            "Download Report Excel",
            data=output.getvalue(),
            file_name="MAWB_Audit_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    except Exception as e:
        st.exception(e)
else:
    st.info("Please upload a Billing Charges Excel file to start.")
