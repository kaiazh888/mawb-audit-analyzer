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
# Config
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
    # ✅ Destination candidates
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
    dest_col = find_first_col(raw_df, BILLING_OPTIONAL["Destination"])

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

    # ✅ Destination
    if dest_col:
        df_all["Destination"] = df_all[dest_col].astype(str).str.strip()
        df_all.loc[df_all["Destination"].isin(["", "nan", "None"]), "Destination"] = ""
    else:
        df_all["Destination"] = ""

    df_all = df_all[df_all["MAWB"].ne("")].copy()

    # line-level profit/margin (审计要求：每个sell/cost旁边都要有profit/margin)
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

    # ---- ETA mapping ----
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

    # ---- MAWB summary (✅ include Destination) ----
    # Destination: first non-empty per MAWB
    def first_non_empty(series: pd.Series) -> str:
        s = series.astype(str).replace(["nan", "None"], "").str.strip()
        s = s[s.ne("")]
        return s.iloc[0] if len(s) else ""

    summary = (
        df.groupby("MAWB", as_index=False)
          .agg(
              Client=("Client", "first"),
              Destination=("Destination", first_non_empty),  # ✅ added
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
    margin_anomalies = summary[
        ((summary["Profit Margin %"] < 0.30) | (summary["Profit Margin %"] > 0.80)) &
        (~summary["Profit Margin %"].isna())
    ].copy().sort_values("Profit Margin %")
    negative_profit = summary[summary["Profit"] < 0].copy().sort_values("Profit")

    both_zero = summary[(summary["Total_Sell"] == 0) & (summary["Total_Cost"] == 0)].copy().sort_values("MAWB")
    sell_zero_only = summary[(summary["Total_Sell"] == 0) & (summary["Total_Cost"] > 0)].copy().sort_values("Total_Cost", ascending=False)
    cost_zero_only = summary[(summary["Total_Cost"] == 0) & (summary["Total_Sell"] > 0)].copy().sort_values("Total_Sell", ascending=False)

    # ---- Client Summary (kept) ----
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

    # ---- Charge Code Profit <= 0 by MAWB (✅ include Destination) ----
    cc_mawb = (
        df.groupby(["MAWB", "Charge Code"], as_index=False)
          .agg(
              Client=("Client", "first"),
              Destination=("Destination", first_non_empty),  # ✅ added
              Vendor=("Vendor", "first"),
              Total_Cost=("Cost Amount", "sum"),
              Total_Sell=("Sell Amount", "sum"),
              ETA=("ETA", "max"),
          )
    )
    cc_mawb["Profit"] = cc_mawb["Total_Sell"] - cc_mawb["Total_Cost"]
    cc_mawb["Profit Margin %"] = ratio_or_nan(cc_mawb["Profit"], cc_mawb["Total_Sell"])
    cc_mawb["ETA Month"] = pd.to_datetime(cc_mawb["ETA"], errors="coerce").dt.to_period("M").astype(str).replace("NaT", "")
    chargecode_profit_le0_mawb = cc_mawb[cc_mawb["Profit"] <= 0].copy().sort_values(["Profit", "Total_Sell"], ascending=[True, False])

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

    def display_df(df_in, date_cols=None):
        out = df_in.copy()
        if date_cols:
            out = to_date_only(out, date_cols)
        if "Profit Margin %" in out.columns:
            out["Profit Margin %"] = out["Profit Margin %"].apply(format_pct_str_or_blank)
        return out

    st.subheader("MAWB Summary (All) — Auditable Only (✅ Destination included)")
    st.dataframe(display_df(summary, date_cols=["ETA"]), use_container_width=True)

    st.subheader("Exceptions (Open items) — Auditable Only (✅ Destination included)")
    st.dataframe(display_df(exceptions, date_cols=["ETA"]), use_container_width=True)

    st.subheader("Margin Anomalies — Auditable Only (✅ Destination included)")
    st.dataframe(display_df(margin_anomalies, date_cols=["ETA"]), use_container_width=True)

    st.subheader("Negative Profit — Auditable Only (✅ Destination included)")
    st.dataframe(display_df(negative_profit, date_cols=["ETA"]), use_container_width=True)

    st.subheader("Cost=Sell=0 — Auditable Only (✅ Destination included)")
    st.dataframe(display_df(both_zero, date_cols=["ETA"]), use_container_width=True)

    st.subheader("Sell=0 ONLY — Auditable Only (✅ Destination included)")
    st.dataframe(display_df(sell_zero_only, date_cols=["ETA"]), use_container_width=True)

    st.subheader("Cost=0 ONLY — Auditable Only (✅ Destination included)")
    st.dataframe(display_df(cost_zero_only, date_cols=["ETA"]), use_container_width=True)

    st.subheader("Client Summary — Auditable Only")
    st.dataframe(display_df(client_summary, date_cols=["Latest_ETA"]), use_container_width=True)

    st.subheader("Charge Code Summary — Auditable Only")
    st.dataframe(display_df(chargecode_summary), use_container_width=True)

    st.subheader("Vendor Summary — Auditable Only")
    st.dataframe(display_df(vendor_summary), use_container_width=True)

    st.subheader("Charge Code Profit <= 0 (by MAWB) — Auditable Only (✅ Destination included)")
    st.dataframe(display_df(chargecode_profit_le0_mawb, date_cols=["ETA"]), use_container_width=True)

    # ---------------- Export ----------------
    output = io.BytesIO()

    summary_x = to_date_only(summary, ["ETA"])
    exceptions_x = to_date_only(exceptions, ["ETA"])
    margin_anomalies_x = to_date_only(margin_anomalies, ["ETA"])
    negative_profit_x = to_date_only(negative_profit, ["ETA"])
    both_zero_x = to_date_only(both_zero, ["ETA"])
    sell_zero_only_x = to_date_only(sell_zero_only, ["ETA"])
    cost_zero_only_x = to_date_only(cost_zero_only, ["ETA"])
    client_summary_x = to_date_only(client_summary, ["Latest_ETA"])
    chargecode_profit_le0_mawb_x = to_date_only(chargecode_profit_le0_mawb, ["ETA"])
    df_x = to_date_only(df, ["ETA"])

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        workbook = writer.book

        def set_percent(ws, dfx, col_name="Profit Margin %"):
            if col_name in dfx.columns:
                col = list(dfx.columns).index(col_name)
                ws.set_column(col, col, 16, workbook.add_format({"num_format": "0.00%"}))

        # Analysis Summary (basic for now)
        ws = workbook.add_worksheet("Analysis Summary")
        writer.sheets["Analysis Summary"] = ws
        ws.write(0, 0, "Analysis Summary (Auditable Only)", workbook.add_format({"bold": True, "font_size": 14}))
        ws.write(2, 0, "Links:", workbook.add_format({"bold": True}))

        tab_links = [
            ("MAWB Summary — detail", "MAWB_Summary"),
            ("Exceptions — detail", "Exceptions"),
            ("Margin Anomalies — detail", "Margin_Anomalies"),
            ("Negative Profit — detail", "Negative_Profit"),
            ("Cost=Sell=0 — detail", "Both_Zero"),
            ("Sell=0 only — detail", "Sell_Zero_Only"),
            ("Cost=0 only — detail", "Cost_Zero_Only"),
            ("Client Summary — detail", "Client_Summary"),
            ("Charge code summary — detail", "ChargeCode_Summary"),
            ("Vendor summary — detail", "Vendor_Summary"),
            ("ChargeCode Profit<=0 by MAWB — detail", "ChargeCode_ProfitLE0_MAWB"),
            ("Raw enriched billing (auditable only) — detail", "Raw_Billing_Enriched"),
        ]
        if mawb_keep:
            tab_links.insert(0, ("MAWB not found from filter — detail", "MAWB_Not_Found"))

        r = 3
        for text, sheet_name in tab_links:
            ws.write_url(r, 0, f"internal:'{sheet_name}'!A1", string=text)
            r += 1

        # Structural totals block
        if structural_totals is not None:
            start = r + 2
            ws.write(start, 0, "Structural Client Totals (Excluded)", workbook.add_format({"bold": True}))
            structural_totals.to_excel(writer, index=False, sheet_name="Analysis Summary", startrow=start + 1, startcol=0)
            set_percent(ws, structural_totals)

        # Write other sheets
        exceptions_x.to_excel(writer, index=False, sheet_name="Exceptions")
        summary_x.to_excel(writer, index=False, sheet_name="MAWB_Summary")
        margin_anomalies_x.to_excel(writer, index=False, sheet_name="Margin_Anomalies")
        negative_profit_x.to_excel(writer, index=False, sheet_name="Negative_Profit")
        both_zero_x.to_excel(writer, index=False, sheet_name="Both_Zero")
        sell_zero_only_x.to_excel(writer, index=False, sheet_name="Sell_Zero_Only")
        cost_zero_only_x.to_excel(writer, index=False, sheet_name="Cost_Zero_Only")
        client_summary_x.to_excel(writer, index=False, sheet_name="Client_Summary")
        chargecode_summary.to_excel(writer, index=False, sheet_name="ChargeCode_Summary")
        vendor_summary.to_excel(writer, index=False, sheet_name="Vendor_Summary")
        chargecode_profit_le0_mawb_x.to_excel(writer, index=False, sheet_name="ChargeCode_ProfitLE0_MAWB")
        df_x.to_excel(writer, index=False, sheet_name="Raw_Billing_Enriched")

        if mawb_keep:
            pd.DataFrame({"MAWB": mawb_not_found}).to_excel(writer, index=False, sheet_name="MAWB_Not_Found")

        # percent formatting
        for sh, dfx in [
            ("Exceptions", exceptions_x),
            ("MAWB_Summary", summary_x),
            ("Margin_Anomalies", margin_anomalies_x),
            ("Negative_Profit", negative_profit_x),
            ("Both_Zero", both_zero_x),
            ("Sell_Zero_Only", sell_zero_only_x),
            ("Cost_Zero_Only", cost_zero_only_x),
            ("Client_Summary", client_summary_x),
            ("ChargeCode_Summary", chargecode_summary),
            ("Vendor_Summary", vendor_summary),
            ("ChargeCode_ProfitLE0_MAWB", chargecode_profit_le0_mawb_x),
            ("Raw_Billing_Enriched", df_x),
        ]:
            ws2 = writer.sheets.get(sh)
            if ws2 is not None:
                set_percent(ws2, dfx)

    st.download_button(
        "Download Report Excel (MAWB tabs include Destination)",
        data=output.getvalue(),
        file_name="MAWB_Audit_Report_WITH_DESTINATION.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

except Exception as e:
    st.exception(e)
