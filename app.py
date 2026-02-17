import io
import re
import numpy as np
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Vendor Code Mapper + TLMF Allocator", layout="wide")
st.title("Vendor → Charge Code Mapper (Audit) + TLMF AR Allocation")
st.caption("Audit view: classify vendors to primary charge code by Cost; handle TLMF AR集中/AP分散 via AR allocation.")

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

def confidence_from_share(share: float) -> str:
    if share >= 0.8:
        return "High"
    if share >= 0.6:
        return "Medium"
    return "Low"

def format_pct(x):
    try:
        return f"{float(x)*100:.2f}%"
    except Exception:
        return ""

# Bucket mapping (可按你内部code体系继续扩展)
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

# ---------------- Upload ----------------
charges_file = st.file_uploader("Upload Mawb Charges (must include Vendor, Charge Code, Cost Amount, Sell Amount)", type=["xlsx"])
vendor_list_file = st.file_uploader("Optional: Upload Vendor List (Code / Full Name...)", type=["xlsx"])

st.divider()

# ---------------- Options ----------------
st.subheader("TLMF Logic Options")
enable_tlmf_alloc = st.checkbox("Enable TLMF AR allocation (recommended)", value=True)
alloc_method = st.radio(
    "TLMF AR allocation method",
    ["By Cost Share (recommended)", "By Line Count"],
    horizontal=True,
    disabled=not enable_tlmf_alloc
)

# ---------------- Required Columns ----------------
REQ = {
    "MAWB": ["MAWB", "Mawb", "Master AWB", "MasterAWB"],
    "Vendor": ["Vendor", "Supplier", "Carrier"],
    "Charge Code": ["Charge Code", "ChargeCode", "Code"],
    "Cost Amount": ["Cost Amount", "Cost", "AP Amount", "Total Cost"],
    "Sell Amount": ["Sell Amount", "Sell", "AR Amount", "Total Sell"],
}

if not charges_file:
    st.info("Please upload Mawb Charges file to start.")
    st.stop()

# ---------------- Read Charges ----------------
xls = pd.ExcelFile(charges_file)
sheet = xls.sheet_names[0]
raw = pd.read_excel(xls, sheet_name=sheet)

mawb_col = find_first_col(raw, REQ["MAWB"])
vendor_col = find_first_col(raw, REQ["Vendor"])
cc_col = find_first_col(raw, REQ["Charge Code"])
cost_col = find_first_col(raw, REQ["Cost Amount"])
sell_col = find_first_col(raw, REQ["Sell Amount"])

missing = [k for k, v in [("MAWB", mawb_col), ("Vendor", vendor_col), ("Charge Code", cc_col), ("Cost Amount", cost_col), ("Sell Amount", sell_col)] if not v]
if missing:
    st.error(f"Missing required columns: {missing}. Please check your export headers.")
    st.stop()

df = raw.copy()
df["MAWB"] = df[mawb_col].astype(str).str.strip().str.upper()
df["Vendor"] = df[vendor_col].fillna("").astype(str).str.strip().str.upper()
df["Charge Code"] = df[cc_col].fillna("").astype(str).str.strip().str.upper()
df["Cost Amount"] = safe_numeric(df[cost_col])
df["Sell Amount"] = safe_numeric(df[sell_col])

# Keep rows with charge code
df = df[df["Charge Code"].ne("")].copy()

# ---------------- Vendor → Primary Charge Code (Audit: by Cost) ----------------
df_v = df[(df["Vendor"].ne("")) & (df["Charge Code"].ne(""))].copy()

vc = (df_v.groupby(["Vendor", "Charge Code"], as_index=False)
        .agg(
            Line_Count=("MAWB", "size"),
            MAWB_Count=("MAWB", pd.Series.nunique),
            Total_Cost=("Cost Amount", "sum"),
            Total_Sell=("Sell Amount", "sum"),
        ))

vt = (vc.groupby("Vendor", as_index=False)
        .agg(
            Vendor_Total_Cost=("Total_Cost", "sum"),
            Vendor_Total_Sell=("Total_Sell", "sum"),
            Total_Lines=("Line_Count", "sum"),
            Total_MAWBs=("MAWB_Count", "sum"),
        ))

vc = vc.merge(vt[["Vendor","Vendor_Total_Cost"]], on="Vendor", how="left")
vc["Cost_Share"] = np.where(vc["Vendor_Total_Cost"]>0, vc["Total_Cost"]/vc["Vendor_Total_Cost"], 0.0)
vc["Category"] = vc["Charge Code"].apply(map_bucket)

vc = vc.sort_values(["Vendor","Total_Cost"], ascending=[True, False])
vc["Rank"] = vc.groupby("Vendor")["Total_Cost"].rank(method="first", ascending=False)

top1 = vc[vc["Rank"]==1].copy()
top2 = vc[vc["Rank"]==2].copy()

vendor_map = vt.merge(
    top1.rename(columns={
        "Charge Code":"Primary_Charge_Code",
        "Total_Cost":"Primary_Cost",
        "Total_Sell":"Primary_Sell",
        "Cost_Share":"Primary_Cost_Share",
        "Category":"Primary_Category",
        "Line_Count":"Primary_Lines",
        "MAWB_Count":"Primary_MAWBs",
    })[["Vendor","Primary_Charge_Code","Primary_Category","Primary_Cost","Primary_Sell","Primary_Cost_Share","Primary_Lines","Primary_MAWBs"]],
    on="Vendor", how="left"
).merge(
    top2.rename(columns={"Charge Code":"Secondary_Charge_Code","Total_Cost":"Secondary_Cost","Category":"Secondary_Category"})[
        ["Vendor","Secondary_Charge_Code","Secondary_Category","Secondary_Cost"]
    ],
    on="Vendor", how="left"
)

vendor_map["Confidence"] = vendor_map["Primary_Cost_Share"].apply(confidence_from_share)

# ---------------- TLMF Allocation ----------------
tlmf_alloc = pd.DataFrame()
if enable_tlmf_alloc:
    tlmf = df[df["Charge Code"].eq("TLMF")].copy()
    if not tlmf.empty:
        # MAWB level total sell and total cost
        mawb_sell = tlmf.groupby("MAWB", as_index=False)["Sell Amount"].sum().rename(columns={"Sell Amount":"TLMF_Total_Sell"})
        mawb_cost = tlmf.groupby("MAWB", as_index=False)["Cost Amount"].sum().rename(columns={"Cost Amount":"TLMF_Total_Cost"})

        # Vendor cost within MAWB
        vcost = (tlmf[tlmf["Vendor"].ne("")]
                    .groupby(["MAWB","Vendor"], as_index=False)
                    .agg(
                        Vendor_TLMF_Cost=("Cost Amount","sum"),
                        Vendor_Lines=("MAWB","size")
                    ))

        tlmf_alloc = vcost.merge(mawb_sell, on="MAWB", how="left").merge(mawb_cost, on="MAWB", how="left")
        tlmf_alloc["TLMF_Total_Sell"] = tlmf_alloc["TLMF_Total_Sell"].fillna(0.0)
        tlmf_alloc["TLMF_Total_Cost"] = tlmf_alloc["TLMF_Total_Cost"].fillna(0.0)

        if alloc_method.startswith("By Cost"):
            tlmf_alloc["Alloc_Share"] = np.where(
                tlmf_alloc["TLMF_Total_Cost"]>0,
                tlmf_alloc["Vendor_TLMF_Cost"] / tlmf_alloc["TLMF_Total_Cost"],
                0.0
            )
        else:
            # By line count within MAWB
            total_lines = tlmf_alloc.groupby("MAWB")["Vendor_Lines"].transform("sum")
            tlmf_alloc["Alloc_Share"] = np.where(total_lines>0, tlmf_alloc["Vendor_Lines"]/total_lines, 0.0)

        tlmf_alloc["Vendor_AR_Allocated"] = tlmf_alloc["TLMF_Total_Sell"] * tlmf_alloc["Alloc_Share"]
        tlmf_alloc["Vendor_Profit_Est"] = tlmf_alloc["Vendor_AR_Allocated"] - tlmf_alloc["Vendor_TLMF_Cost"]
        tlmf_alloc["Vendor_Margin_Est"] = np.where(
            tlmf_alloc["Vendor_AR_Allocated"]>0,
            tlmf_alloc["Vendor_Profit_Est"] / tlmf_alloc["Vendor_AR_Allocated"],
            0.0
        )
        tlmf_alloc["Vendor_Margin_Est_Display"] = tlmf_alloc["Vendor_Margin_Est"].apply(format_pct)

# ---------------- MAWB Summary (charge-code aware) ----------------
mawb_sum = (df.groupby("MAWB", as_index=False)
              .agg(
                  Total_Cost=("Cost Amount","sum"),
                  Total_Sell=("Sell Amount","sum"),
                  Line_Count=("MAWB","size")
              ))
mawb_sum["Profit"] = mawb_sum["Total_Sell"] - mawb_sum["Total_Cost"]
mawb_sum["Profit Margin %"] = np.where(mawb_sum["Total_Sell"]>0, mawb_sum["Profit"]/mawb_sum["Total_Sell"], 0.0)
mawb_sum["Profit Margin % Display"] = mawb_sum["Profit Margin %"].apply(format_pct)

# ---------------- Merge vendor list (optional) ----------------
merged_vendor_list = pd.DataFrame()
if vendor_list_file:
    vxls = pd.ExcelFile(vendor_list_file)
    vraw = pd.read_excel(vxls, sheet_name=vxls.sheet_names[0])
    # expected columns: Code, Full Name, UNLOCO, State (if present)
    code_col = find_first_col(vraw, ["Code", "Vendor Code", "Vendor"])
    if code_col:
        vraw2 = vraw.copy()
        vraw2["Code"] = vraw2[code_col].fillna("").astype(str).str.strip().str.upper()
        merged_vendor_list = vraw2.merge(vendor_map, left_on="Code", right_on="Vendor", how="left")
    else:
        st.warning("Vendor list uploaded but could not find a 'Code' column. Skipped merge.")

# ---------------- UI ----------------
st.subheader("Vendor → Primary Charge Code (Audit: by Cost Total)")
st.dataframe(vendor_map.sort_values(["Confidence","Vendor_Total_Cost"], ascending=[True, False]), use_container_width=True)

st.subheader("Top5 Charge Code distribution per Vendor (evidence)")
st.dataframe(vc.groupby("Vendor").head(5), use_container_width=True)

if enable_tlmf_alloc:
    st.subheader("TLMF Vendor-level View (Allocated AR, audit-friendly)")
    if tlmf_alloc.empty:
        st.info("No TLMF rows found (or no vendor on TLMF lines) in uploaded charges.")
    else:
        st.dataframe(
            tlmf_alloc.sort_values(["MAWB","Vendor_TLMF_Cost"], ascending=[True, False]),
            use_container_width=True
        )

st.subheader("MAWB Summary (All codes)")
st.dataframe(mawb_sum.sort_values("Profit", ascending=False), use_container_width=True)

if not merged_vendor_list.empty:
    st.subheader("Vendor List + Data-driven Mapping (merged)")
    st.dataframe(merged_vendor_list, use_container_width=True)

# ---------------- Export ----------------
output = io.BytesIO()
with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
    vendor_map.to_excel(writer, index=False, sheet_name="Vendor_Mapping")
    vc.groupby("Vendor").head(5).to_excel(writer, index=False, sheet_name="Top5_Code_Distribution")
    mawb_sum.to_excel(writer, index=False, sheet_name="MAWB_Summary")
    if enable_tlmf_alloc and not tlmf_alloc.empty:
        tlmf_alloc.to_excel(writer, index=False, sheet_name="TLMF_Allocated")
    if not merged_vendor_list.empty:
        merged_vendor_list.to_excel(writer, index=False, sheet_name="VendorList_Merged")

st.download_button(
    "Download Excel Output",
    data=output.getvalue(),
    file_name="Vendor_Code_Mapping_with_TLMF.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
