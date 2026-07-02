import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="StrategicERP Sub Project Mapper", layout="wide")

st.title("StrategicERP GRN / Purchase Bill Sub Project Mapper")

st.write("Upload GRN vs Purchase Bill Excel and PR Report Excel. The app will add Sub Project and create costing summary.")

grn_file = st.file_uploader("Upload GRN vs Purchase Bill Excel", type=["xlsx"])
pr_file = st.file_uploader("Upload PR Report Excel", type=["xlsx"])

def read_grn(file):
    df = pd.read_excel(file, header=1)
    df = df.dropna(how="all")
    df.columns = df.columns.astype(str).str.strip()
    return df

def read_pr(file):
    df = pd.read_excel(file)
    df = df.dropna(how="all")
    df.columns = df.columns.astype(str).str.strip()
    return df

def clean_pr_no(value):
    if pd.isna(value):
        return ""
    return str(value).strip().upper()

def create_excel(mapped_df, summary_df, unmatched_df):
    output = BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        mapped_df.to_excel(writer, sheet_name="Mapped Detail", index=False)
        summary_df.to_excel(writer, sheet_name="Sub Project Summary", index=False)
        unmatched_df.to_excel(writer, sheet_name="Unmatched PR", index=False)

    output.seek(0)
    return output

if grn_file and pr_file:
    grn_df = read_grn(grn_file)
    pr_df = read_pr(pr_file)

    st.subheader("Uploaded Data Preview")
    st.write("GRN / Purchase Bill Data")
    st.dataframe(grn_df.head())

    st.write("PR Report Data")
    st.dataframe(pr_df.head())

    grn_pr_col = "PRNo"
    pr_pr_col = "Purchase Requisition (PR) No."
    sub_project_col = "Sub Project"
    amount_col = "Bill Item Amt"

    if grn_pr_col not in grn_df.columns:
        st.error("PRNo column not found in GRN/Purchase Bill file.")
        st.stop()

    if pr_pr_col not in pr_df.columns:
        st.error("Purchase Requisition (PR) No. column not found in PR file.")
        st.stop()

    if sub_project_col not in pr_df.columns:
        st.error("Sub Project column not found in PR file.")
        st.stop()

    if amount_col not in grn_df.columns:
        st.error("Bill Item Amt column not found in GRN/Purchase Bill file.")
        st.stop()

    grn_df["PR_Clean"] = grn_df[grn_pr_col].apply(clean_pr_no)
    pr_df["PR_Clean"] = pr_df[pr_pr_col].apply(clean_pr_no)

    pr_mapping = pr_df[["PR_Clean", sub_project_col]].drop_duplicates()

    mapped_df = grn_df.merge(pr_mapping, on="PR_Clean", how="left")

    mapped_df["Mapping Status"] = mapped_df[sub_project_col].apply(
        lambda x: "Matched" if pd.notna(x) and str(x).strip() != "" else "Not Matched"
    )

    mapped_df[amount_col] = pd.to_numeric(mapped_df[amount_col], errors="coerce").fillna(0)

    summary_df = (
        mapped_df.groupby(sub_project_col, dropna=False)[amount_col]
        .sum()
        .reset_index()
        .sort_values(amount_col, ascending=False)
    )

    unmatched_df = mapped_df[mapped_df["Mapping Status"] == "Not Matched"]

    st.subheader("Result Summary")
    st.write("Sub Project Wise Costing")
    st.dataframe(summary_df)

    st.write("Unmatched PR Count:", len(unmatched_df))

    excel_output = create_excel(mapped_df, summary_df, unmatched_df)

    st.download_button(
        label="Download Final Excel",
        data=excel_output,
        file_name="sub_project_wise_purchase_bill_costing.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )