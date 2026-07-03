import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Sub Project Mapper", layout="wide")

st.title("GRN vs Purchase Bill - Sub Project Mapper")

st.write(
    "Upload GRN vs Purchase Bill Excel and PR Report Excel. "
    "This tool will add only one new column: Sub Project, beside Project Name."
)

grn_file = st.file_uploader("Upload GRN vs Purchase Bill Excel", type=["xlsx"])
pr_file = st.file_uploader("Upload PR Report Excel", type=["xlsx"])


def clean_pr(value):
    if pd.isna(value):
        return ""
    return str(value).strip().upper()


def create_output_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="GRN Purchase Bill")
    output.seek(0)
    return output


if grn_file and pr_file:
    grn_df = pd.read_excel(grn_file, header=1)
    pr_df = pd.read_excel(pr_file)

    grn_df.columns = grn_df.columns.astype(str).str.strip()
    pr_df.columns = pr_df.columns.astype(str).str.strip()

    grn_pr_col = "PRNo"
    project_col = "Project Name"

    pr_pr_col = "Purchase Requisition (PR) No."
    sub_project_col = "Sub Project"

    if grn_pr_col not in grn_df.columns:
        st.error("PRNo column not found in GRN file.")
        st.stop()

    if project_col not in grn_df.columns:
        st.error("Project Name column not found in GRN file.")
        st.stop()

    if pr_pr_col not in pr_df.columns:
        st.error("Purchase Requisition (PR) No. column not found in PR file.")
        st.stop()

    if sub_project_col not in pr_df.columns:
        st.error("Sub Project column not found in PR file.")
        st.stop()

    grn_df["PR_Clean"] = grn_df[grn_pr_col].apply(clean_pr)
    pr_df["PR_Clean"] = pr_df[pr_pr_col].apply(clean_pr)

    pr_mapping = (
        pr_df[["PR_Clean", sub_project_col]]
        .dropna(subset=["PR_Clean"])
        .drop_duplicates(subset=["PR_Clean"], keep="first")
    )

    mapped_df = grn_df.merge(pr_mapping, on="PR_Clean", how="left")

    cols = list(mapped_df.columns)

    cols.remove(sub_project_col)
    project_index = cols.index(project_col)
    cols.insert(project_index + 1, sub_project_col)

    mapped_df = mapped_df[cols]

    mapped_df = mapped_df.drop(columns=["PR_Clean"])

    st.success("Sub Project column added successfully.")

    st.subheader("Preview")
    st.dataframe(mapped_df.head(50))

    output_excel = create_output_excel(mapped_df)

    st.download_button(
        label="Download Updated Excel",
        data=output_excel,
        file_name="grn_purchase_bill_with_sub_project.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
