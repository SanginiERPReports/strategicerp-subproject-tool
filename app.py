import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime

st.set_page_config(page_title="Sub Project Mapper", layout="wide")

st.title("GRN vs Purchase Bill - Sub Project Mapper")

st.write(
    "Upload GRN vs Purchase Bill Excel and PR Report Excel. "
    "This tool will add Sub Project beside Project Name and remove unnecessary columns."
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

        workbook = writer.book
        worksheet = writer.sheets["GRN Purchase Bill"]

        header_format = workbook.add_format({
            "bold": True,
            "bg_color": "#1C2551",
            "font_color": "#FFFFFF",
            "border": 1
        })

        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)
            worksheet.set_column(col_num, col_num, 18)

    output.seek(0)
    return output


if grn_file and pr_file:
    try:
        grn_df = pd.read_excel(grn_file, header=1)
        pr_df = pd.read_excel(pr_file)

        grn_df = grn_df.dropna(how="all")
        pr_df = pr_df.dropna(how="all")

        grn_df.columns = grn_df.columns.astype(str).str.strip()
        pr_df.columns = pr_df.columns.astype(str).str.strip()

        grn_pr_col = "PRNo"
        project_col = "Project Name"

        pr_pr_col = "Purchase Requisition (PR) No."
        sub_project_col = "Sub Project"

        required_grn_columns = [grn_pr_col, project_col]
        required_pr_columns = [pr_pr_col, sub_project_col]

        missing_grn_columns = [
            col for col in required_grn_columns if col not in grn_df.columns
        ]

        missing_pr_columns = [
            col for col in required_pr_columns if col not in pr_df.columns
        ]

        if missing_grn_columns:
            st.error(f"Missing column in GRN file: {', '.join(missing_grn_columns)}")
            st.stop()

        if missing_pr_columns:
            st.error(f"Missing column in PR file: {', '.join(missing_pr_columns)}")
            st.stop()

        grn_df["PR_Clean"] = grn_df[grn_pr_col].apply(clean_pr)
        pr_df["PR_Clean"] = pr_df[pr_pr_col].apply(clean_pr)

        pr_mapping = (
            pr_df[["PR_Clean", sub_project_col]]
            .dropna(subset=["PR_Clean"])
            .drop_duplicates(subset=["PR_Clean"], keep="first")
        )

        mapped_df = grn_df.merge(pr_mapping, on="PR_Clean", how="left")

        columns_to_remove = [
            "Excise Duty Amt",
            "Loading / Unloading Chgs",
            "Others Chgs",
            "CESS Amt"
        ]

        mapped_df = mapped_df.drop(
            columns=[col for col in columns_to_remove if col in mapped_df.columns],
            errors="ignore"
        )

        cols = list(mapped_df.columns)

        if sub_project_col in cols:
            cols.remove(sub_project_col)

        project_index = cols.index(project_col)
        cols.insert(project_index + 1, sub_project_col)

        mapped_df = mapped_df[cols]

        mapped_df = mapped_df.drop(columns=["PR_Clean"], errors="ignore")

        total_records = len(mapped_df)
        matched_records = mapped_df[sub_project_col].notna().sum()
        unmatched_records = total_records - matched_records

        st.success("Sub Project column added successfully.")

        col1, col2, col3 = st.columns(3)

        col1.metric("Total Records", f"{total_records:,}")
        col2.metric("Matched", f"{matched_records:,}")
        col3.metric("Not Matched", f"{unmatched_records:,}")

        st.subheader("Preview")
        st.dataframe(mapped_df.head(100), use_container_width=True)

        output_excel = create_output_excel(mapped_df)

        filename = f"GRN_SubProject_{datetime.now():%Y%m%d_%H%M}.xlsx"

        st.download_button(
            label="Download Updated Excel",
            data=output_excel,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    except Exception as e:
        st.error("Something went wrong while processing the files.")
        st.exception(e)
