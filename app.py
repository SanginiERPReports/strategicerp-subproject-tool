import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime

st.set_page_config(page_title="GRN Sub Project Mapper", layout="wide")

st.title("GRN vs Purchase Bill - Sub Project Mapper")
st.write("Double checking: PR No + Item Desc. If item match is not found, PR-only mapping is used only when safe.")


grn_file = st.file_uploader("Upload GRN vs Purchase Bill Excel", type=["xlsx"])
pr_file = st.file_uploader("Upload PR Report Excel", type=["xlsx"])


def clean_text(value):
    if pd.isna(value):
        return ""
    return " ".join(str(value).strip().upper().split())


def read_grn_file(file):
    return pd.read_excel(file, header=1)


def read_pr_file(file):
    return pd.read_excel(file, header=0)


def join_unique(values):
    cleaned = []
    for v in values:
        if pd.notna(v) and str(v).strip() != "":
            cleaned.append(str(v).strip())

    unique_values = sorted(set(cleaned))
    return " | ".join(unique_values)


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

        for col_num, value in enumerate(df.columns):
            worksheet.write(0, col_num, value, header_format)
            worksheet.set_column(col_num, col_num, 18)

    output.seek(0)
    return output


if grn_file and pr_file:
    try:
        grn_df = read_grn_file(grn_file)
        pr_df = read_pr_file(pr_file)

        grn_df = grn_df.dropna(how="all")
        pr_df = pr_df.dropna(how="all")

        grn_df.columns = grn_df.columns.astype(str).str.strip()
        pr_df.columns = pr_df.columns.astype(str).str.strip()

        grn_pr_col = "PRNo"
        grn_item_col = "Item Desc"
        project_col = "Project Name"

        pr_pr_col = "Purchase Requisition (PR) No."
        pr_item_col = "Item Desc"
        sub_project_col = "Sub Project"

        required_grn_columns = [grn_pr_col, grn_item_col, project_col]
        required_pr_columns = [pr_pr_col, pr_item_col, sub_project_col]

        missing_grn_columns = [col for col in required_grn_columns if col not in grn_df.columns]
        missing_pr_columns = [col for col in required_pr_columns if col not in pr_df.columns]

        if missing_grn_columns:
            st.error(f"Missing column in GRN file: {', '.join(missing_grn_columns)}")
            st.stop()

        if missing_pr_columns:
            st.error(f"Missing column in PR file: {', '.join(missing_pr_columns)}")
            st.stop()

        grn_df["PR_Clean"] = grn_df[grn_pr_col].apply(clean_text)
        grn_df["Item_Clean"] = grn_df[grn_item_col].apply(clean_text)

        pr_df["PR_Clean"] = pr_df[pr_pr_col].apply(clean_text)
        pr_df["Item_Clean"] = pr_df[pr_item_col].apply(clean_text)

        pr_item_mapping = (
            pr_df
            .groupby(["PR_Clean", "Item_Clean"], as_index=False)[sub_project_col]
            .agg(join_unique)
            .rename(columns={sub_project_col: "Sub Project From PR Item"})
        )

        pr_only_check = (
            pr_df
            .groupby("PR_Clean")[sub_project_col]
            .agg(lambda x: sorted(set([str(v).strip() for v in x if pd.notna(v) and str(v).strip() != ""])))
            .reset_index()
        )

        pr_only_check["Unique Sub Project Count"] = pr_only_check[sub_project_col].apply(len)
        pr_only_check["Sub Project From PR Only"] = pr_only_check[sub_project_col].apply(
            lambda x: x[0] if len(x) == 1 else ""
        )

        pr_only_mapping = pr_only_check[["PR_Clean", "Unique Sub Project Count", "Sub Project From PR Only"]]

        mapped_df = grn_df.merge(
            pr_item_mapping,
            on=["PR_Clean", "Item_Clean"],
            how="left"
        )

        mapped_df = mapped_df.merge(
            pr_only_mapping,
            on="PR_Clean",
            how="left"
        )

        mapped_df[sub_project_col] = mapped_df["Sub Project From PR Item"]

        safe_pr_only_condition = (
            mapped_df[sub_project_col].isna() |
            (mapped_df[sub_project_col].astype(str).str.strip() == "")
        ) & (
            mapped_df["Unique Sub Project Count"] == 1
        )

        mapped_df.loc[safe_pr_only_condition, sub_project_col] = mapped_df.loc[
            safe_pr_only_condition, "Sub Project From PR Only"
        ]

        mapped_df["Mapping Status"] = "Not Matched"

        mapped_df.loc[
            mapped_df["Sub Project From PR Item"].notna() &
            (mapped_df["Sub Project From PR Item"].astype(str).str.strip() != ""),
            "Mapping Status"
        ] = "Matched by PR + Item"

        mapped_df.loc[
            safe_pr_only_condition,
            "Mapping Status"
        ] = "Matched by PR Only - Safe"

        mapped_df.loc[
            (
                mapped_df[sub_project_col].isna() |
                (mapped_df[sub_project_col].astype(str).str.strip() == "")
            ) &
            (mapped_df["Unique Sub Project Count"] > 1),
            "Mapping Status"
        ] = "Not Matched - Multiple Sub Projects in Same PR"

        columns_to_remove = [
            "Excise Duty Amt",
            "Loading / Unloading Chgs",
            "Others Chgs",
            "CESS Amt",
            "PR_Clean",
            "Item_Clean",
            "Sub Project From PR Item",
            "Sub Project From PR Only",
            "Unique Sub Project Count"
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

        if "Mapping Status" in cols:
            cols.remove("Mapping Status")
            cols.append("Mapping Status")

        mapped_df = mapped_df[cols]

        total_records = len(mapped_df)
        matched_pr_item = (mapped_df["Mapping Status"] == "Matched by PR + Item").sum()
        matched_pr_only = (mapped_df["Mapping Status"] == "Matched by PR Only - Safe").sum()
        not_matched = total_records - matched_pr_item - matched_pr_only

        st.success("File processed successfully.")

        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Total Records", f"{total_records:,}")
        col2.metric("Matched by PR + Item", f"{matched_pr_item:,}")
        col3.metric("Matched by PR Only", f"{matched_pr_only:,}")
        col4.metric("Not Matched", f"{not_matched:,}")

        st.subheader("Preview")
        st.dataframe(mapped_df.head(100), use_container_width=True)

        output_excel = create_output_excel(mapped_df)

        filename = f"GRN_SubProject_Double_Check_{datetime.now():%Y%m%d_%H%M}.xlsx"

        st.download_button(
            label="Download Updated Excel",
            data=output_excel,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error("Something went wrong while processing the files.")
        st.exception(e)
