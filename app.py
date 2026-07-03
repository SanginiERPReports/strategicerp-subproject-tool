import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
import re

st.set_page_config(page_title="StrategicERP Cost Allocation", layout="wide")

st.title("StrategicERP PR Based Cost Allocation")
st.write("Purchase Bill = cost source | PR = sub-project source | Stock Ledger = consumed vs stock-in-hand status")

grn_file = st.file_uploader("Upload GRN vs Purchase Bill Excel", type=["xlsx"])
pr_file = st.file_uploader("Upload PR Excel", type=["xlsx"])
stock_file = st.file_uploader("Upload Stock Ledger Excel", type=["xlsx"])


def clean_text(value):
    if pd.isna(value):
        return ""
    value = str(value).upper().strip()
    value = re.sub(r"\s+", " ", value)
    return value


def clean_item(value):
    if pd.isna(value):
        return ""
    value = str(value).upper().strip()
    value = re.sub(r"[^A-Z0-9]+", " ", value)
    value = re.sub(r"\s+", " ", value)
    return value


def split_sub_projects(value):
    if pd.isna(value):
        return []
    text = str(value).strip()
    text = text.replace("\n", ",")
    text = text.replace("|", ",")
    text = text.replace(";", ",")
    return [x.strip() for x in text.split(",") if x.strip()]


def join_unique(values):
    output = []
    for v in values:
        if pd.notna(v) and str(v).strip() != "":
            output.append(str(v).strip())
    return " | ".join(sorted(set(output)))


def create_excel(sheets):
    output = BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        for sheet_name, df in sheets.items():
            sheet = sheet_name[:31]
            df.to_excel(writer, sheet_name=sheet, index=False)

            workbook = writer.book
            worksheet = writer.sheets[sheet]

            header_format = workbook.add_format({
                "bold": True,
                "bg_color": "#1C2551",
                "font_color": "#FFFFFF",
                "border": 1
            })

            money_format = workbook.add_format({"num_format": "#,##0.00"})
            percent_format = workbook.add_format({"num_format": "0.00%"})

            for col_num, col_name in enumerate(df.columns):
                worksheet.write(0, col_num, col_name, header_format)

                width = min(max(len(str(col_name)) + 4, 16), 35)

                if "allocation %" in str(col_name).lower() or "consumed %" in str(col_name).lower():
                    worksheet.set_column(col_num, col_num, 16, percent_format)
                elif any(x in str(col_name).lower() for x in ["amount", "amt", "gst", "cost", "principal", "total"]):
                    worksheet.set_column(col_num, col_num, 18, money_format)
                else:
                    worksheet.set_column(col_num, col_num, width)

            worksheet.freeze_panes(1, 0)

    output.seek(0)
    return output


if grn_file and pr_file and stock_file:
    try:
        grn_df = pd.read_excel(grn_file, header=1)
        pr_df = pd.read_excel(pr_file, header=0)
        stock_df = pd.read_excel(stock_file, header=0)

        grn_df = grn_df.dropna(how="all")
        pr_df = pr_df.dropna(how="all")
        stock_df = stock_df.dropna(how="all")

        grn_df.columns = grn_df.columns.astype(str).str.strip()
        pr_df.columns = pr_df.columns.astype(str).str.strip()
        stock_df.columns = stock_df.columns.astype(str).str.strip()

        grn_pr_col = "PRNo"
        grn_po_col = "PO No"
        grn_grn_col = "GR No"
        grn_item_col = "Item Desc"
        project_col = "Project Name"
        gst_col = "GST Amt"
        bill_amt_col = "Bill Item Amt"

        pr_pr_col = "Purchase Requisition (PR) No."
        pr_item_col = "Item Desc"
        pr_qty_col = "Quantity"
        sub_project_col = "Sub Project"

        stock_pr_col = "P.RNo"
        stock_po_col = "P.O. No"
        stock_grn_col = "G.R. No"
        stock_item_col = "Item Desc"
        stock_sub_project_col = "Sub Project"
        stock_received_qty_col = "Received Qty"
        stock_issued_qty_col = "Issued Qty"

        required_grn_cols = [
            grn_pr_col, grn_po_col, grn_grn_col, grn_item_col,
            project_col, gst_col, bill_amt_col
        ]

        required_pr_cols = [
            pr_pr_col, pr_item_col, pr_qty_col, sub_project_col
        ]

        required_stock_cols = [
            stock_pr_col, stock_po_col, stock_grn_col,
            stock_item_col, stock_sub_project_col, stock_issued_qty_col
        ]

        missing_grn = [c for c in required_grn_cols if c not in grn_df.columns]
        missing_pr = [c for c in required_pr_cols if c not in pr_df.columns]
        missing_stock = [c for c in required_stock_cols if c not in stock_df.columns]

        if missing_grn:
            st.error(f"Missing columns in GRN Purchase Bill file: {', '.join(missing_grn)}")
            st.stop()

        if missing_pr:
            st.error(f"Missing columns in PR file: {', '.join(missing_pr)}")
            st.stop()

        if missing_stock:
            st.error(f"Missing columns in Stock Ledger file: {', '.join(missing_stock)}")
            st.stop()

        grn_df["Original Row No"] = range(1, len(grn_df) + 1)

        grn_df["PR_Clean"] = grn_df[grn_pr_col].apply(clean_text)
        grn_df["PO_Clean"] = grn_df[grn_po_col].apply(clean_text)
        grn_df["GRN_Clean"] = grn_df[grn_grn_col].apply(clean_text)
        grn_df["Item_Clean"] = grn_df[grn_item_col].apply(clean_item)

        pr_df["PR_Clean"] = pr_df[pr_pr_col].apply(clean_text)
        pr_df["Item_Clean"] = pr_df[pr_item_col].apply(clean_item)
        pr_df["PR_Qty"] = pd.to_numeric(pr_df[pr_qty_col], errors="coerce").fillna(0)

        stock_df["PR_Clean"] = stock_df[stock_pr_col].apply(clean_text)
        stock_df["PO_Clean"] = stock_df[stock_po_col].apply(clean_text)
        stock_df["GRN_Clean"] = stock_df[stock_grn_col].apply(clean_text)
        stock_df["Item_Clean"] = stock_df[stock_item_col].apply(clean_item)
        stock_df["Issued_Qty_Clean"] = pd.to_numeric(stock_df[stock_issued_qty_col], errors="coerce").fillna(0)

        if stock_received_qty_col in stock_df.columns:
            stock_df["Received_Qty_Clean"] = pd.to_numeric(stock_df[stock_received_qty_col], errors="coerce").fillna(0)
        else:
            stock_df["Received_Qty_Clean"] = 0

        pr_rows = []

        for _, row in pr_df.iterrows():
            pr_no = row["PR_Clean"]
            item = row["Item_Clean"]
            qty = row["PR_Qty"]
            sub_projects = split_sub_projects(row[sub_project_col])

            if pr_no == "" or item == "" or not sub_projects:
                continue

            if qty <= 0:
                qty = 1

            each_qty = qty / len(sub_projects)

            for sp in sub_projects:
                pr_rows.append({
                    "PR_Clean": pr_no,
                    "Item_Clean": item,
                    "Final Sub Project": sp,
                    "PR Allocation Qty": each_qty
                })

        pr_alloc_df = pd.DataFrame(pr_rows)

        if pr_alloc_df.empty:
            st.error("No PR sub-project allocation data found.")
            st.stop()

        pr_alloc_df["Total Allocation Qty"] = pr_alloc_df.groupby(
            ["PR_Clean", "Item_Clean"]
        )["PR Allocation Qty"].transform("sum")

        pr_alloc_df["Allocation %"] = pr_alloc_df["PR Allocation Qty"] / pr_alloc_df["Total Allocation Qty"]

        pr_item_mapping = pr_alloc_df[
            ["PR_Clean", "Item_Clean", "Final Sub Project", "Allocation %"]
        ].copy()

        pr_only_alloc = (
            pr_alloc_df
            .groupby(["PR_Clean", "Final Sub Project"], as_index=False)["PR Allocation Qty"]
            .sum()
        )

        pr_only_alloc["Total PR Qty"] = pr_only_alloc.groupby("PR_Clean")["PR Allocation Qty"].transform("sum")
        pr_only_alloc["Allocation %"] = pr_only_alloc["PR Allocation Qty"] / pr_only_alloc["Total PR Qty"]

        item_mapped = grn_df.merge(
            pr_item_mapping,
            on=["PR_Clean", "Item_Clean"],
            how="left"
        )

        item_mapped["Mapping Source"] = "PR + Item"

        matched = item_mapped[item_mapped["Final Sub Project"].notna()].copy()
        unmatched_original_rows = item_mapped[item_mapped["Final Sub Project"].isna()]["Original Row No"].unique()

        fallback_base = grn_df[grn_df["Original Row No"].isin(unmatched_original_rows)].copy()

        fallback = fallback_base.merge(
            pr_only_alloc[["PR_Clean", "Final Sub Project", "Allocation %"]],
            on="PR_Clean",
            how="left"
        )

        fallback["Mapping Source"] = "PR Only Fallback"

        detail_df = pd.concat([matched, fallback], ignore_index=True)

        detail_df["Allocation %"] = pd.to_numeric(detail_df["Allocation %"], errors="coerce").fillna(0)

        detail_df["Bill Item Amt Numeric"] = pd.to_numeric(detail_df[bill_amt_col], errors="coerce").fillna(0)
        detail_df["GST Amt Numeric"] = pd.to_numeric(detail_df[gst_col], errors="coerce").fillna(0)

        detail_df["Allocated Procurement Amount"] = detail_df["Bill Item Amt Numeric"] * detail_df["Allocation %"]
        detail_df["Allocated GST Amount"] = detail_df["GST Amt Numeric"] * detail_df["Allocation %"]
        detail_df["Allocated Principal Amount"] = detail_df["Allocated Procurement Amount"] - detail_df["Allocated GST Amount"]

        stock_summary = (
            stock_df
            .groupby(["PR_Clean", "PO_Clean", "GRN_Clean", "Item_Clean"], as_index=False)
            .agg({
                "Received_Qty_Clean": "sum",
                "Issued_Qty_Clean": "sum",
                stock_sub_project_col: join_unique
            })
            .rename(columns={
                "Received_Qty_Clean": "Stock Ledger Received Qty",
                "Issued_Qty_Clean": "Stock Ledger Issued Qty",
                stock_sub_project_col: "Stock Ledger Issue Sub Projects"
            })
        )

        detail_df = detail_df.merge(
            stock_summary,
            on=["PR_Clean", "PO_Clean", "GRN_Clean", "Item_Clean"],
            how="left"
        )

        detail_df["Stock Ledger Received Qty"] = pd.to_numeric(
            detail_df["Stock Ledger Received Qty"], errors="coerce"
        ).fillna(0)

        detail_df["Stock Ledger Issued Qty"] = pd.to_numeric(
            detail_df["Stock Ledger Issued Qty"], errors="coerce"
        ).fillna(0)

        detail_df["Consumed %"] = 0.0

        has_received_qty = detail_df["Stock Ledger Received Qty"] > 0
        detail_df.loc[has_received_qty, "Consumed %"] = (
            detail_df.loc[has_received_qty, "Stock Ledger Issued Qty"] /
            detail_df.loc[has_received_qty, "Stock Ledger Received Qty"]
        )

        no_received_but_issued = (
            (detail_df["Stock Ledger Received Qty"] <= 0) &
            (detail_df["Stock Ledger Issued Qty"] > 0)
        )

        detail_df.loc[no_received_but_issued, "Consumed %"] = 1.0

        detail_df["Consumed %"] = detail_df["Consumed %"].clip(lower=0, upper=1)

        detail_df["Consumed Principal Amount"] = detail_df["Allocated Principal Amount"] * detail_df["Consumed %"]
        detail_df["Consumed GST Amount"] = detail_df["Allocated GST Amount"] * detail_df["Consumed %"]
        detail_df["Consumed Total Amount"] = detail_df["Allocated Procurement Amount"] * detail_df["Consumed %"]

        detail_df["Stock In Hand Principal Amount"] = detail_df["Allocated Principal Amount"] - detail_df["Consumed Principal Amount"]
        detail_df["Stock In Hand GST Amount"] = detail_df["Allocated GST Amount"] - detail_df["Consumed GST Amount"]
        detail_df["Stock In Hand Total Amount"] = detail_df["Allocated Procurement Amount"] - detail_df["Consumed Total Amount"]

        detail_df["Cost Status"] = "Stock In Hand / Received Not Issued"

        detail_df.loc[
            detail_df["Consumed %"] >= 1,
            "Cost Status"
        ] = "Fully Consumed / Issued"

        detail_df.loc[
            (detail_df["Consumed %"] > 0) & (detail_df["Consumed %"] < 1),
            "Cost Status"
        ] = "Partially Consumed / Partly In Stock"

        detail_df["Mapping Status"] = "Matched"

        detail_df.loc[
            detail_df["Final Sub Project"].isna() | (detail_df["Final Sub Project"].astype(str).str.strip() == ""),
            "Mapping Status"
        ] = "Manual Review Required"

        detail_df["Stock Validation"] = "No Issue Entry Yet"

        detail_df.loc[
            detail_df["Stock Ledger Issued Qty"] > 0,
            "Stock Validation"
        ] = "Issue Entry Available"

        columns_to_remove = [
            "Excise Duty Amt",
            "Loading / Unloading Chgs",
            "Others Chgs",
            "CESS Amt",
            "PR_Clean",
            "PO_Clean",
            "GRN_Clean",
            "Item_Clean",
            "Bill Item Amt Numeric",
            "GST Amt Numeric"
        ]

        detail_df = detail_df.drop(
            columns=[c for c in columns_to_remove if c in detail_df.columns],
            errors="ignore"
        )

        cols = list(detail_df.columns)

        for c in ["Final Sub Project", "Allocation %", "Cost Status"]:
            if c in cols:
                cols.remove(c)

        project_index = cols.index(project_col)
        cols.insert(project_index + 1, "Final Sub Project")
        cols.insert(project_index + 2, "Allocation %")
        cols.insert(project_index + 3, "Cost Status")

        end_cols = [
            "Allocated Principal Amount",
            "Allocated GST Amount",
            "Allocated Procurement Amount",
            "Consumed %",
            "Consumed Principal Amount",
            "Consumed GST Amount",
            "Consumed Total Amount",
            "Stock In Hand Principal Amount",
            "Stock In Hand GST Amount",
            "Stock In Hand Total Amount",
            "Stock Ledger Received Qty",
            "Stock Ledger Issued Qty",
            "Mapping Source",
            "Mapping Status",
            "Stock Ledger Issue Sub Projects",
            "Stock Validation"
        ]

        for c in end_cols:
            if c in cols:
                cols.remove(c)
                cols.append(c)

        detail_df = detail_df[cols]

        sub_project_summary = (
            detail_df
            .groupby("Final Sub Project", dropna=False)[[
                "Allocated Principal Amount",
                "Allocated GST Amount",
                "Allocated Procurement Amount",
                "Consumed Principal Amount",
                "Consumed GST Amount",
                "Consumed Total Amount",
                "Stock In Hand Principal Amount",
                "Stock In Hand GST Amount",
                "Stock In Hand Total Amount"
            ]]
            .sum()
            .reset_index()
            .sort_values("Allocated Procurement Amount", ascending=False)
        )

        cost_status_summary = (
            detail_df
            .groupby(["Final Sub Project", "Cost Status"], dropna=False)[[
                "Allocated Procurement Amount",
                "Consumed Total Amount",
                "Stock In Hand Total Amount"
            ]]
            .sum()
            .reset_index()
            .sort_values("Allocated Procurement Amount", ascending=False)
        )

        manual_review = detail_df[
            detail_df["Mapping Status"] == "Manual Review Required"
        ].copy()

        stock_in_hand = detail_df[
            detail_df["Stock In Hand Total Amount"] > 0
        ].copy()

        consumed = detail_df[
            detail_df["Consumed Total Amount"] > 0
        ].copy()

        output_sheets = {
            "Allocated Detail": detail_df,
            "Sub Project Summary": sub_project_summary,
            "Cost Status Summary": cost_status_summary,
            "Stock In Hand": stock_in_hand,
            "Consumed": consumed,
            "Manual Review": manual_review
        }

        total_procurement = detail_df["Allocated Procurement Amount"].sum()
        total_consumed = detail_df["Consumed Total Amount"].sum()
        total_stock = detail_df["Stock In Hand Total Amount"].sum()

        st.success("Cost allocation completed successfully.")

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Procurement Cost", f"{total_procurement:,.2f}")
        c2.metric("Consumed Cost", f"{total_consumed:,.2f}")
        c3.metric("Stock In Hand Cost", f"{total_stock:,.2f}")
        c4.metric("Manual Review Rows", f"{len(manual_review):,}")

        st.subheader("Sub Project Summary")
        st.dataframe(sub_project_summary, use_container_width=True)

        st.subheader("Preview")
        st.dataframe(detail_df.head(100), use_container_width=True)

        excel_output = create_excel(output_sheets)

        filename = f"StrategicERP_Cost_Allocation_{datetime.now():%Y%m%d_%H%M}.xlsx"

        st.download_button(
            label="Download Final Excel",
            data=excel_output,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error("Something went wrong while processing the files.")
        st.exception(e)
