import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
import re

st.set_page_config(
    page_title="StrategicERP Subproject Consumption Cost",
    layout="wide"
)

st.title("StrategicERP Subproject Consumption Cost")
st.write(
    "GRN/Purchase Bill provides actual receipt cost. "
    "Stock Ledger issue quantity and issue subproject provide actual consumption. "
    "PR is used only as a reference and validation source."
)

grn_file = st.file_uploader(
    "Upload GRN vs Purchase Bill Excel",
    type=["xlsx"]
)

pr_file = st.file_uploader(
    "Upload PR Excel",
    type=["xlsx"]
)

stock_file = st.file_uploader(
    "Upload Stock Ledger Excel",
    type=["xlsx"]
)

with st.expander("Cost assumptions", expanded=False):
    bill_includes_gst = st.checkbox(
        "Bill Item Amt includes GST",
        value=True
    )

    include_freight = st.checkbox(
        "Add Freight Chgs separately to material cost",
        value=False
    )


# ============================================================
# HELPERS
# ============================================================

def clean_text(value):
    if pd.isna(value):
        return ""

    value = str(value).upper().strip()
    return re.sub(r"\s+", " ", value)


def clean_item(value):
    if pd.isna(value):
        return ""

    value = str(value).upper().strip()
    value = re.sub(r"[^A-Z0-9]+", " ", value)
    return re.sub(r"\s+", " ", value).strip()


def clean_subproject(value):
    if pd.isna(value):
        return ""

    value = str(value).upper().strip()
    return re.sub(r"\s+", " ", value)


def split_subprojects(value):
    if pd.isna(value):
        return []

    text = str(value)

    for separator in ["\n", "|", ";"]:
        text = text.replace(separator, ",")

    output = []

    for part in text.split(","):
        part = clean_subproject(part)

        if part:
            output.append(part)

    return sorted(set(output))


def join_unique(values):
    output = []

    for value in values:
        if pd.notna(value):
            text = str(value).strip()

            if text:
                output.append(text)

    return " | ".join(sorted(set(output)))


def first_nonblank(values):
    for value in values:
        if pd.notna(value) and str(value).strip():
            return value

    return ""


def find_column(df, candidates, required=True):
    actual_columns = {
        str(column).strip().upper(): column
        for column in df.columns
    }

    for candidate in candidates:
        candidate_key = str(candidate).strip().upper()

        if candidate_key in actual_columns:
            return actual_columns[candidate_key]

    if required:
        raise ValueError(
            "Required column not found. Expected one of: "
            + ", ".join(candidates)
        )

    return None


def to_number(series):
    return pd.to_numeric(
        series,
        errors="coerce"
    ).fillna(0)


def safe_divide(numerator, denominator):
    numerator = to_number(numerator)
    denominator = to_number(denominator)

    result = pd.Series(
        0.0,
        index=numerator.index
    )

    valid = denominator != 0

    result.loc[valid] = (
        numerator.loc[valid]
        / denominator.loc[valid]
    )

    return result


def create_excel(sheets):
    output = BytesIO()

    with pd.ExcelWriter(
        output,
        engine="xlsxwriter"
    ) as writer:

        workbook = writer.book

        header_format = workbook.add_format({
            "bold": True,
            "bg_color": "#1C2551",
            "font_color": "#FFFFFF",
            "border": 1,
            "text_wrap": True,
            "valign": "top"
        })

        money_format = workbook.add_format({
            "num_format": "#,##0.00"
        })

        quantity_format = workbook.add_format({
            "num_format": "#,##0.000"
        })

        percentage_format = workbook.add_format({
            "num_format": "0.00%"
        })

        warning_format = workbook.add_format({
            "bg_color": "#FFF2CC",
            "font_color": "#7F6000"
        })

        error_format = workbook.add_format({
            "bg_color": "#F4CCCC",
            "font_color": "#990000"
        })

        for sheet_name, dataframe in sheets.items():
            sheet_name = sheet_name[:31]
            export_df = dataframe.copy()

            export_df.to_excel(
                writer,
                sheet_name=sheet_name,
                index=False
            )

            worksheet = writer.sheets[sheet_name]

            worksheet.freeze_panes(1, 0)

            if len(export_df.columns) > 0:
                worksheet.autofilter(
                    0,
                    0,
                    len(export_df),
                    len(export_df.columns) - 1
                )

            for column_number, column_name in enumerate(
                export_df.columns
            ):
                worksheet.write(
                    0,
                    column_number,
                    column_name,
                    header_format
                )

                lower_name = str(column_name).lower()

                width = min(
                    max(
                        len(str(column_name)) + 4,
                        15
                    ),
                    38
                )

                if "%" in str(column_name):
                    worksheet.set_column(
                        column_number,
                        column_number,
                        15,
                        percentage_format
                    )

                elif any(
                    word in lower_name
                    for word in [
                        "amount",
                        "cost",
                        "value",
                        "gst",
                        "principal",
                        "freight",
                        "difference"
                    ]
                ):
                    worksheet.set_column(
                        column_number,
                        column_number,
                        18,
                        money_format
                    )

                elif any(
                    word in lower_name
                    for word in [
                        "qty",
                        "quantity",
                        "rate"
                    ]
                ):
                    worksheet.set_column(
                        column_number,
                        column_number,
                        16,
                        quantity_format
                    )

                else:
                    worksheet.set_column(
                        column_number,
                        column_number,
                        width
                    )

                if (
                    any(
                        word in lower_name
                        for word in [
                            "status",
                            "validation",
                            "quality"
                        ]
                    )
                    and len(export_df) > 0
                ):
                    worksheet.conditional_format(
                        1,
                        column_number,
                        len(export_df),
                        column_number,
                        {
                            "type": "text",
                            "criteria": "containing",
                            "value": "REVIEW",
                            "format": warning_format
                        }
                    )

                    worksheet.conditional_format(
                        1,
                        column_number,
                        len(export_df),
                        column_number,
                        {
                            "type": "text",
                            "criteria": "containing",
                            "value": "ERROR",
                            "format": error_format
                        }
                    )

    output.seek(0)

    return output


# ============================================================
# MAIN PROCESS
# ============================================================

if grn_file and pr_file and stock_file:

    try:

        # ----------------------------------------------------
        # READ FILES
        # ----------------------------------------------------

        grn_df = pd.read_excel(
            grn_file,
            header=1
        )

        pr_df = pd.read_excel(
            pr_file,
            header=0
        )

        stock_df = pd.read_excel(
            stock_file,
            header=0
        )

        grn_df = grn_df.dropna(
            how="all"
        ).copy()

        pr_df = pr_df.dropna(
            how="all"
        ).copy()

        stock_df = stock_df.dropna(
            how="all"
        ).copy()

        grn_df.columns = (
            grn_df.columns
            .astype(str)
            .str.strip()
        )

        pr_df.columns = (
            pr_df.columns
            .astype(str)
            .str.strip()
        )

        stock_df.columns = (
            stock_df.columns
            .astype(str)
            .str.strip()
        )

        # ----------------------------------------------------
        # IDENTIFY GRN / PURCHASE BILL COLUMNS
        # ----------------------------------------------------

        grn_pr_col = find_column(
            grn_df,
            [
                "PRNo",
                "PR No",
                "P.RNo",
                "P.R. No"
            ]
        )

        grn_po_col = find_column(
            grn_df,
            [
                "PO No",
                "P.O. No",
                "PONo"
            ]
        )

        grn_no_col = find_column(
            grn_df,
            [
                "GR No",
                "GRN No",
                "G.R. No"
            ]
        )

        grn_item_col = find_column(
            grn_df,
            [
                "Item Desc",
                "Item Description"
            ]
        )

        project_col = find_column(
            grn_df,
            [
                "Project Name",
                "Project"
            ]
        )

        gst_col = find_column(
            grn_df,
            [
                "GST Amt",
                "GST Amount"
            ]
        )

        bill_amount_col = find_column(
            grn_df,
            [
                "Bill Item Amt",
                "Bill Item Amount",
                "Bill Amount"
            ]
        )

        grn_qty_col = find_column(
            grn_df,
            [
                "Received Qty",
                "GRN Qty",
                "GR Qty",
                "Bill Qty",
                "Quantity"
            ]
        )

        freight_col = find_column(
            grn_df,
            [
                "Freight Chgs",
                "Freight Charges",
                "Freight Amt"
            ],
            required=False
        )

        supplier_col = find_column(
            grn_df,
            [
                "Supplier Name",
                "Vendor Name",
                "Party Name"
            ],
            required=False
        )

        bill_no_col = find_column(
            grn_df,
            [
                "Bill No",
                "Invoice No",
                "Purchase Bill No"
            ],
            required=False
        )

        unit_col = find_column(
            grn_df,
            [
                "Unit",
                "UOM"
            ],
            required=False
        )

        # ----------------------------------------------------
        # IDENTIFY PR COLUMNS
        # ----------------------------------------------------

        pr_no_col = find_column(
            pr_df,
            [
                "Purchase Requisition (PR) No.",
                "PRNo",
                "PR No",
                "P.RNo"
            ]
        )

        pr_item_col = find_column(
            pr_df,
            [
                "Item Desc",
                "Item Description"
            ]
        )

        pr_qty_col = find_column(
            pr_df,
            [
                "Quantity",
                "PR Qty",
                "Requested Qty"
            ]
        )

        pr_subproject_col = find_column(
            pr_df,
            [
                "Sub Project",
                "SubProject",
                "Sub-Project"
            ]
        )

        # ----------------------------------------------------
        # IDENTIFY STOCK LEDGER COLUMNS
        # ----------------------------------------------------

        stock_pr_col = find_column(
            stock_df,
            [
                "P.RNo",
                "PRNo",
                "PR No",
                "P.R. No"
            ]
        )

        stock_po_col = find_column(
            stock_df,
            [
                "P.O. No",
                "PO No",
                "PONo"
            ]
        )

        stock_grn_col = find_column(
            stock_df,
            [
                "G.R. No",
                "GR No",
                "GRN No"
            ]
        )

        stock_item_col = find_column(
            stock_df,
            [
                "Item Desc",
                "Item Description"
            ]
        )

        stock_subproject_col = find_column(
            stock_df,
            [
                "Sub Project",
                "SubProject",
                "Sub-Project"
            ]
        )

        stock_issued_qty_col = find_column(
            stock_df,
            [
                "Issued Qty",
                "Issue Qty",
                "Consumed Qty"
            ]
        )

        stock_received_qty_col = find_column(
            stock_df,
            [
                "Received Qty",
                "Receipt Qty"
            ],
            required=False
        )

        # ----------------------------------------------------
        # PREPARE GRN / PURCHASE BILL DATA
        # ----------------------------------------------------

        grn_df["Source PB Row No"] = range(
            1,
            len(grn_df) + 1
        )

        grn_df["PR_Clean"] = (
            grn_df[grn_pr_col]
            .apply(clean_text)
        )

        grn_df["PO_Clean"] = (
            grn_df[grn_po_col]
            .apply(clean_text)
        )

        grn_df["GRN_Clean"] = (
            grn_df[grn_no_col]
            .apply(clean_text)
        )

        grn_df["Item_Clean"] = (
            grn_df[grn_item_col]
            .apply(clean_item)
        )

        grn_df["GRN Received Qty"] = to_number(
            grn_df[grn_qty_col]
        )

        grn_df["Bill Item Amount"] = to_number(
            grn_df[bill_amount_col]
        )

        grn_df["Receipt GST Amount"] = to_number(
            grn_df[gst_col]
        )

        if freight_col:
            grn_df["Freight Amount"] = to_number(
                grn_df[freight_col]
            )
        else:
            grn_df["Freight Amount"] = 0.0

        if bill_includes_gst:
            grn_df["Base Principal Amount"] = (
                grn_df["Bill Item Amount"]
                - grn_df["Receipt GST Amount"]
            )
        else:
            grn_df["Base Principal Amount"] = (
                grn_df["Bill Item Amount"]
            )

        if include_freight:
            grn_df["Receipt Principal Cost"] = (
                grn_df["Base Principal Amount"]
                + grn_df["Freight Amount"]
            )
        else:
            grn_df["Receipt Principal Cost"] = (
                grn_df["Base Principal Amount"]
            )

        grn_df["Receipt Total Including GST"] = (
            grn_df["Receipt Principal Cost"]
            + grn_df["Receipt GST Amount"]
        )

        grn_df["Receipt Key"] = (
            grn_df["PR_Clean"]
            + " || "
            + grn_df["PO_Clean"]
            + " || "
            + grn_df["GRN_Clean"]
            + " || "
            + grn_df["Item_Clean"]
        )

        grn_df["PB Data Quality"] = "OK"

        missing_key = (
            (grn_df["PR_Clean"] == "")
            | (grn_df["PO_Clean"] == "")
            | (grn_df["GRN_Clean"] == "")
            | (grn_df["Item_Clean"] == "")
        )

        grn_df.loc[
            missing_key,
            "PB Data Quality"
        ] = (
            "MANUAL REVIEW: "
            "Missing PR/PO/GRN/Item"
        )

        grn_df.loc[
            grn_df["GRN Received Qty"] <= 0,
            "PB Data Quality"
        ] = (
            "MANUAL REVIEW: "
            "Zero or missing received quantity"
        )

        # ----------------------------------------------------
        # RECEIPT COST TABLE
        # ----------------------------------------------------

        receipt_aggregation = {
            project_col: first_nonblank,
            grn_pr_col: first_nonblank,
            grn_po_col: first_nonblank,
            grn_no_col: first_nonblank,
            grn_item_col: first_nonblank,
            "GRN Received Qty": "sum",
            "Receipt Principal Cost": "sum",
            "Receipt GST Amount": "sum",
            "Receipt Total Including GST": "sum",
            "Freight Amount": "sum",
            "Source PB Row No": (
                lambda x: join_unique(
                    x.astype(str)
                )
            ),
            "PB Data Quality": join_unique
        }

        if supplier_col:
            receipt_aggregation[
                supplier_col
            ] = join_unique

        if bill_no_col:
            receipt_aggregation[
                bill_no_col
            ] = join_unique

        if unit_col:
            receipt_aggregation[
                unit_col
            ] = first_nonblank

        receipt_df = (
            grn_df
            .groupby(
                [
                    "Receipt Key",
                    "PR_Clean",
                    "PO_Clean",
                    "GRN_Clean",
                    "Item_Clean"
                ],
                as_index=False,
                dropna=False
            )
            .agg(receipt_aggregation)
        )

        receipt_df["Principal Unit Cost"] = safe_divide(
            receipt_df["Receipt Principal Cost"],
            receipt_df["GRN Received Qty"]
        )

        receipt_df["GST Unit Cost"] = safe_divide(
            receipt_df["Receipt GST Amount"],
            receipt_df["GRN Received Qty"]
        )

        receipt_df[
            "Total Unit Cost Including GST"
        ] = safe_divide(
            receipt_df[
                "Receipt Total Including GST"
            ],
            receipt_df["GRN Received Qty"]
        )

        # ----------------------------------------------------
        # PREPARE PR REFERENCE
        # PR DOES NOT ALLOCATE CONSUMPTION COST
        # ----------------------------------------------------

        pr_df["PR_Clean"] = (
            pr_df[pr_no_col]
            .apply(clean_text)
        )

        pr_df["Item_Clean"] = (
            pr_df[pr_item_col]
            .apply(clean_item)
        )

        pr_df["PR Quantity"] = to_number(
            pr_df[pr_qty_col]
        )

        pr_reference_rows = []

        for _, row in pr_df.iterrows():

            if (
                not row["PR_Clean"]
                or not row["Item_Clean"]
            ):
                continue

            subprojects = split_subprojects(
                row[pr_subproject_col]
            )

            if not subprojects:
                pr_reference_rows.append({
                    "PR_Clean":
                        row["PR_Clean"],
                    "Item_Clean":
                        row["Item_Clean"],
                    "Intended PR Subproject":
                        "",
                    "PR Quantity":
                        row["PR Quantity"]
                })

            else:
                for subproject in subprojects:
                    pr_reference_rows.append({
                        "PR_Clean":
                            row["PR_Clean"],
                        "Item_Clean":
                            row["Item_Clean"],
                        "Intended PR Subproject":
                            subproject,
                        "PR Quantity":
                            row["PR Quantity"]
                    })

        pr_reference_detail = pd.DataFrame(
            pr_reference_rows,
            columns=[
                "PR_Clean",
                "Item_Clean",
                "Intended PR Subproject",
                "PR Quantity"
            ]
        )

        if pr_reference_detail.empty:

            pr_reference_summary = pd.DataFrame(
                columns=[
                    "PR_Clean",
                    "Item_Clean",
                    "Intended PR Subprojects",
                    "PR Reference Quantity"
                ]
            )

        else:

            pr_reference_summary = (
                pr_reference_detail
                .groupby(
                    [
                        "PR_Clean",
                        "Item_Clean"
                    ],
                    as_index=False,
                    dropna=False
                )
                .agg({
                    "Intended PR Subproject":
                        join_unique,
                    "PR Quantity":
                        "max"
                })
                .rename(columns={
                    "Intended PR Subproject":
                        "Intended PR Subprojects",
                    "PR Quantity":
                        "PR Reference Quantity"
                })
            )

        receipt_df = receipt_df.merge(
            pr_reference_summary,
            on=[
                "PR_Clean",
                "Item_Clean"
            ],
            how="left"
        )

        receipt_df[
            "Intended PR Subprojects"
        ] = (
            receipt_df[
                "Intended PR Subprojects"
            ]
            .fillna("")
        )

        receipt_df[
            "PR Reference Quantity"
        ] = to_number(
            receipt_df[
                "PR Reference Quantity"
            ]
        )

        receipt_df[
            "PR Reference Status"
        ] = "Reference available"

        receipt_df.loc[
            receipt_df[
                "Intended PR Subprojects"
            ] == "",
            "PR Reference Status"
        ] = (
            "MANUAL REVIEW: "
            "PR + Item reference not found"
        )

        # ----------------------------------------------------
        # PREPARE STOCK LEDGER
        # ----------------------------------------------------

        stock_df[
            "Source Stock Row No"
        ] = range(
            1,
            len(stock_df) + 1
        )

        stock_df["PR_Clean"] = (
            stock_df[stock_pr_col]
            .apply(clean_text)
        )

        stock_df["PO_Clean"] = (
            stock_df[stock_po_col]
            .apply(clean_text)
        )

        stock_df["GRN_Clean"] = (
            stock_df[stock_grn_col]
            .apply(clean_text)
        )

        stock_df["Item_Clean"] = (
            stock_df[stock_item_col]
            .apply(clean_item)
        )

        stock_df[
            "Actual Issue Subproject"
        ] = (
            stock_df[stock_subproject_col]
            .apply(clean_subproject)
        )

        stock_df[
            "Actual Issued Qty"
        ] = to_number(
            stock_df[stock_issued_qty_col]
        )

        if stock_received_qty_col:
            stock_df[
                "Ledger Received Qty"
            ] = to_number(
                stock_df[
                    stock_received_qty_col
                ]
            )
        else:
            stock_df[
                "Ledger Received Qty"
            ] = 0.0

        stock_df["Receipt Key"] = (
            stock_df["PR_Clean"]
            + " || "
            + stock_df["PO_Clean"]
            + " || "
            + stock_df["GRN_Clean"]
            + " || "
            + stock_df["Item_Clean"]
        )

        stock_df[
            "Stock Data Quality"
        ] = "OK"

        stock_missing_key = (
            (stock_df["PR_Clean"] == "")
            | (stock_df["PO_Clean"] == "")
            | (stock_df["GRN_Clean"] == "")
            | (stock_df["Item_Clean"] == "")
        )

        stock_df.loc[
            stock_missing_key,
            "Stock Data Quality"
        ] = (
            "MANUAL REVIEW: "
            "Missing PR/PO/GRN/Item"
        )

        stock_df.loc[
            (
                (
                    stock_df[
                        "Actual Issued Qty"
                    ] > 0
                )
                & (
                    stock_df[
                        "Actual Issue Subproject"
                    ] == ""
                )
            ),
            "Stock Data Quality"
        ] = (
            "MANUAL REVIEW: "
            "Issue subproject is blank"
        )

        issue_rows = stock_df[
            stock_df[
                "Actual Issued Qty"
            ] != 0
        ].copy()

        # ----------------------------------------------------
        # ACTUAL ISSUE SUMMARY
        # ----------------------------------------------------

        if issue_rows.empty:

            issue_summary = pd.DataFrame(
                columns=[
                    "Receipt Key",
                    "PR_Clean",
                    "PO_Clean",
                    "GRN_Clean",
                    "Item_Clean",
                    "Actual Issue Subproject",
                    "Actual Issued Qty",
                    "Source Stock Row Nos",
                    "Stock Data Quality"
                ]
            )

        else:

            issue_summary = (
                issue_rows
                .groupby(
                    [
                        "Receipt Key",
                        "PR_Clean",
                        "PO_Clean",
                        "GRN_Clean",
                        "Item_Clean",
                        "Actual Issue Subproject"
                    ],
                    as_index=False,
                    dropna=False
                )
                .agg({
                    "Actual Issued Qty":
                        "sum",
                    "Source Stock Row No":
                        lambda x: join_unique(
                            x.astype(str)
                        ),
                    "Stock Data Quality":
                        join_unique
                })
                .rename(columns={
                    "Source Stock Row No":
                        "Source Stock Row Nos"
                })
            )

        # Stock Ledger may repeat received quantity
        # for every issue row. Therefore MAX is used.

        ledger_validation = (
            stock_df
            .groupby(
                [
                    "Receipt Key",
                    "PR_Clean",
                    "PO_Clean",
                    "GRN_Clean",
                    "Item_Clean"
                ],
                as_index=False,
                dropna=False
            )
            .agg({
                "Ledger Received Qty":
                    "max",
                "Source Stock Row No":
                    "count"
            })
            .rename(columns={
                "Source Stock Row No":
                    "Stock Ledger Row Count"
            })
        )

        # ----------------------------------------------------
        # COST ACTUAL CONSUMPTION
        # ----------------------------------------------------

        consumption_df = issue_summary.merge(
            receipt_df,
            on=[
                "Receipt Key",
                "PR_Clean",
                "PO_Clean",
                "GRN_Clean",
                "Item_Clean"
            ],
            how="left",
            indicator=True
        )

        consumption_df[
            "Receipt Match Status"
        ] = (
            "Matched to GRN/Purchase Bill"
        )

        consumption_df.loc[
            consumption_df["_merge"] != "both",
            "Receipt Match Status"
        ] = (
            "MANUAL REVIEW: "
            "Stock issue not matched "
            "to GRN/Purchase Bill"
        )

        consumption_df = (
            consumption_df
            .drop(columns=["_merge"])
        )

        for column in [
            "GRN Received Qty",
            "Receipt Principal Cost",
            "Receipt GST Amount",
            "Receipt Total Including GST",
            "Principal Unit Cost",
            "GST Unit Cost",
            "Total Unit Cost Including GST",
            "PR Reference Quantity"
        ]:
            if column in consumption_df.columns:
                consumption_df[column] = to_number(
                    consumption_df[column]
                )

        if not consumption_df.empty:

            consumption_df[
                "Total Actual Issued Qty for Receipt"
            ] = (
                consumption_df
                .groupby(
                    "Receipt Key"
                )["Actual Issued Qty"]
                .transform("sum")
            )

        else:

            consumption_df[
                "Total Actual Issued Qty for Receipt"
            ] = pd.Series(dtype=float)

        consumption_df[
            "Costing Factor"
        ] = 1.0

        over_issue_mask = (
            (
                consumption_df[
                    "Total Actual Issued Qty for Receipt"
                ]
                > consumption_df[
                    "GRN Received Qty"
                ]
            )
            & (
                consumption_df[
                    "Total Actual Issued Qty for Receipt"
                ] > 0
            )
        )

        consumption_df.loc[
            over_issue_mask,
            "Costing Factor"
        ] = (
            consumption_df.loc[
                over_issue_mask,
                "GRN Received Qty"
            ]
            / consumption_df.loc[
                over_issue_mask,
                "Total Actual Issued Qty for Receipt"
            ]
        )

        consumption_df[
            "Costed Issued Qty"
        ] = (
            consumption_df[
                "Actual Issued Qty"
            ]
            * consumption_df[
                "Costing Factor"
            ]
        )

        consumption_df[
            "Uncosted Over-Issue Qty"
        ] = (
            consumption_df[
                "Actual Issued Qty"
            ]
            - consumption_df[
                "Costed Issued Qty"
            ]
        )

        consumption_df[
            "Consumed Principal Cost"
        ] = (
            consumption_df[
                "Costed Issued Qty"
            ]
            * consumption_df[
                "Principal Unit Cost"
            ]
        )

        consumption_df[
            "Consumed GST Value"
        ] = (
            consumption_df[
                "Costed Issued Qty"
            ]
            * consumption_df[
                "GST Unit Cost"
            ]
        )

        consumption_df[
            "Consumed Total Including GST"
        ] = (
            consumption_df[
                "Costed Issued Qty"
            ]
            * consumption_df[
                "Total Unit Cost Including GST"
            ]
        )

        consumption_df[
            "Consumption Status"
        ] = "Actual issue costed"

        consumption_df.loc[
            (
                consumption_df[
                    "Actual Issue Subproject"
                ]
                .fillna("")
                == ""
            ),
            "Consumption Status"
        ] = (
            "MANUAL REVIEW: "
            "Blank actual issue subproject"
        )

        consumption_df.loc[
            consumption_df[
                "Receipt Match Status"
            ]
            .str.contains(
                "MANUAL REVIEW",
                na=False
            ),
            "Consumption Status"
        ] = (
            "MANUAL REVIEW: "
            "Receipt cost unavailable"
        )

        consumption_df.loc[
            consumption_df[
                "Uncosted Over-Issue Qty"
            ] > 0,
            "Consumption Status"
        ] = (
            "MANUAL REVIEW: "
            "Issue exceeds receipt; "
            "value proportionately capped"
        )

        def validate_pr_subproject(row):
            actual = str(
                row.get(
                    "Actual Issue Subproject",
                    ""
                )
            ).strip()

            intended_text = str(
                row.get(
                    "Intended PR Subprojects",
                    ""
                )
            ).strip()

            if not intended_text:
                return (
                    "REVIEW: "
                    "PR reference unavailable"
                )

            intended = [
                item.strip()
                for item in intended_text.split(" | ")
                if item.strip()
            ]

            if actual in intended:
                return (
                    "Actual issue agrees "
                    "with PR reference"
                )

            return (
                "REVIEW: Actual issue "
                "subproject differs "
                "from PR reference"
            )

        if not consumption_df.empty:

            consumption_df[
                "PR vs Actual Issue Validation"
            ] = consumption_df.apply(
                validate_pr_subproject,
                axis=1
            )

        else:

            consumption_df[
                "PR vs Actual Issue Validation"
            ] = pd.Series(dtype=str)

        # ----------------------------------------------------
        # STOCK IN HAND BY RECEIPT
        # ----------------------------------------------------

        if issue_rows.empty:

            issued_by_receipt = pd.DataFrame(
                columns=[
                    "Receipt Key",
                    "PR_Clean",
                    "PO_Clean",
                    "GRN_Clean",
                    "Item_Clean",
                    "Total Actual Issued Qty"
                ]
            )

        else:

            issued_by_receipt = (
                issue_rows
                .groupby(
                    [
                        "Receipt Key",
                        "PR_Clean",
                        "PO_Clean",
                        "GRN_Clean",
                        "Item_Clean"
                    ],
                    as_index=False,
                    dropna=False
                )["Actual Issued Qty"]
                .sum()
                .rename(columns={
                    "Actual Issued Qty":
                        "Total Actual Issued Qty"
                })
            )

        stock_in_hand_df = receipt_df.merge(
            issued_by_receipt,
            on=[
                "Receipt Key",
                "PR_Clean",
                "PO_Clean",
                "GRN_Clean",
                "Item_Clean"
            ],
            how="left"
        )

        stock_in_hand_df[
            "Total Actual Issued Qty"
        ] = to_number(
            stock_in_hand_df[
                "Total Actual Issued Qty"
            ]
        )

        stock_in_hand_df = (
            stock_in_hand_df
            .merge(
                ledger_validation,
                on=[
                    "Receipt Key",
                    "PR_Clean",
                    "PO_Clean",
                    "GRN_Clean",
                    "Item_Clean"
                ],
                how="left"
            )
        )

        stock_in_hand_df[
            "Ledger Received Qty"
        ] = to_number(
            stock_in_hand_df[
                "Ledger Received Qty"
            ]
        )

        stock_in_hand_df[
            "Stock Ledger Row Count"
        ] = to_number(
            stock_in_hand_df[
                "Stock Ledger Row Count"
            ]
        )

        stock_in_hand_df[
            "Costed Issued Qty"
        ] = (
            stock_in_hand_df[
                [
                    "GRN Received Qty",
                    "Total Actual Issued Qty"
                ]
            ]
            .min(axis=1)
            .clip(lower=0)
        )

        stock_in_hand_df[
            "Stock In Hand Qty"
        ] = (
            stock_in_hand_df[
                "GRN Received Qty"
            ]
            - stock_in_hand_df[
                "Costed Issued Qty"
            ]
        ).clip(lower=0)

        stock_in_hand_df[
            "Over-Issue Qty"
        ] = (
            stock_in_hand_df[
                "Total Actual Issued Qty"
            ]
            - stock_in_hand_df[
                "GRN Received Qty"
            ]
        ).clip(lower=0)

        stock_in_hand_df[
            "Consumed Principal Cost"
        ] = (
            stock_in_hand_df[
                "Costed Issued Qty"
            ]
            * stock_in_hand_df[
                "Principal Unit Cost"
            ]
        )

        stock_in_hand_df[
            "Consumed GST Value"
        ] = (
            stock_in_hand_df[
                "Costed Issued Qty"
            ]
            * stock_in_hand_df[
                "GST Unit Cost"
            ]
        )

        stock_in_hand_df[
            "Consumed Total Including GST"
        ] = (
            stock_in_hand_df[
                "Costed Issued Qty"
            ]
            * stock_in_hand_df[
                "Total Unit Cost Including GST"
            ]
        )

        stock_in_hand_df[
            "Stock In Hand Principal Cost"
        ] = (
            stock_in_hand_df[
                "Stock In Hand Qty"
            ]
            * stock_in_hand_df[
                "Principal Unit Cost"
            ]
        )

        stock_in_hand_df[
            "Stock In Hand GST Value"
        ] = (
            stock_in_hand_df[
                "Stock In Hand Qty"
            ]
            * stock_in_hand_df[
                "GST Unit Cost"
            ]
        )

        stock_in_hand_df[
            "Stock In Hand Total Including GST"
        ] = (
            stock_in_hand_df[
                "Stock In Hand Qty"
            ]
            * stock_in_hand_df[
                "Total Unit Cost Including GST"
            ]
        )

        stock_in_hand_df[
            "Quantity Validation"
        ] = "OK"

        stock_in_hand_df.loc[
            stock_in_hand_df[
                "Over-Issue Qty"
            ] > 0,
            "Quantity Validation"
        ] = (
            "MANUAL REVIEW: "
            "Issued quantity exceeds "
            "GRN received quantity"
        )

        ledger_qty_available = (
            stock_in_hand_df[
                "Ledger Received Qty"
            ] > 0
        )

        quantity_difference = (
            stock_in_hand_df[
                "Ledger Received Qty"
            ]
            - stock_in_hand_df[
                "GRN Received Qty"
            ]
        ).abs()

        quantity_tolerance = (
            stock_in_hand_df[
                "GRN Received Qty"
            ]
            .abs()
            * 0.001
        ).clip(lower=0.001)

        stock_in_hand_df.loc[
            (
                ledger_qty_available
                & (
                    quantity_difference
                    > quantity_tolerance
                )
            ),
            "Quantity Validation"
        ] = (
            "REVIEW: Stock Ledger "
            "received quantity differs "
            "from GRN quantity"
        )

        stock_in_hand_df[
            "Cost Reconciliation Difference"
        ] = (
            stock_in_hand_df[
                "Receipt Principal Cost"
            ]
            - stock_in_hand_df[
                "Consumed Principal Cost"
            ]
            - stock_in_hand_df[
                "Stock In Hand Principal Cost"
            ]
        )

        stock_in_hand_df[
            "Cost Reconciliation Status"
        ] = "Reconciled"

        stock_in_hand_df.loc[
            stock_in_hand_df[
                "Cost Reconciliation Difference"
            ].abs() > 0.01,
            "Cost Reconciliation Status"
        ] = (
            "ERROR: Receipt cost "
            "does not reconcile"
        )

        # ----------------------------------------------------
        # SUMMARIES
        # ----------------------------------------------------

        valid_consumption = consumption_df[
            consumption_df[
                "Actual Issue Subproject"
            ]
            .fillna("")
            .astype(str)
            .str.strip()
            != ""
        ].copy()

        if valid_consumption.empty:

            subproject_summary = pd.DataFrame(
                columns=[
                    "Actual Issue Subproject",
                    "Actual Issued Qty",
                    "Costed Issued Qty",
                    "Uncosted Over-Issue Qty",
                    "Consumed Principal Cost",
                    "Consumed GST Value",
                    "Consumed Total Including GST"
                ]
            )

        else:

            subproject_summary = (
                valid_consumption
                .groupby(
                    "Actual Issue Subproject",
                    as_index=False,
                    dropna=False
                )[[
                    "Actual Issued Qty",
                    "Costed Issued Qty",
                    "Uncosted Over-Issue Qty",
                    "Consumed Principal Cost",
                    "Consumed GST Value",
                    "Consumed Total Including GST"
                ]]
                .sum()
                .sort_values(
                    "Consumed Principal Cost",
                    ascending=False
                )
            )

        if valid_consumption.empty:

            project_subproject_summary = (
                pd.DataFrame()
            )

        else:

            project_subproject_summary = (
                valid_consumption
                .groupby(
                    [
                        project_col,
                        "Actual Issue Subproject"
                    ],
                    as_index=False,
                    dropna=False
                )[[
                    "Actual Issued Qty",
                    "Costed Issued Qty",
                    "Consumed Principal Cost",
                    "Consumed GST Value",
                    "Consumed Total Including GST"
                ]]
                .sum()
                .sort_values(
                    "Consumed Principal Cost",
                    ascending=False
                )
            )

        stock_available_df = stock_in_hand_df[
            stock_in_hand_df[
                "Stock In Hand Qty"
            ] > 0
        ].copy()

        # ----------------------------------------------------
        # RECONCILIATION
        # ----------------------------------------------------

        total_receipt_principal = (
            receipt_df[
                "Receipt Principal Cost"
            ].sum()
        )

        total_receipt_gst = (
            receipt_df[
                "Receipt GST Amount"
            ].sum()
        )

        total_receipt_total = (
            receipt_df[
                "Receipt Total Including GST"
            ].sum()
        )

        total_consumed_principal = (
            stock_in_hand_df[
                "Consumed Principal Cost"
            ].sum()
        )

        total_consumed_gst = (
            stock_in_hand_df[
                "Consumed GST Value"
            ].sum()
        )

        total_consumed_total = (
            stock_in_hand_df[
                "Consumed Total Including GST"
            ].sum()
        )

        total_stock_principal = (
            stock_in_hand_df[
                "Stock In Hand Principal Cost"
            ].sum()
        )

        total_stock_gst = (
            stock_in_hand_df[
                "Stock In Hand GST Value"
            ].sum()
        )

        total_stock_total = (
            stock_in_hand_df[
                "Stock In Hand Total Including GST"
            ].sum()
        )

        reconciliation_df = pd.DataFrame([
            {
                "Measure":
                    "Principal / Material Cost",
                "Receipt Value":
                    total_receipt_principal,
                "Consumed Value":
                    total_consumed_principal,
                "Stock In Hand Value":
                    total_stock_principal,
                "Reconciliation Difference":
                    total_receipt_principal
                    - total_consumed_principal
                    - total_stock_principal
            },
            {
                "Measure":
                    "GST Value",
                "Receipt Value":
                    total_receipt_gst,
                "Consumed Value":
                    total_consumed_gst,
                "Stock In Hand Value":
                    total_stock_gst,
                "Reconciliation Difference":
                    total_receipt_gst
                    - total_consumed_gst
                    - total_stock_gst
            },
            {
                "Measure":
                    "Total Including GST",
                "Receipt Value":
                    total_receipt_total,
                "Consumed Value":
                    total_consumed_total,
                "Stock In Hand Value":
                    total_stock_total,
                "Reconciliation Difference":
                    total_receipt_total
                    - total_consumed_total
                    - total_stock_total
            }
        ])

        reconciliation_df["Status"] = (
            reconciliation_df[
                "Reconciliation Difference"
            ]
            .abs()
            .apply(
                lambda value:
                    "Reconciled"
                    if value <= 0.01
                    else "ERROR: Not reconciled"
            )
        )

        # ----------------------------------------------------
        # REVIEW TABLES
        # ----------------------------------------------------

        pb_review = grn_df[
            grn_df[
                "PB Data Quality"
            ] != "OK"
        ].copy()

        receipt_review = stock_in_hand_df[
            (
                stock_in_hand_df[
                    "Quantity Validation"
                ] != "OK"
            )
            | (
                stock_in_hand_df[
                    "Cost Reconciliation Status"
                ] != "Reconciled"
            )
            | (
                stock_in_hand_df[
                    "PR Reference Status"
                ]
                .str.contains(
                    "MANUAL REVIEW",
                    na=False
                )
            )
        ].copy()

        consumption_review = consumption_df[
            (
                consumption_df[
                    "Consumption Status"
                ]
                .str.contains(
                    "MANUAL REVIEW",
                    na=False
                )
            )
            | (
                consumption_df[
                    "PR vs Actual Issue Validation"
                ]
                .str.contains(
                    "REVIEW",
                    na=False
                )
            )
        ].copy()

        unmatched_stock_issues = consumption_df[
            consumption_df[
                "Receipt Match Status"
            ]
            .str.contains(
                "MANUAL REVIEW",
                na=False
            )
        ].copy()

        data_quality_summary = pd.DataFrame([
            {
                "Check":
                    "Purchase Bill source rows",
                "Total Rows":
                    len(grn_df),
                "Review Rows":
                    len(pb_review)
            },
            {
                "Check":
                    "Receipt cost records",
                "Total Rows":
                    len(receipt_df),
                "Review Rows":
                    len(receipt_review)
            },
            {
                "Check":
                    "Consumption records",
                "Total Rows":
                    len(consumption_df),
                "Review Rows":
                    len(consumption_review)
            },
            {
                "Check":
                    "Unmatched stock issues",
                "Total Rows":
                    len(consumption_df),
                "Review Rows":
                    len(unmatched_stock_issues)
            },
            {
                "Check":
                    "Over-issued receipts",
                "Total Rows":
                    len(stock_in_hand_df),
                "Review Rows":
                    int(
                        (
                            stock_in_hand_df[
                                "Over-Issue Qty"
                            ] > 0
                        ).sum()
                    )
            }
        ])

        data_quality_summary["Status"] = (
            data_quality_summary[
                "Review Rows"
            ]
            .apply(
                lambda count:
                    "OK"
                    if count == 0
                    else "REVIEW REQUIRED"
            )
        )

        # ----------------------------------------------------
        # EXCEL OUTPUT
        # ----------------------------------------------------

        output_sheets = {
            "Reconciliation":
                reconciliation_df,

            "Subproject Consumption":
                subproject_summary,

            "Project Subproject":
                project_subproject_summary,

            "Consumption Detail":
                consumption_df,

            "Stock In Hand":
                stock_available_df,

            "Receipt Cost Detail":
                receipt_df,

            "PR Reference":
                pr_reference_summary,

            "Data Quality":
                data_quality_summary,

            "Consumption Review":
                consumption_review,

            "Receipt Review":
                receipt_review,

            "Unmatched Stock Issues":
                unmatched_stock_issues,

            "PB Source Review":
                pb_review
        }

        excel_output = create_excel(
            output_sheets
        )

        # ----------------------------------------------------
        # DISPLAY RESULTS
        # ----------------------------------------------------

        principal_difference = (
            total_receipt_principal
            - total_consumed_principal
            - total_stock_principal
        )

        st.success(
            "Actual subproject consumption "
            "costing completed."
        )

        c1, c2, c3, c4 = st.columns(4)

        c1.metric(
            "Receipt Principal Cost",
            f"{total_receipt_principal:,.2f}"
        )

        c2.metric(
            "Consumed Principal Cost",
            f"{total_consumed_principal:,.2f}"
        )

        c3.metric(
            "Stock In Hand Principal Cost",
            f"{total_stock_principal:,.2f}"
        )

        c4.metric(
            "Reconciliation Difference",
            f"{principal_difference:,.2f}"
        )

        c5, c6, c7, c8 = st.columns(4)

        c5.metric(
            "Receipt Records",
            f"{len(receipt_df):,}"
        )

        c6.metric(
            "Consumption Records",
            f"{len(consumption_df):,}"
        )

        c7.metric(
            "Review Rows",
            f"{len(consumption_review):,}"
        )

        over_issue_count = int(
    (
        stock_in_hand_df["Over-Issue Qty"] > 0
    ).sum()
)

c8.metric(
    "Over-Issue Records",
    f"{over_issue_count:,}"
)
        st.subheader("Reconciliation")

        st.dataframe(
            reconciliation_df,
            use_container_width=True,
            hide_index=True
        )

        st.subheader(
            "Actual Subproject Consumption Summary"
        )

        st.dataframe(
            subproject_summary,
            use_container_width=True,
            hide_index=True
        )

        if not project_subproject_summary.empty:

            st.subheader(
                "Project and Subproject Summary"
            )

            st.dataframe(
                project_subproject_summary,
                use_container_width=True,
                hide_index=True
            )

        st.subheader(
            "Stock In Hand Preview"
        )

        st.dataframe(
            stock_available_df.head(100),
            use_container_width=True,
            hide_index=True
        )

        st.subheader(
            "Data Quality Summary"
        )

        st.dataframe(
            data_quality_summary,
            use_container_width=True,
            hide_index=True
        )

        output_filename = (
            "StrategicERP_Actual_"
            "Subproject_Consumption_"
            f"{datetime.now():%Y%m%d_%H%M}.xlsx"
        )

        st.download_button(
            label="Download Final Excel",
            data=excel_output,
            file_name=output_filename,
            mime=(
                "application/vnd.openxmlformats-"
                "officedocument.spreadsheetml.sheet"
            )
        )

    except Exception as error:

        st.error(
            "Something went wrong while "
            "processing the files."
        )

        st.exception(error)

else:

    st.info(
        "Upload all three Excel files "
        "to generate the report."
    )
