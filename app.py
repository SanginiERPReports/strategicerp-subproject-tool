import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
import re


st.set_page_config(
    page_title="StrategicERP Subproject Consumption",
    layout="wide"
)


st.title("StrategicERP Subproject Consumption Cost")

st.caption(
    "Actual consumption is taken directly from Stock Ledger Issued Amount "
    "and mapped to the Sub Project entered in the Goods Issue Note."
)


purchase_file = st.file_uploader(
    "1. Upload GRN vs Purchase Bill Excel",
    type=["xlsx"]
)


stock_file = st.file_uploader(
    "2. Upload Stock Ledger Excel",
    type=["xlsx"]
)


pr_file = st.file_uploader(
    "3. Upload PR Excel",
    type=["xlsx"]
)


# ============================================================
# HELPER FUNCTIONS
# ============================================================


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

    return value.strip()


def clean_subproject(value):
    if pd.isna(value):
        return ""

    value = str(value).upper().strip()
    value = re.sub(r"\s+", " ", value)
    value = re.sub(r",+$", "", value)

    return value.strip()


def split_subprojects(value):
    if pd.isna(value):
        return []

    text = str(value)

    for separator in ["\n", "|", ";"]:
        text = text.replace(separator, ",")

    output = []

    for part in text.split(","):
        cleaned = clean_subproject(part)

        if cleaned:
            output.append(cleaned)

    return sorted(set(output))


def join_unique(values):
    output = []

    for value in values:
        if pd.notna(value):
            text = str(value).strip()

            if text:
                output.append(text)

    return " | ".join(sorted(set(output)))


def to_number(series):
    return pd.to_numeric(
        series,
        errors="coerce"
    ).fillna(0)


def find_column(
    dataframe,
    candidates,
    required=True
):
    column_map = {
        str(column).strip().upper(): column
        for column in dataframe.columns
    }

    for candidate in candidates:
        key = str(candidate).strip().upper()

        if key in column_map:
            return column_map[key]

    if required:
        raise ValueError(
            "Missing required column. Expected one of: "
            + ", ".join(candidates)
        )

    return None


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

        warning_format = workbook.add_format({
            "bg_color": "#FFF2CC",
            "font_color": "#7F6000"
        })

        error_format = workbook.add_format({
            "bg_color": "#F4CCCC",
            "font_color": "#990000"
        })

        for sheet_name, dataframe in sheets.items():
            if dataframe is None:
                continue

            safe_name = sheet_name[:31]
            export_df = dataframe.copy()

            export_df.to_excel(
                writer,
                sheet_name=safe_name,
                index=False
            )

            worksheet = writer.sheets[safe_name]

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
                    40
                )

                if any(
                    word in lower_name
                    for word in [
                        "amount",
                        "amt",
                        "cost",
                        "value",
                        "gst",
                        "difference",
                        "rate"
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
                        "quantity"
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
                            "review"
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


if purchase_file and stock_file and pr_file:

    try:

        # ----------------------------------------------------
        # READ FILES
        # ----------------------------------------------------

        purchase_df = pd.read_excel(
            purchase_file,
            header=1
        )

        stock_df = pd.read_excel(
            stock_file,
            header=0
        )

        pr_df = pd.read_excel(
            pr_file,
            header=0
        )

        purchase_df = purchase_df.dropna(
            how="all"
        ).copy()

        stock_df = stock_df.dropna(
            how="all"
        ).copy()

        pr_df = pr_df.dropna(
            how="all"
        ).copy()

        purchase_df.columns = (
            purchase_df.columns
            .astype(str)
            .str.strip()
        )

        stock_df.columns = (
            stock_df.columns
            .astype(str)
            .str.strip()
        )

        pr_df.columns = (
            pr_df.columns
            .astype(str)
            .str.strip()
        )

        # ----------------------------------------------------
        # PURCHASE BILL COLUMNS
        # ----------------------------------------------------

        pb_company_col = find_column(
            purchase_df,
            [
                "Name Of Company",
                "Name of Company"
            ],
            required=False
        )

        pb_project_col = find_column(
            purchase_df,
            [
                "Project Name",
                "Project"
            ]
        )

        pb_supplier_col = find_column(
            purchase_df,
            [
                "Supplier Name",
                "Vendor Name",
                "Party Name"
            ],
            required=False
        )

        pb_pr_col = find_column(
            purchase_df,
            [
                "PRNo",
                "PR No",
                "P.RNo"
            ]
        )

        pb_po_col = find_column(
            purchase_df,
            [
                "PO No",
                "P.O. No",
                "PONo"
            ]
        )

        pb_grn_col = find_column(
            purchase_df,
            [
                "GR No",
                "G.R. No",
                "GRN No"
            ]
        )

        pb_item_col = find_column(
            purchase_df,
            [
                "Item Desc",
                "Item Description"
            ]
        )

        pb_unit_col = find_column(
            purchase_df,
            [
                "Unit",
                "UOM"
            ],
            required=False
        )

        pb_received_qty_col = find_column(
            purchase_df,
            [
                "Received Qty",
                "GRN Qty",
                "GR Qty",
                "Quantity"
            ]
        )

        pb_gst_col = find_column(
            purchase_df,
            [
                "GST Amt",
                "GST Amount"
            ],
            required=False
        )

        pb_bill_amount_col = find_column(
            purchase_df,
            [
                "Bill Item Amt",
                "Bill Item Amount",
                "Bill Amount"
            ]
        )

        pb_bill_rate_col = find_column(
            purchase_df,
            [
                "Bill Rate",
                "Rate"
            ],
            required=False
        )

        pb_bill_no_col = find_column(
            purchase_df,
            [
                "Bill No",
                "Invoice No",
                "Purchase Bill No"
            ],
            required=False
        )

        # ----------------------------------------------------
        # STOCK LEDGER COLUMNS
        # ----------------------------------------------------

        stock_date_col = find_column(
            stock_df,
            [
                "Date",
                "Transaction Date"
            ],
            required=False
        )

        stock_company_col = find_column(
            stock_df,
            [
                "Name of Company",
                "Name Of Company"
            ],
            required=False
        )

        stock_project_col = find_column(
            stock_df,
            [
                "Project Name",
                "Project"
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

        stock_activity_col = find_column(
            stock_df,
            [
                "Activity Code"
            ],
            required=False
        )

        stock_contractor_col = find_column(
            stock_df,
            [
                "Contractor / Service Provider Name",
                "Contractor Name"
            ],
            required=False
        )

        stock_godown_col = find_column(
            stock_df,
            [
                "Godown Name",
                "Store Name"
            ],
            required=False
        )

        stock_item_group_col = find_column(
            stock_df,
            [
                "Item Group"
            ],
            required=False
        )

        stock_item_col = find_column(
            stock_df,
            [
                "Item Desc",
                "Item Description"
            ]
        )

        stock_from_voucher_col = find_column(
            stock_df,
            [
                "From Voucher"
            ],
            required=False
        )

        stock_voucher_col = find_column(
            stock_df,
            [
                "Voucher No",
                "GIN No",
                "Issue Voucher No"
            ],
            required=False
        )

        stock_grn_col = find_column(
            stock_df,
            [
                "G.R. No",
                "GR No",
                "GRN No"
            ],
            required=False
        )

        stock_po_col = find_column(
            stock_df,
            [
                "P.O. No",
                "PO No",
                "PONo"
            ],
            required=False
        )

        stock_pr_col = find_column(
            stock_df,
            [
                "P.RNo",
                "PRNo",
                "PR No"
            ],
            required=False
        )

        stock_grn_line_col = find_column(
            stock_df,
            [
                "GRN Line ID"
            ],
            required=False
        )

        stock_unit_col = find_column(
            stock_df,
            [
                "Unit",
                "UOM"
            ],
            required=False
        )

        stock_received_qty_col = find_column(
            stock_df,
            [
                "Received Qty",
                "Receipt Qty"
            ],
            required=False
        )

        stock_issued_qty_col = find_column(
            stock_df,
            [
                "Issued Qty",
                "Issue Qty",
                "Consumed Qty"
            ]
        )

        stock_received_amount_col = find_column(
            stock_df,
            [
                "Received Amt",
                "Received Amount"
            ],
            required=False
        )

        stock_issued_amount_col = find_column(
            stock_df,
            [
                "Issued Amt",
                "Issued Amount"
            ]
        )

        stock_status_col = find_column(
            stock_df,
            [
                "Status"
            ],
            required=False
        )

        # ----------------------------------------------------
        # PR COLUMNS
        # ----------------------------------------------------

        pr_number_col = find_column(
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

        pr_quantity_col = find_column(
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
        # PREPARE PURCHASE BILL DATA
        # ----------------------------------------------------

        purchase_df["PR_Clean"] = (
            purchase_df[pb_pr_col]
            .apply(clean_text)
        )

        purchase_df["PO_Clean"] = (
            purchase_df[pb_po_col]
            .apply(clean_text)
        )

        purchase_df["GRN_Clean"] = (
            purchase_df[pb_grn_col]
            .apply(clean_text)
        )

        purchase_df["Item_Clean"] = (
            purchase_df[pb_item_col]
            .apply(clean_item)
        )

        purchase_df["PB Received Qty"] = to_number(
            purchase_df[pb_received_qty_col]
        )

        purchase_df["PB Bill Amount"] = to_number(
            purchase_df[pb_bill_amount_col]
        )

        if pb_gst_col:
            purchase_df["PB GST Amount"] = to_number(
                purchase_df[pb_gst_col]
            )
        else:
            purchase_df["PB GST Amount"] = 0.0

        purchase_df["PB Principal Amount"] = (
            purchase_df["PB Bill Amount"]
            - purchase_df["PB GST Amount"]
        )

        purchase_df["Purchase Data Status"] = "OK"

        purchase_df.loc[
            (
                (purchase_df["PR_Clean"] == "")
                | (purchase_df["PO_Clean"] == "")
                | (purchase_df["GRN_Clean"] == "")
                | (purchase_df["Item_Clean"] == "")
            ),
            "Purchase Data Status"
        ] = "REVIEW: Missing PR/PO/GRN/Item"

        purchase_df.loc[
            purchase_df["PB Received Qty"] <= 0,
            "Purchase Data Status"
        ] = "REVIEW: Zero or missing received quantity"

        # ----------------------------------------------------
        # PREPARE PR REFERENCE
        # ----------------------------------------------------

        pr_df["PR_Clean"] = (
            pr_df[pr_number_col]
            .apply(clean_text)
        )

        pr_df["Item_Clean"] = (
            pr_df[pr_item_col]
            .apply(clean_item)
        )

        pr_df["PR Quantity"] = to_number(
            pr_df[pr_quantity_col]
        )

        pr_reference_rows = []

        for _, row in pr_df.iterrows():

            pr_number = row["PR_Clean"]
            item = row["Item_Clean"]

            if not pr_number or not item:
                continue

            subprojects = split_subprojects(
                row[pr_subproject_col]
            )

            if not subprojects:
                pr_reference_rows.append({
                    "PR_Clean": pr_number,
                    "Item_Clean": item,
                    "Intended PR Subproject": "",
                    "PR Quantity": row["PR Quantity"]
                })

            else:
                for subproject in subprojects:
                    pr_reference_rows.append({
                        "PR_Clean": pr_number,
                        "Item_Clean": item,
                        "Intended PR Subproject": subproject,
                        "PR Quantity": row["PR Quantity"]
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
                    "Intended PR Subproject": join_unique,
                    "PR Quantity": "max"
                })
                .rename(columns={
                    "Intended PR Subproject":
                        "Intended PR Subprojects",

                    "PR Quantity":
                        "PR Reference Quantity"
                })
            )

        # ----------------------------------------------------
        # PREPARE STOCK LEDGER
        # ----------------------------------------------------

        if stock_pr_col:
            stock_df["PR_Clean"] = (
                stock_df[stock_pr_col]
                .apply(clean_text)
            )
        else:
            stock_df["PR_Clean"] = ""

        if stock_po_col:
            stock_df["PO_Clean"] = (
                stock_df[stock_po_col]
                .apply(clean_text)
            )
        else:
            stock_df["PO_Clean"] = ""

        if stock_grn_col:
            stock_df["GRN_Clean"] = (
                stock_df[stock_grn_col]
                .apply(clean_text)
            )
        else:
            stock_df["GRN_Clean"] = ""

        stock_df["Item_Clean"] = (
            stock_df[stock_item_col]
            .apply(clean_item)
        )

        stock_df["Subproject_Clean"] = (
            stock_df[stock_subproject_col]
            .apply(clean_subproject)
        )

        stock_df["Issued Qty Numeric"] = to_number(
            stock_df[stock_issued_qty_col]
        )

        stock_df["Issued Amount Numeric"] = to_number(
            stock_df[stock_issued_amount_col]
        )

        if stock_received_qty_col:
            stock_df["Received Qty Numeric"] = to_number(
                stock_df[stock_received_qty_col]
            )
        else:
            stock_df["Received Qty Numeric"] = 0.0

        if stock_received_amount_col:
            stock_df["Received Amount Numeric"] = to_number(
                stock_df[stock_received_amount_col]
            )
        else:
            stock_df["Received Amount Numeric"] = 0.0

        stock_df["Stock Data Status"] = "OK"

        stock_df.loc[
            (
                (stock_df["Issued Qty Numeric"] > 0)
                & (stock_df["Subproject_Clean"] == "")
            ),
            "Stock Data Status"
        ] = "REVIEW: Issued quantity has blank subproject"

        stock_df.loc[
            (
                (stock_df["Issued Amount Numeric"] != 0)
                & (stock_df["Issued Qty Numeric"] == 0)
            ),
            "Stock Data Status"
        ] = (
            "REVIEW: Issued amount exists "
            "but issued quantity is zero"
        )

        stock_df.loc[
            (
                (stock_df["Issued Qty Numeric"] != 0)
                & (stock_df["Issued Amount Numeric"] == 0)
            ),
            "Stock Data Status"
        ] = (
            "REVIEW: Issued quantity exists "
            "but issued amount is zero"
        )

        # ----------------------------------------------------
        # ACTUAL CONSUMPTION
        # ----------------------------------------------------

        issue_df = stock_df[
            (
                stock_df["Issued Qty Numeric"] != 0
            )
            | (
                stock_df["Issued Amount Numeric"] != 0
            )
        ].copy()

        issue_df["Actual Issue Subproject"] = (
            issue_df["Subproject_Clean"]
        )

        issue_df["Actual Consumption Qty"] = (
            issue_df["Issued Qty Numeric"]
        )

        issue_df["Actual Consumption Cost"] = (
            issue_df["Issued Amount Numeric"]
        )

        issue_df["ERP Average Issue Rate"] = 0.0

        nonzero_issue_quantity = (
            issue_df["Actual Consumption Qty"] != 0
        )

        issue_df.loc[
            nonzero_issue_quantity,
            "ERP Average Issue Rate"
        ] = (
            issue_df.loc[
                nonzero_issue_quantity,
                "Actual Consumption Cost"
            ]
            / issue_df.loc[
                nonzero_issue_quantity,
                "Actual Consumption Qty"
            ]
        )

        issue_df = issue_df.merge(
            pr_reference_summary,
            on=[
                "PR_Clean",
                "Item_Clean"
            ],
            how="left"
        )

        issue_df["Intended PR Subprojects"] = (
            issue_df["Intended PR Subprojects"]
            .fillna("")
        )

        issue_df["PR Reference Quantity"] = to_number(
            issue_df["PR Reference Quantity"]
        )

        def compare_pr_and_issue(row):
            actual_subproject = str(
                row["Actual Issue Subproject"]
            ).strip()

            intended_text = str(
                row["Intended PR Subprojects"]
            ).strip()

            if not intended_text:
                return "REVIEW: PR + Item reference not found"

            intended_subprojects = [
                part.strip()
                for part in intended_text.split(" | ")
                if part.strip()
            ]

            if actual_subproject in intended_subprojects:
                return "Matched with PR reference"

            return (
                "REVIEW: Issue subproject differs "
                "from PR reference"
            )

        if issue_df.empty:
            issue_df["PR Validation"] = pd.Series(
                dtype=str
            )

        else:
            issue_df["PR Validation"] = issue_df.apply(
                compare_pr_and_issue,
                axis=1
            )

        # ----------------------------------------------------
        # CONSUMPTION DETAIL
        # ----------------------------------------------------

        detail_columns = []

        for column in [
            stock_date_col,
            stock_company_col,
            stock_project_col,
            stock_subproject_col,
            stock_activity_col,
            stock_contractor_col,
            stock_godown_col,
            stock_item_group_col,
            stock_item_col,
            stock_from_voucher_col,
            stock_voucher_col,
            stock_grn_col,
            stock_po_col,
            stock_pr_col,
            stock_grn_line_col,
            stock_unit_col,
            stock_status_col
        ]:
            if column and column in issue_df.columns:
                detail_columns.append(column)

        detail_columns += [
            "Actual Issue Subproject",
            "Actual Consumption Qty",
            "ERP Average Issue Rate",
            "Actual Consumption Cost",
            "Intended PR Subprojects",
            "PR Reference Quantity",
            "PR Validation",
            "Stock Data Status"
        ]

        detail_columns = list(
            dict.fromkeys(detail_columns)
        )

        consumption_detail = issue_df[
            detail_columns
        ].copy()

        # ----------------------------------------------------
        # SUBPROJECT SUMMARY
        # ----------------------------------------------------

        subproject_summary = (
            issue_df
            .groupby(
                [
                    stock_project_col,
                    "Actual Issue Subproject"
                ],
                as_index=False,
                dropna=False
            )[[
                "Actual Consumption Qty",
                "Actual Consumption Cost"
            ]]
            .sum()
        )

        if stock_voucher_col:

            issue_voucher_count = (
                issue_df
                .groupby(
                    [
                        stock_project_col,
                        "Actual Issue Subproject"
                    ],
                    as_index=False,
                    dropna=False
                )[stock_voucher_col]
                .nunique()
                .rename(columns={
                    stock_voucher_col:
                        "Issue Voucher Count"
                })
            )

            subproject_summary = (
                subproject_summary
                .merge(
                    issue_voucher_count,
                    on=[
                        stock_project_col,
                        "Actual Issue Subproject"
                    ],
                    how="left"
                )
            )

        else:
            subproject_summary[
                "Issue Voucher Count"
            ] = 0

        subproject_summary = (
            subproject_summary
            .sort_values(
                "Actual Consumption Cost",
                ascending=False
            )
        )

        # ----------------------------------------------------
        # ACTIVITY SUMMARY
        # ----------------------------------------------------

        if stock_activity_col:

            activity_summary = (
                issue_df
                .groupby(
                    [
                        stock_project_col,
                        "Actual Issue Subproject",
                        stock_activity_col
                    ],
                    as_index=False,
                    dropna=False
                )[[
                    "Actual Consumption Qty",
                    "Actual Consumption Cost"
                ]]
                .sum()
                .sort_values(
                    "Actual Consumption Cost",
                    ascending=False
                )
            )

        else:
            activity_summary = pd.DataFrame()

        # ----------------------------------------------------
        # CONTRACTOR SUMMARY
        # ----------------------------------------------------

        if stock_contractor_col:

            contractor_summary = (
                issue_df
                .groupby(
                    [
                        stock_project_col,
                        "Actual Issue Subproject",
                        stock_contractor_col
                    ],
                    as_index=False,
                    dropna=False
                )[[
                    "Actual Consumption Qty",
                    "Actual Consumption Cost"
                ]]
                .sum()
                .sort_values(
                    "Actual Consumption Cost",
                    ascending=False
                )
            )

        else:
            contractor_summary = pd.DataFrame()

        # ----------------------------------------------------
        # ITEM SUMMARY
        # ----------------------------------------------------

        item_summary = (
            issue_df
            .groupby(
                [
                    stock_project_col,
                    "Actual Issue Subproject",
                    stock_item_col
                ],
                as_index=False,
                dropna=False
            )[[
                "Actual Consumption Qty",
                "Actual Consumption Cost"
            ]]
            .sum()
            .sort_values(
                "Actual Consumption Cost",
                ascending=False
            )
        )

        # ----------------------------------------------------
        # INVENTORY SUMMARY
        # ----------------------------------------------------

        inventory_group_columns = [
            stock_project_col,
            stock_item_col
        ]

        if stock_godown_col:
            inventory_group_columns.insert(
                1,
                stock_godown_col
            )

        inventory_summary = (
            stock_df
            .groupby(
                inventory_group_columns,
                as_index=False,
                dropna=False
            )
            .agg({
                "Received Qty Numeric": "sum",
                "Issued Qty Numeric": "sum",
                "Received Amount Numeric": "sum",
                "Issued Amount Numeric": "sum"
            })
        )

        inventory_summary["Balance Qty"] = (
            inventory_summary["Received Qty Numeric"]
            - inventory_summary["Issued Qty Numeric"]
        )

        inventory_summary["Balance Value"] = (
            inventory_summary["Received Amount Numeric"]
            - inventory_summary["Issued Amount Numeric"]
        )

        inventory_summary["Inventory Status"] = "OK"

        inventory_summary.loc[
            inventory_summary["Balance Qty"] < -0.001,
            "Inventory Status"
        ] = "REVIEW: Negative stock quantity"

        inventory_summary.loc[
            inventory_summary["Balance Value"] < -0.01,
            "Inventory Status"
        ] = "REVIEW: Negative stock value"

        # ----------------------------------------------------
        # PROCUREMENT SUMMARY
        # ----------------------------------------------------

        procurement_group_columns = [
            pb_project_col
        ]

        if pb_supplier_col:
            procurement_group_columns.append(
                pb_supplier_col
            )

        procurement_summary = (
            purchase_df
            .groupby(
                procurement_group_columns,
                as_index=False,
                dropna=False
            )
            .agg({
                "PB Received Qty": "sum",
                "PB Principal Amount": "sum",
                "PB GST Amount": "sum",
                "PB Bill Amount": "sum"
            })
            .sort_values(
                "PB Principal Amount",
                ascending=False
            )
        )

        # ----------------------------------------------------
        # RECONCILIATION
        # ----------------------------------------------------

        total_purchase_bill_amount = (
            purchase_df["PB Bill Amount"].sum()
        )

        total_purchase_principal_amount = (
            purchase_df["PB Principal Amount"].sum()
        )

        total_stock_received_amount = (
            stock_df["Received Amount Numeric"].sum()
        )

        total_stock_issued_amount = (
            stock_df["Issued Amount Numeric"].sum()
        )

        total_stock_balance_value = (
            inventory_summary["Balance Value"].sum()
        )

        reconciliation = pd.DataFrame([
            {
                "Check":
                    "Purchase Bill total vs "
                    "Stock Ledger received amount",

                "Purchase Bill Value":
                    total_purchase_bill_amount,

                "Stock Ledger Value":
                    total_stock_received_amount,

                "Difference":
                    total_purchase_bill_amount
                    - total_stock_received_amount,

                "Comment":
                    "Difference may include GST, opening stock, "
                    "timing, returns, adjustments, or valuation basis."
            },
            {
                "Check":
                    "Stock Ledger movement reconciliation",

                "Purchase Bill Value":
                    total_stock_received_amount,

                "Stock Ledger Value":
                    total_stock_issued_amount
                    + total_stock_balance_value,

                "Difference":
                    total_stock_received_amount
                    - total_stock_issued_amount
                    - total_stock_balance_value,

                "Comment":
                    "Received Amount should equal Issued Amount "
                    "plus Balance Value."
            }
        ])

        reconciliation["Status"] = (
            reconciliation["Difference"]
            .abs()
            .apply(
                lambda value:
                    "Reconciled"
                    if value <= 0.01
                    else "REVIEW REQUIRED"
            )
        )

        # ----------------------------------------------------
        # REVIEW REPORTS
        # ----------------------------------------------------

        issue_review = issue_df[
            (
                issue_df["Stock Data Status"] != "OK"
            )
            | (
                issue_df["PR Validation"]
                .str.contains(
                    "REVIEW",
                    na=False
                )
            )
        ].copy()

        purchase_review = purchase_df[
            purchase_df[
                "Purchase Data Status"
            ] != "OK"
        ].copy()

        inventory_review = inventory_summary[
            inventory_summary[
                "Inventory Status"
            ] != "OK"
        ].copy()

        data_quality = pd.DataFrame([
            {
                "Check":
                    "Stock Ledger issue rows",

                "Total Rows":
                    len(issue_df),

                "Review Rows":
                    len(issue_review),

                "Status":
                    "OK"
                    if len(issue_review) == 0
                    else "REVIEW REQUIRED"
            },
            {
                "Check":
                    "Purchase Bill rows",

                "Total Rows":
                    len(purchase_df),

                "Review Rows":
                    len(purchase_review),

                "Status":
                    "OK"
                    if len(purchase_review) == 0
                    else "REVIEW REQUIRED"
            },
            {
                "Check":
                    "Inventory summary rows",

                "Total Rows":
                    len(inventory_summary),

                "Review Rows":
                    len(inventory_review),

                "Status":
                    "OK"
                    if len(inventory_review) == 0
                    else "REVIEW REQUIRED"
            }
        ])

        # ----------------------------------------------------
        # DISPLAY RESULTS
        # ----------------------------------------------------

        total_consumption_cost = (
            issue_df[
                "Actual Consumption Cost"
            ].sum()
        )

        total_received_value = (
            stock_df[
                "Received Amount Numeric"
            ].sum()
        )

        total_closing_stock_value = (
            inventory_summary[
                "Balance Value"
            ].sum()
        )

        blank_subproject_count = int(
            (
                issue_df[
                    "Actual Issue Subproject"
                ] == ""
            ).sum()
        )

        st.success(
            "Actual subproject consumption report "
            "generated successfully."
        )

        c1, c2, c3, c4 = st.columns(4)

        c1.metric(
            "Actual Consumption Cost",
            f"{total_consumption_cost:,.2f}"
        )

        c2.metric(
            "Stock Ledger Received Value",
            f"{total_received_value:,.2f}"
        )

        c3.metric(
            "Closing Stock Value",
            f"{total_closing_stock_value:,.2f}"
        )

        c4.metric(
            "Blank Issue Subprojects",
            f"{blank_subproject_count:,}"
        )

        st.subheader(
            "Subproject Consumption Summary"
        )

        st.dataframe(
            subproject_summary,
            use_container_width=True,
            hide_index=True
        )

        st.subheader(
            "Inventory Summary"
        )

        st.dataframe(
            inventory_summary,
            use_container_width=True,
            hide_index=True
        )

        st.subheader(
            "Reconciliation"
        )

        st.dataframe(
            reconciliation,
            use_container_width=True,
            hide_index=True
        )

        st.subheader(
            "Data Quality"
        )

        st.dataframe(
            data_quality,
            use_container_width=True,
            hide_index=True
        )

        # ----------------------------------------------------
        # EXCEL OUTPUT
        # ----------------------------------------------------

        output_sheets = {
            "Subproject Consumption":
                subproject_summary,

            "Activity Summary":
                activity_summary,

            "Contractor Summary":
                contractor_summary,

            "Item Summary":
                item_summary,

            "Consumption Detail":
                consumption_detail,

            "Inventory Summary":
                inventory_summary,

            "Procurement Summary":
                procurement_summary,

            "Purchase Bill Detail":
                purchase_df,

            "PR Reference":
                pr_reference_summary,

            "Reconciliation":
                reconciliation,

            "Data Quality":
                data_quality,

            "Issue Review":
                issue_review,

            "Purchase Review":
                purchase_review,

            "Inventory Review":
                inventory_review
        }

        excel_output = create_excel(
            output_sheets
        )

        filename = (
            "StrategicERP_Subproject_Consumption_"
            f"{datetime.now():%Y%m%d_%H%M}.xlsx"
        )

        st.download_button(
            label="Download Final Excel",
            data=excel_output,
            file_name=filename,
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            )
        )

    except Exception as error:

        st.error(
            "Something went wrong while processing the files."
        )

        st.exception(error)


else:

    st.info(
        "Upload the Purchase Bill, Stock Ledger, "
        "and PR Excel files."
    )
