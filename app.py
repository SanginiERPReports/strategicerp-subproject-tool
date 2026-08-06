import re
from datetime import datetime
from io import BytesIO

import pandas as pd
import streamlit as st


# ============================================================
# PAGE CONFIGURATION
# ============================================================

st.set_page_config(
    page_title="StrategicERP Cost Intelligence",
    layout="centered"
)

st.title("StrategicERP Cost Intelligence")
st.caption(
    "Upload the required StrategicERP reports and download the final Excel output."
)


# ============================================================
# FILE UPLOADS
# ============================================================

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

gst_file = st.file_uploader(
    "4. Upload Item GST Master Excel",
    type=["xlsx"],
    help=(
        "The GST master should contain Item Description and GST Rate. "
        "Item Group may also be used as a fallback."
    )
)


with st.expander("Excel header-row settings", expanded=False):

    purchase_header_row = st.number_input(
        "Purchase Bill header row",
        min_value=1,
        max_value=10,
        value=2,
        step=1
    )

    stock_header_row = st.number_input(
        "Stock Ledger header row",
        min_value=1,
        max_value=10,
        value=1,
        step=1
    )

    pr_header_row = st.number_input(
        "PR header row",
        min_value=1,
        max_value=10,
        value=1,
        step=1
    )

    gst_header_row = st.number_input(
        "GST Master header row",
        min_value=1,
        max_value=10,
        value=1,
        step=1
    )

    inventory_value_tolerance = st.number_input(
        "Inventory value tolerance",
        min_value=0.00,
        value=1.00,
        step=0.50
    )

    inventory_qty_tolerance = st.number_input(
        "Inventory quantity tolerance",
        min_value=0.000,
        value=0.001,
        step=0.001,
        format="%.3f"
    )


generate_report = st.button(
    "Generate Final Report",
    type="primary",
    use_container_width=True
)


# ============================================================
# GENERAL HELPERS
# ============================================================

def clean_text(value):
    if pd.isna(value):
        return ""

    text = str(value).upper().strip()
    text = re.sub(r"\s+", " ", text)

    return text


def clean_item(value):
    if pd.isna(value):
        return ""

    text = str(value).upper().strip()
    text = re.sub(r"[^A-Z0-9]+", " ", text)
    text = re.sub(r"\s+", " ", text)

    return text.strip()


def clean_subproject(value):
    if pd.isna(value):
        return ""

    text = str(value).upper().strip()
    text = re.sub(r"\s+", " ", text)
    text = re.sub(r",+$", "", text)

    return text.strip()


def clean_item_group(value):
    if pd.isna(value):
        return ""

    text = str(value).upper().strip()
    text = re.sub(r"[^A-Z0-9]+", " ", text)
    text = re.sub(r"\s+", " ", text)

    return text.strip()


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


def to_number(values):
    return pd.to_numeric(
        values,
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


def prepare_dataframe(dataframe):
    dataframe = dataframe.dropna(
        how="all"
    ).copy()

    dataframe.columns = (
        dataframe.columns
        .astype(str)
        .str.strip()
    )

    return dataframe


def read_excel_report(
    uploaded_file,
    header_row
):
    dataframe = pd.read_excel(
        uploaded_file,
        header=int(header_row) - 1
    )

    return prepare_dataframe(dataframe)


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
        candidate_key = str(candidate).strip().upper()

        if candidate_key in column_map:
            return column_map[candidate_key]

    if required:
        raise ValueError(
            "Missing required column. Expected one of: "
            + ", ".join(candidates)
        )

    return None


def select_existing_columns(
    dataframe,
    columns
):
    output = []

    for column in columns:
        if column and column in dataframe.columns:
            output.append(column)

    return list(dict.fromkeys(output))


# ============================================================
# PURCHASE BILL PROCESSING
# ============================================================

def process_purchase_bill(purchase_raw_df):

    purchase_df = purchase_raw_df.copy()

    company_col = find_column(
        purchase_df,
        [
            "Name Of Company",
            "Name of Company",
            "Company Name",
            "Company"
        ],
        required=False
    )

    project_col = find_column(
        purchase_df,
        [
            "Project Name",
            "Project"
        ]
    )

    supplier_col = find_column(
        purchase_df,
        [
            "Supplier Name",
            "Vendor Name",
            "Party Name",
            "Supplier",
            "Vendor"
        ],
        required=False
    )

    pr_col = find_column(
        purchase_df,
        [
            "PRNo",
            "PR No",
            "P.RNo",
            "P.R. No"
        ]
    )

    po_col = find_column(
        purchase_df,
        [
            "PO No",
            "P.O. No",
            "PONo"
        ]
    )

    grn_col = find_column(
        purchase_df,
        [
            "GR No",
            "G.R. No",
            "GRN No"
        ]
    )

    item_group_col = find_column(
        purchase_df,
        [
            "Item Group",
            "Item Category",
            "Material Group"
        ],
        required=False
    )

    item_col = find_column(
        purchase_df,
        [
            "Item Desc",
            "Item Description",
            "Item Name"
        ]
    )

    unit_col = find_column(
        purchase_df,
        [
            "Unit",
            "UOM"
        ],
        required=False
    )

    received_qty_col = find_column(
        purchase_df,
        [
            "Received Qty",
            "GRN Qty",
            "GR Qty",
            "Bill Qty",
            "Quantity"
        ]
    )

    gst_col = find_column(
        purchase_df,
        [
            "GST Amt",
            "GST Amount",
            "Tax Amount"
        ],
        required=False
    )

    bill_amount_col = find_column(
        purchase_df,
        [
            "Bill Item Amt",
            "Bill Item Amount",
            "Bill Amount"
        ]
    )

    purchase_df["PR_Clean"] = (
        purchase_df[pr_col]
        .apply(clean_text)
    )

    purchase_df["PO_Clean"] = (
        purchase_df[po_col]
        .apply(clean_text)
    )

    purchase_df["GRN_Clean"] = (
        purchase_df[grn_col]
        .apply(clean_text)
    )

    purchase_df["Item_Clean"] = (
        purchase_df[item_col]
        .apply(clean_item)
    )

    if item_group_col:
        purchase_df["Item_Group_Clean"] = (
            purchase_df[item_group_col]
            .apply(clean_item_group)
        )
    else:
        purchase_df["Item_Group_Clean"] = ""

    purchase_df["PB Received Qty"] = to_number(
        purchase_df[received_qty_col]
    )

    purchase_df["PB Total Bill Amount"] = to_number(
        purchase_df[bill_amount_col]
    )

    if gst_col:
        purchase_df["PB GST Amount"] = to_number(
            purchase_df[gst_col]
        )
    else:
        purchase_df["PB GST Amount"] = 0.0

    purchase_df["PB Principal Amount"] = (
        purchase_df["PB Total Bill Amount"]
        - purchase_df["PB GST Amount"]
    )

    purchase_df["PB Principal Rate"] = safe_divide(
        purchase_df["PB Principal Amount"],
        purchase_df["PB Received Qty"]
    )

    purchase_df["Purchase Data Status"] = "OK"

    purchase_df.loc[
        purchase_df["PR_Clean"] == "",
        "Purchase Data Status"
    ] = "REVIEW: PR number is blank"

    purchase_df.loc[
        purchase_df["PO_Clean"] == "",
        "Purchase Data Status"
    ] = "REVIEW: PO number is blank"

    purchase_df.loc[
        purchase_df["GRN_Clean"] == "",
        "Purchase Data Status"
    ] = "REVIEW: GRN number is blank"

    purchase_df.loc[
        purchase_df["Item_Clean"] == "",
        "Purchase Data Status"
    ] = "REVIEW: Item description is blank"

    purchase_df.loc[
        purchase_df["PB Received Qty"] <= 0,
        "Purchase Data Status"
    ] = "REVIEW: Received quantity is zero or blank"

    summary_group = [project_col]

    if company_col:
        summary_group.insert(
            0,
            company_col
        )

    if supplier_col:
        summary_group.append(
            supplier_col
        )

    procurement_summary = (
        purchase_df
        .groupby(
            summary_group,
            as_index=False,
            dropna=False
        )
        .agg({
            "PB Received Qty": "sum",
            "PB Principal Amount": "sum",
            "PB GST Amount": "sum",
            "PB Total Bill Amount": "sum"
        })
        .sort_values(
            "PB Principal Amount",
            ascending=False
        )
    )

    purchase_review = purchase_df[
        purchase_df["Purchase Data Status"] != "OK"
    ].copy()

    columns = {
        "company": company_col,
        "project": project_col,
        "supplier": supplier_col,
        "pr": pr_col,
        "po": po_col,
        "grn": grn_col,
        "item_group": item_group_col,
        "item": item_col,
        "unit": unit_col
    }

    return (
        purchase_df,
        procurement_summary,
        purchase_review,
        columns
    )


# ============================================================
# GST MASTER PROCESSING
# ============================================================

def process_gst_master(gst_raw_df):

    gst_df = gst_raw_df.copy()

    gst_item_col = find_column(
        gst_df,
        [
            "Item Desc",
            "Item Description",
            "Item Name",
            "Description"
        ],
        required=False
    )

    gst_item_group_col = find_column(
        gst_df,
        [
            "Item Group",
            "Item Category",
            "Material Group",
            "Group Name"
        ],
        required=False
    )

    gst_rate_col = find_column(
        gst_df,
        [
            "GST %",
            "GST%",
            "GST Rate",
            "GST",
            "Tax Rate",
            "GST Percentage"
        ]
    )

    if not gst_item_col and not gst_item_group_col:
        raise ValueError(
            "GST Master must contain Item Description "
            "or Item Group."
        )

    if gst_item_col:
        gst_df["GST_Item_Clean"] = (
            gst_df[gst_item_col]
            .apply(clean_item)
        )
    else:
        gst_df["GST_Item_Clean"] = ""

    if gst_item_group_col:
        gst_df["GST_Item_Group_Clean"] = (
            gst_df[gst_item_group_col]
            .apply(clean_item_group)
        )
    else:
        gst_df["GST_Item_Group_Clean"] = ""

    gst_rate_text = (
        gst_df[gst_rate_col]
        .astype(str)
        .str.replace("%", "", regex=False)
        .str.extract(r"([-+]?\d*\.?\d+)")[0]
    )

    gst_df["Applicable GST %"] = pd.to_numeric(
        gst_rate_text,
        errors="coerce"
    ).fillna(0)

    item_gst_map = (
        gst_df[
            gst_df["GST_Item_Clean"] != ""
        ]
        .groupby(
            "GST_Item_Clean",
            as_index=False
        )["Applicable GST %"]
        .max()
    )

    group_gst_map = (
        gst_df[
            gst_df["GST_Item_Group_Clean"] != ""
        ]
        .groupby(
            "GST_Item_Group_Clean",
            as_index=False
        )["Applicable GST %"]
        .max()
    )

    return (
        gst_df,
        item_gst_map,
        group_gst_map
    )


def add_gst_to_dataframe(
    dataframe,
    item_gst_map,
    group_gst_map,
    cost_column
):
    output_df = dataframe.copy()

    output_df = output_df.merge(
        item_gst_map.rename(columns={
            "Applicable GST %":
                "Item Level GST %"
        }),
        left_on="Item_Clean",
        right_on="GST_Item_Clean",
        how="left"
    )

    output_df = output_df.drop(
        columns=["GST_Item_Clean"],
        errors="ignore"
    )

    output_df = output_df.merge(
        group_gst_map.rename(columns={
            "Applicable GST %":
                "Group Level GST %"
        }),
        left_on="Item_Group_Clean",
        right_on="GST_Item_Group_Clean",
        how="left"
    )

    output_df = output_df.drop(
        columns=["GST_Item_Group_Clean"],
        errors="ignore"
    )

    output_df["Item Level GST %"] = to_number(
        output_df["Item Level GST %"]
    )

    output_df["Group Level GST %"] = to_number(
        output_df["Group Level GST %"]
    )

    output_df["Applicable GST %"] = (
        output_df["Item Level GST %"]
        .where(
            output_df["Item Level GST %"] > 0,
            output_df["Group Level GST %"]
        )
    )

    output_df["GST Mapping Source"] = "Not Mapped"

    output_df.loc[
        output_df["Item Level GST %"] > 0,
        "GST Mapping Source"
    ] = "Item Description"

    output_df.loc[
        (
            output_df["Item Level GST %"] <= 0
        )
        & (
            output_df["Group Level GST %"] > 0
        ),
        "GST Mapping Source"
    ] = "Item Group"

    output_df["Applicable GST %"] = to_number(
    output_df["Applicable GST %"]
)

output_df["GST Rate Decimal"] = (
    output_df["Applicable GST %"]
    .where(
        output_df["Applicable GST %"] <= 1,
        output_df["Applicable GST %"] / 100
    )
)

output_df["Consumption GST Amount"] = (
    to_number(
        output_df[cost_column]
    )
    * output_df["GST Rate Decimal"]
)

output_df["Total Cost Including GST"] = (
    to_number(
        output_df[cost_column]
    )
    + output_df["Consumption GST Amount"]
)

    return output_df


# ============================================================
# STOCK LEDGER PROCESSING
# ============================================================

def process_stock_ledger(
    stock_raw_df,
    item_gst_map,
    group_gst_map,
    inventory_value_tolerance,
    inventory_qty_tolerance
):

    stock_df = stock_raw_df.copy()

    company_col = find_column(
        stock_df,
        [
            "Name of Company",
            "Name Of Company",
            "Company Name",
            "Company"
        ],
        required=False
    )

    project_col = find_column(
        stock_df,
        [
            "Project Name",
            "Project"
        ]
    )

    subproject_col = find_column(
        stock_df,
        [
            "Sub Project",
            "SubProject",
            "Sub-Project"
        ]
    )

    activity_col = find_column(
        stock_df,
        [
            "Activity Code",
            "Activity"
        ],
        required=False
    )

    contractor_col = find_column(
        stock_df,
        [
            "Contractor / Service Provider Name",
            "Contractor Name",
            "Contractor"
        ],
        required=False
    )

    godown_col = find_column(
        stock_df,
        [
            "Godown Name",
            "Store Name",
            "Godown",
            "Store"
        ],
        required=False
    )

    item_group_col = find_column(
        stock_df,
        [
            "Item Group",
            "Item Category",
            "Material Group"
        ],
        required=False
    )

    item_col = find_column(
        stock_df,
        [
            "Item Desc",
            "Item Description",
            "Item Name"
        ]
    )

    unit_col = find_column(
        stock_df,
        [
            "Unit",
            "UOM"
        ],
        required=False
    )

    voucher_col = find_column(
        stock_df,
        [
            "Voucher No",
            "GIN No",
            "Issue Voucher No"
        ],
        required=False
    )

    pr_col = find_column(
        stock_df,
        [
            "P.RNo",
            "PRNo",
            "PR No"
        ],
        required=False
    )

    po_col = find_column(
        stock_df,
        [
            "P.O. No",
            "PO No",
            "PONo"
        ],
        required=False
    )

    grn_col = find_column(
        stock_df,
        [
            "G.R. No",
            "GR No",
            "GRN No"
        ],
        required=False
    )

    status_col = find_column(
        stock_df,
        [
            "Status",
            "Voucher Status"
        ],
        required=False
    )

    received_qty_col = find_column(
        stock_df,
        [
            "Received Qty",
            "Receipt Qty"
        ],
        required=False
    )

    received_amount_col = find_column(
        stock_df,
        [
            "Received Amt",
            "Received Amount"
        ],
        required=False
    )

    issued_qty_col = find_column(
        stock_df,
        [
            "Issued Qty",
            "Issue Qty",
            "Consumed Qty"
        ]
    )

    issued_amount_col = find_column(
        stock_df,
        [
            "Issued Amt",
            "Issued Amount",
            "Consumed Amount"
        ]
    )

    stock_df["Project_Clean"] = (
        stock_df[project_col]
        .apply(clean_text)
    )

    stock_df["Subproject_Clean"] = (
        stock_df[subproject_col]
        .apply(clean_subproject)
    )

    stock_df["Item_Clean"] = (
        stock_df[item_col]
        .apply(clean_item)
    )

    if item_group_col:
        stock_df["Item_Group_Clean"] = (
            stock_df[item_group_col]
            .apply(clean_item_group)
        )
    else:
        stock_df["Item_Group_Clean"] = ""

    if pr_col:
        stock_df["PR_Clean"] = (
            stock_df[pr_col]
            .apply(clean_text)
        )
    else:
        stock_df["PR_Clean"] = ""

    if po_col:
        stock_df["PO_Clean"] = (
            stock_df[po_col]
            .apply(clean_text)
        )
    else:
        stock_df["PO_Clean"] = ""

    if grn_col:
        stock_df["GRN_Clean"] = (
            stock_df[grn_col]
            .apply(clean_text)
        )
    else:
        stock_df["GRN_Clean"] = ""

    stock_df["Stock Received Qty"] = (
        to_number(
            stock_df[received_qty_col]
        )
        if received_qty_col
        else 0.0
    )

    stock_df["Stock Received Amount"] = (
        to_number(
            stock_df[received_amount_col]
        )
        if received_amount_col
        else 0.0
    )

    stock_df["Stock Issued Qty"] = to_number(
        stock_df[issued_qty_col]
    )

    stock_df["Stock Issued Amount"] = to_number(
        stock_df[issued_amount_col]
    )

    stock_df["Stock Data Status"] = "OK"

    stock_df.loc[
        stock_df["Item_Clean"] == "",
        "Stock Data Status"
    ] = "REVIEW: Item description is blank"

    stock_df.loc[
        (
            stock_df["Stock Issued Qty"] != 0
        )
        & (
            stock_df["Subproject_Clean"] == ""
        ),
        "Stock Data Status"
    ] = "REVIEW: Issued quantity has blank subproject"

    stock_df.loc[
        (
            stock_df["Stock Issued Qty"] != 0
        )
        & (
            stock_df["Stock Issued Amount"] == 0
        ),
        "Stock Data Status"
    ] = "REVIEW: Issued quantity exists but issued amount is zero"

    issue_df = stock_df[
        (
            stock_df["Stock Issued Qty"] != 0
        )
        | (
            stock_df["Stock Issued Amount"] != 0
        )
    ].copy()

    issue_df["Actual Issue Subproject"] = (
        issue_df["Subproject_Clean"]
    )

    issue_df["Actual Consumption Qty"] = (
        issue_df["Stock Issued Qty"]
    )

    issue_df["Actual Consumption Cost"] = (
        issue_df["Stock Issued Amount"]
    )

    issue_df["Actual Consumption Rate"] = safe_divide(
        issue_df["Actual Consumption Cost"],
        issue_df["Actual Consumption Qty"]
    )

    issue_df = add_gst_to_dataframe(
        issue_df,
        item_gst_map,
        group_gst_map,
        "Actual Consumption Cost"
    )

    consumption_columns = select_existing_columns(
        issue_df,
        [
            company_col,
            project_col,
            subproject_col,
            activity_col,
            contractor_col,
            godown_col,
            item_group_col,
            item_col,
            unit_col,
            voucher_col,
            pr_col,
            po_col,
            grn_col,
            status_col,
            "Actual Issue Subproject",
            "Actual Consumption Qty",
            "Actual Consumption Rate",
            "Actual Consumption Cost",
            "Applicable GST %",
            "Consumption GST Amount",
            "Total Cost Including GST",
            "GST Mapping Source",
            "PR_Clean",
            "PO_Clean",
            "GRN_Clean",
            "Item_Clean",
            "Item_Group_Clean",
            "Stock Data Status"
        ]
    )

    consumption_register = issue_df[
        consumption_columns
    ].copy()

    subproject_group = [
        project_col,
        "Actual Issue Subproject"
    ]

    if company_col:
        subproject_group.insert(
            0,
            company_col
        )

    subproject_summary = (
        issue_df
        .groupby(
            subproject_group,
            as_index=False,
            dropna=False
        )
        .agg({
            "Actual Consumption Qty": "sum",
            "Actual Consumption Cost": "sum",
            "Consumption GST Amount": "sum",
            "Total Cost Including GST": "sum"
        })
        .sort_values(
            "Actual Consumption Cost",
            ascending=False
        )
    )

    activity_summary = pd.DataFrame()

    if activity_col:
        activity_summary = (
            issue_df
            .groupby(
                [
                    project_col,
                    "Actual Issue Subproject",
                    activity_col
                ],
                as_index=False,
                dropna=False
            )
            .agg({
                "Actual Consumption Qty": "sum",
                "Actual Consumption Cost": "sum",
                "Consumption GST Amount": "sum",
                "Total Cost Including GST": "sum"
            })
            .sort_values(
                "Actual Consumption Cost",
                ascending=False
            )
        )

    contractor_summary = pd.DataFrame()

    if contractor_col:
        contractor_summary = (
            issue_df
            .groupby(
                [
                    project_col,
                    "Actual Issue Subproject",
                    contractor_col
                ],
                as_index=False,
                dropna=False
            )
            .agg({
                "Actual Consumption Qty": "sum",
                "Actual Consumption Cost": "sum",
                "Consumption GST Amount": "sum",
                "Total Cost Including GST": "sum"
            })
            .sort_values(
                "Actual Consumption Cost",
                ascending=False
            )
        )

    item_group_columns = [
        project_col,
        "Actual Issue Subproject",
        item_col
    ]

    if item_group_col:
        item_group_columns.insert(
            2,
            item_group_col
        )

    if unit_col:
        item_group_columns.append(
            unit_col
        )

    item_summary = (
        issue_df
        .groupby(
            item_group_columns,
            as_index=False,
            dropna=False
        )
        .agg({
            "Actual Consumption Qty": "sum",
            "Actual Consumption Cost": "sum",
            "Consumption GST Amount": "sum",
            "Total Cost Including GST": "sum"
        })
        .sort_values(
            "Actual Consumption Cost",
            ascending=False
        )
    )

    item_summary["Weighted Average Issue Rate"] = safe_divide(
        item_summary["Actual Consumption Cost"],
        item_summary["Actual Consumption Qty"]
    )

    inventory_group = [
        project_col,
        item_col
    ]

    if godown_col:
        inventory_group.insert(
            1,
            godown_col
        )

    if item_group_col:
        inventory_group.append(
            item_group_col
        )

    if unit_col:
        inventory_group.append(
            unit_col
        )

    inventory_summary = (
        stock_df
        .groupby(
            inventory_group,
            as_index=False,
            dropna=False
        )
        .agg({
            "Stock Received Qty": "sum",
            "Stock Issued Qty": "sum",
            "Stock Received Amount": "sum",
            "Stock Issued Amount": "sum"
        })
    )

    inventory_summary["Closing Stock Qty"] = (
        inventory_summary["Stock Received Qty"]
        - inventory_summary["Stock Issued Qty"]
    )

    inventory_summary["Closing Stock Value"] = (
        inventory_summary["Stock Received Amount"]
        - inventory_summary["Stock Issued Amount"]
    )

    inventory_summary["Inventory Status"] = "OK"

    inventory_summary.loc[
        inventory_summary["Closing Stock Qty"]
        < -abs(inventory_qty_tolerance),
        "Inventory Status"
    ] = "REVIEW: Negative closing stock quantity"

    inventory_summary.loc[
        inventory_summary["Closing Stock Value"]
        < -abs(inventory_value_tolerance),
        "Inventory Status"
    ] = "REVIEW: Negative closing stock value"

    issue_review = issue_df[
        issue_df["Stock Data Status"] != "OK"
    ].copy()

    inventory_review = inventory_summary[
        inventory_summary["Inventory Status"] != "OK"
    ].copy()

    missing_gst = issue_df[
        issue_df["Applicable GST %"] <= 0
    ].copy()

    columns = {
        "company": company_col,
        "project": project_col,
        "subproject": subproject_col,
        "activity": activity_col,
        "contractor": contractor_col,
        "item_group": item_group_col,
        "item": item_col
    }

    return (
        stock_df,
        consumption_register,
        subproject_summary,
        activity_summary,
        contractor_summary,
        item_summary,
        inventory_summary,
        issue_review,
        inventory_review,
        missing_gst,
        columns
    )


# ============================================================
# PR PROCESSING
# ============================================================

def process_pr_report(
    pr_raw_df,
    consumption_register
):

    pr_df = pr_raw_df.copy()

    pr_number_col = find_column(
        pr_df,
        [
            "Purchase Requisition (PR) No.",
            "PRNo",
            "PR No",
            "P.RNo"
        ]
    )

    item_col = find_column(
        pr_df,
        [
            "Item Desc",
            "Item Description",
            "Item Name"
        ]
    )

    quantity_col = find_column(
        pr_df,
        [
            "Quantity",
            "PR Qty",
            "Requested Qty"
        ]
    )

    subproject_col = find_column(
        pr_df,
        [
            "Sub Project",
            "SubProject",
            "Sub-Project"
        ]
    )

    pr_df["PR_Clean"] = (
        pr_df[pr_number_col]
        .apply(clean_text)
    )

    pr_df["Item_Clean"] = (
        pr_df[item_col]
        .apply(clean_item)
    )

    pr_df["PR Quantity Numeric"] = to_number(
        pr_df[quantity_col]
    )

    reference_rows = []

    for _, row in pr_df.iterrows():

        subprojects = split_subprojects(
            row[subproject_col]
        )

        if not subprojects:
            subprojects = [""]

        for subproject in subprojects:
            reference_rows.append({
                "PR_Clean":
                    row["PR_Clean"],

                "Item_Clean":
                    row["Item_Clean"],

                "Intended PR Subproject":
                    subproject,

                "PR Reference Quantity":
                    row["PR Quantity Numeric"]
            })

    reference_detail = pd.DataFrame(
        reference_rows
    )

    if reference_detail.empty:
        reference_summary = pd.DataFrame(
            columns=[
                "PR_Clean",
                "Item_Clean",
                "Intended PR Subprojects",
                "PR Reference Quantity"
            ]
        )
    else:
        reference_summary = (
            reference_detail
            .groupby(
                [
                    "PR_Clean",
                    "Item_Clean"
                ],
                as_index=False
            )
            .agg({
                "Intended PR Subproject":
                    join_unique,

                "PR Reference Quantity":
                    "max"
            })
            .rename(columns={
                "Intended PR Subproject":
                    "Intended PR Subprojects"
            })
        )

    consumption_with_pr = (
        consumption_register
        .merge(
            reference_summary,
            on=[
                "PR_Clean",
                "Item_Clean"
            ],
            how="left"
        )
    )

    consumption_with_pr[
        "Intended PR Subprojects"
    ] = (
        consumption_with_pr[
            "Intended PR Subprojects"
        ]
        .fillna("")
    )

    return consumption_with_pr


# ============================================================
# EXCEL OUTPUT
# ============================================================

def create_excel_workbook(sheets):

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
            "align": "center",
            "valign": "top"
        })

        money_format = workbook.add_format({
            "num_format": "₹#,##0.00"
        })

        quantity_format = workbook.add_format({
            "num_format": "#,##0.000"
        })

        percent_format = workbook.add_format({
            "num_format": "0.00%"
        })

        for sheet_name, dataframe in sheets.items():

            if dataframe is None:
                dataframe = pd.DataFrame()

            safe_name = str(
                sheet_name
            )[:31]

            export_df = dataframe.copy()

            export_df.to_excel(
                writer,
                sheet_name=safe_name,
                index=False
            )

            worksheet = writer.sheets[
                safe_name
            ]

            worksheet.freeze_panes(
                1,
                0
            )

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

                lower_name = str(
                    column_name
                ).lower()

                width = min(
                    max(
                        len(str(column_name)) + 4,
                        15
                    ),
                    40
                )

                if "gst %" in lower_name:
                    worksheet.set_column(
                        column_number,
                        column_number,
                        14
                    )

                elif any(
                    word in lower_name
                    for word in [
                        "amount",
                        "cost",
                        "value",
                        "gst",
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

    output.seek(0)

    return output


# ============================================================
# MAIN PROCESS
# ============================================================

if generate_report:

    if not all([
        purchase_file,
        stock_file,
        pr_file,
        gst_file
    ]):
        st.error(
            "Please upload all four Excel files before generating the report."
        )

    else:

        try:

            with st.spinner(
                "Processing reports and creating the final Excel workbook..."
            ):

                purchase_raw_df = read_excel_report(
                    purchase_file,
                    purchase_header_row
                )

                stock_raw_df = read_excel_report(
                    stock_file,
                    stock_header_row
                )

                pr_raw_df = read_excel_report(
                    pr_file,
                    pr_header_row
                )

                gst_raw_df = read_excel_report(
                    gst_file,
                    gst_header_row
                )

                (
                    purchase_register,
                    procurement_summary,
                    purchase_review,
                    purchase_columns
                ) = process_purchase_bill(
                    purchase_raw_df
                )

                (
                    gst_master,
                    item_gst_map,
                    group_gst_map
                ) = process_gst_master(
                    gst_raw_df
                )

                (
                    stock_register,
                    consumption_register,
                    subproject_summary,
                    activity_summary,
                    contractor_summary,
                    item_summary,
                    inventory_summary,
                    stock_issue_review,
                    inventory_review,
                    missing_gst,
                    stock_columns
                ) = process_stock_ledger(
                    stock_raw_df,
                    item_gst_map,
                    group_gst_map,
                    inventory_value_tolerance,
                    inventory_qty_tolerance
                )

                consumption_register = process_pr_report(
                    pr_raw_df,
                    consumption_register
                )

                total_procurement_principal = (
                    purchase_register[
                        "PB Principal Amount"
                    ].sum()
                )

                total_procurement_gst = (
                    purchase_register[
                        "PB GST Amount"
                    ].sum()
                )

                total_purchase_bill = (
                    purchase_register[
                        "PB Total Bill Amount"
                    ].sum()
                )

                total_consumption_excl_gst = (
                    consumption_register[
                        "Actual Consumption Cost"
                    ].sum()
                )

                total_consumption_gst = (
                    consumption_register[
                        "Consumption GST Amount"
                    ].sum()
                )

                total_consumption_incl_gst = (
                    consumption_register[
                        "Total Cost Including GST"
                    ].sum()
                )

                total_received_value = (
                    stock_register[
                        "Stock Received Amount"
                    ].sum()
                )

                total_closing_stock = (
                    inventory_summary[
                        "Closing Stock Value"
                    ].sum()
                )

                dashboard_summary = pd.DataFrame([
                    {
                        "Report Generated On":
                            datetime.now().strftime(
                                "%d-%m-%Y %H:%M"
                            ),

                        "Procurement Principal Cost":
                            total_procurement_principal,

                        "Procurement GST":
                            total_procurement_gst,

                        "Total Purchase Bill Value":
                            total_purchase_bill,

                        "Stock Ledger Received Value":
                            total_received_value,

                        "Consumption Cost Excluding GST":
                            total_consumption_excl_gst,

                        "Consumption GST Amount":
                            total_consumption_gst,

                        "Consumption Cost Including GST":
                            total_consumption_incl_gst,

                        "Closing Stock Value":
                            total_closing_stock,

                        "Consumption Rows":
                            len(consumption_register),

                        "Missing GST Mapping Rows":
                            len(missing_gst),

                        "Purchase Review Rows":
                            len(purchase_review),

                        "Stock Issue Review Rows":
                            len(stock_issue_review),

                        "Inventory Review Rows":
                            len(inventory_review)
                    }
                ])

                audit_frames = []

                if not purchase_review.empty:
                    frame = purchase_review.copy()
                    frame.insert(
                        0,
                        "Audit Source",
                        "Purchase Bill"
                    )
                    audit_frames.append(frame)

                if not stock_issue_review.empty:
                    frame = stock_issue_review.copy()
                    frame.insert(
                        0,
                        "Audit Source",
                        "Stock Ledger Issue"
                    )
                    audit_frames.append(frame)

                if not inventory_review.empty:
                    frame = inventory_review.copy()
                    frame.insert(
                        0,
                        "Audit Source",
                        "Inventory"
                    )
                    audit_frames.append(frame)

                if not missing_gst.empty:
                    frame = missing_gst.copy()
                    frame.insert(
                        0,
                        "Audit Source",
                        "Missing GST Mapping"
                    )
                    audit_frames.append(frame)

                if audit_frames:

                    all_columns = []

                    for frame in audit_frames:
                        for column in frame.columns:
                            if column not in all_columns:
                                all_columns.append(column)

                    aligned_frames = [
                        frame.reindex(
                            columns=all_columns
                        )
                        for frame in audit_frames
                    ]

                    audit_report = pd.concat(
                        aligned_frames,
                        ignore_index=True
                    )

                else:
                    audit_report = pd.DataFrame(
                        columns=[
                            "Audit Source",
                            "Status"
                        ]
                    )

                final_output_sheets = {
                    "Dashboard":
                        dashboard_summary,

                    "Subproject Cost Summary":
                        subproject_summary,

                    "Activity Summary":
                        activity_summary,

                    "Contractor Summary":
                        contractor_summary,

                    "Item Summary":
                        item_summary,

                    "Inventory Summary":
                        inventory_summary,

                    "Consumption Register":
                        consumption_register,

                    "Procurement Register":
                        purchase_register,

                    "Audit Report":
                        audit_report
                }

                final_excel_output = create_excel_workbook(
                    final_output_sheets
                )

                final_filename = (
                    "StrategicERP_Cost_Intelligence_GST_"
                    f"{datetime.now():%Y%m%d_%H%M}.xlsx"
                )

            st.success(
                "Final report generated successfully."
            )

            st.download_button(
                label="Download Final Cost Intelligence Excel",
                data=final_excel_output,
                file_name=final_filename,
                mime=(
                    "application/vnd.openxmlformats-"
                    "officedocument.spreadsheetml.sheet"
                ),
                use_container_width=True
            )

        except Exception as error:

            st.error(
                "The report could not be generated."
            )

            st.exception(error)
