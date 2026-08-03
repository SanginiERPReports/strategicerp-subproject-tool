import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
import re


# ============================================================
# STREAMLIT PAGE CONFIGURATION
# ============================================================

st.set_page_config(
    page_title="StrategicERP Cost Intelligence",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)


# ============================================================
# APPLICATION HEADER
# ============================================================

st.title("StrategicERP Cost Intelligence")

st.caption(
    "Procurement, actual subproject consumption, inventory, "
    "activity, contractor and item-wise cost reporting."
)

st.info(
    """
    **Business methodology**

    - Purchase Bill report is used for procurement reporting.
    - Stock Ledger `Issued Amt` is used as actual consumption cost.
    - Stock Ledger issue `Sub Project` is used as the consuming subproject.
    - PR report is used only as an intended-use reference.
    - Closing inventory is calculated as Received minus Issued.
    """
)


# ============================================================
# FILE UPLOAD SECTION
# ============================================================

st.sidebar.header("Upload StrategicERP Reports")

purchase_file = st.sidebar.file_uploader(
    "1. GRN vs Purchase Bill",
    type=["xlsx"],
    key="purchase_file"
)

stock_file = st.sidebar.file_uploader(
    "2. Stock Ledger",
    type=["xlsx"],
    key="stock_file"
)

pr_file = st.sidebar.file_uploader(
    "3. Purchase Requisition Report",
    type=["xlsx"],
    key="pr_file"
)


# ============================================================
# APPLICATION SETTINGS
# ============================================================

st.sidebar.header("Report Settings")

purchase_header_row = st.sidebar.number_input(
    "Purchase Bill header row",
    min_value=1,
    max_value=10,
    value=2,
    step=1,
    help=(
        "Use 2 when the Purchase Bill export contains one title row "
        "before the actual column headers."
    )
)

stock_header_row = st.sidebar.number_input(
    "Stock Ledger header row",
    min_value=1,
    max_value=10,
    value=1,
    step=1
)

pr_header_row = st.sidebar.number_input(
    "PR report header row",
    min_value=1,
    max_value=10,
    value=1,
    step=1
)

inventory_value_tolerance = st.sidebar.number_input(
    "Inventory value tolerance",
    min_value=0.00,
    value=1.00,
    step=0.50,
    help=(
        "Negative stock values within this amount will be treated "
        "as insignificant rounding differences."
    )
)

inventory_qty_tolerance = st.sidebar.number_input(
    "Inventory quantity tolerance",
    min_value=0.000,
    value=0.001,
    step=0.001,
    format="%.3f"
)


# ============================================================
# TEXT CLEANING FUNCTIONS
# ============================================================

def clean_text(value):
    """
    Standardise document numbers and general text.

    Example:
    '  pr-001  ' becomes 'PR-001'
    """
    if pd.isna(value):
        return ""

    text = str(value).upper().strip()
    text = re.sub(r"\s+", " ", text)

    return text


def clean_item(value):
    """
    Standardise item descriptions for comparison.

    Example:
    'Cement - OPC 53' becomes 'CEMENT OPC 53'
    """
    if pd.isna(value):
        return ""

    text = str(value).upper().strip()
    text = re.sub(r"[^A-Z0-9]+", " ", text)
    text = re.sub(r"\s+", " ", text)

    return text.strip()


def clean_subproject(value):
    """
    Standardise subproject names without removing useful characters.
    """
    if pd.isna(value):
        return ""

    text = str(value).upper().strip()
    text = re.sub(r"\s+", " ", text)
    text = re.sub(r",+$", "", text)

    return text.strip()


def split_subprojects(value):
    """
    Convert a PR cell containing multiple subprojects into a list.

    Supported separators:
    comma, semicolon, pipe and line break.
    """
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


# ============================================================
# NUMERIC FUNCTIONS
# ============================================================

def to_number(values):
    """
    Convert a pandas Series to numeric values.

    Invalid or blank values become zero.
    """
    return pd.to_numeric(
        values,
        errors="coerce"
    ).fillna(0)


def safe_divide(numerator, denominator):
    """
    Divide two pandas Series without division-by-zero errors.
    """
    numerator = to_number(numerator)
    denominator = to_number(denominator)

    result = pd.Series(
        0.0,
        index=numerator.index
    )

    valid_rows = denominator != 0

    result.loc[valid_rows] = (
        numerator.loc[valid_rows]
        / denominator.loc[valid_rows]
    )

    return result


# ============================================================
# GENERAL DATA FUNCTIONS
# ============================================================

def join_unique(values):
    """
    Join unique nonblank values using a pipe separator.
    """
    output = []

    for value in values:
        if pd.notna(value):
            text = str(value).strip()

            if text:
                output.append(text)

    return " | ".join(sorted(set(output)))


def first_nonblank(values):
    """
    Return the first nonblank value from a pandas Series.
    """
    for value in values:
        if pd.notna(value) and str(value).strip():
            return value

    return ""


def find_column(
    dataframe,
    candidates,
    required=True
):
    """
    Find an Excel column using a list of possible column names.

    Matching is case-insensitive and ignores leading/trailing spaces.
    """
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


def prepare_dataframe(dataframe):
    """
    Apply basic cleaning to every imported Excel report.
    """
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
    """
    Read and clean an uploaded Excel report.

    Streamlit header rows are entered starting from 1.
    Pandas header rows start from 0.
    """
    dataframe = pd.read_excel(
        uploaded_file,
        header=int(header_row) - 1
    )

    return prepare_dataframe(dataframe)


# ============================================================
# DISPLAY FUNCTIONS
# ============================================================

def show_metric(
    container,
    label,
    value,
    value_type="money"
):
    """
    Display a consistently formatted Streamlit metric.
    """
    if value_type == "money":
        formatted_value = f"₹{value:,.2f}"

    elif value_type == "integer":
        formatted_value = f"{int(value):,}"

    elif value_type == "quantity":
        formatted_value = f"{value:,.3f}"

    else:
        formatted_value = str(value)

    container.metric(
        label,
        formatted_value
    )


def show_dataframe(
    title,
    dataframe,
    maximum_rows=None
):
    """
    Display a dataframe using consistent Streamlit formatting.
    """
    st.subheader(title)

    if dataframe is None or dataframe.empty:
        st.warning(
            f"No data available for {title}."
        )
        return

    display_df = dataframe.copy()

    if maximum_rows:
        display_df = display_df.head(
            maximum_rows
        )

    st.dataframe(
        display_df,
        use_container_width=True,
        hide_index=True
    )


# ============================================================
# EXCEL EXPORT FUNCTION
# ============================================================

def create_excel_workbook(sheets):
    """
    Create a professionally formatted multi-sheet Excel workbook.

    Expected input:
    {
        "Sheet Name": dataframe,
        "Another Sheet": dataframe
    }
    """
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
            "valign": "top",
            "align": "center"
        })

        money_format = workbook.add_format({
            "num_format": "₹#,##0.00"
        })

        quantity_format = workbook.add_format({
            "num_format": "#,##0.000"
        })

        integer_format = workbook.add_format({
            "num_format": "#,##0"
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

            if dataframe is None:
                continue

            safe_sheet_name = str(
                sheet_name
            )[:31]

            export_df = dataframe.copy()

            export_df.to_excel(
                writer,
                sheet_name=safe_sheet_name,
                index=False
            )

            worksheet = writer.sheets[
                safe_sheet_name
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

                column_width = min(
                    max(
                        len(str(column_name)) + 4,
                        15
                    ),
                    42
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

                elif any(
                    word in lower_name
                    for word in [
                        "count",
                        "number"
                    ]
                ):
                    worksheet.set_column(
                        column_number,
                        column_number,
                        14,
                        integer_format
                    )

                else:
                    worksheet.set_column(
                        column_number,
                        column_number,
                        column_width
                    )

                if (
                    len(export_df) > 0
                    and any(
                        word in lower_name
                        for word in [
                            "status",
                            "validation",
                            "review",
                            "quality"
                        ]
                    )
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
# FILE STATUS DISPLAY
# ============================================================

st.sidebar.divider()
st.sidebar.subheader("Upload Status")

if purchase_file:
    st.sidebar.success(
        "Purchase Bill uploaded"
    )
else:
    st.sidebar.warning(
        "Purchase Bill pending"
    )

if stock_file:
    st.sidebar.success(
        "Stock Ledger uploaded"
    )
else:
    st.sidebar.warning(
        "Stock Ledger pending"
    )

if pr_file:
    st.sidebar.success(
        "PR Report uploaded"
    )
else:
    st.sidebar.warning(
        "PR Report pending"
    )


# ============================================================
# MAIN APPLICATION PLACEHOLDER
# ============================================================

if not (
    purchase_file
    and stock_file
    and pr_file
):
    st.warning(
        "Upload all three StrategicERP Excel reports "
        "from the left sidebar to continue."
    )

    st.stop()


try:
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

    st.success(
        "All three Excel reports were read successfully."
    )

    file_c1, file_c2, file_c3 = st.columns(3)

    show_metric(
        file_c1,
        "Purchase Bill Rows",
        len(purchase_raw_df),
        "integer"
    )

    show_metric(
        file_c2,
        "Stock Ledger Rows",
        len(stock_raw_df),
        "integer"
    )

    show_metric(
        file_c3,
        "PR Rows",
        len(pr_raw_df),
        "integer"
    )

    with st.expander(
        "Preview uploaded reports",
        expanded=False
    ):
        show_dataframe(
            "Purchase Bill Preview",
            purchase_raw_df,
            maximum_rows=20
        )

        show_dataframe(
            "Stock Ledger Preview",
            stock_raw_df,
            maximum_rows=20
        )

        show_dataframe(
            "PR Report Preview",
            pr_raw_df,
            maximum_rows=20
        )

    st.info(
        "Module 1 is working. "
        "Module 2 code should be pasted immediately below this line."
    )

except Exception as error:
    st.error(
        "The uploaded reports could not be read."
    )

    st.exception(error)

    st.stop()
# ============================================================
# MODULE 2 — PURCHASE BILL / PROCUREMENT ENGINE
# ============================================================


def process_purchase_bill(purchase_raw_df):
    """
    Process the GRN vs Purchase Bill export.

    The function preserves all original ERP columns and appends:
    - Clean PR, PO, GRN and Item fields
    - Numeric received quantity
    - Principal amount
    - GST amount
    - Total bill amount
    - Calculated principal rate
    - Data-quality status

    Returns:
    1. Detailed procurement register
    2. Project and supplier summary
    3. Project and item summary
    4. Purchase Bill review report
    5. Dictionary containing detected column names
    """

    purchase_df = purchase_raw_df.copy()

    # --------------------------------------------------------
    # FIND REQUIRED PURCHASE BILL COLUMNS
    # --------------------------------------------------------

    project_col = find_column(
        purchase_df,
        [
            "Project Name",
            "Project"
        ]
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
            "PONo",
            "Purchase Order No"
        ]
    )

    grn_col = find_column(
        purchase_df,
        [
            "GR No",
            "G.R. No",
            "GRN No",
            "Goods Receipt No"
        ]
    )

    item_col = find_column(
        purchase_df,
        [
            "Item Desc",
            "Item Description",
            "Item Name"
        ]
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

    bill_amount_col = find_column(
        purchase_df,
        [
            "Bill Item Amt",
            "Bill Item Amount",
            "Bill Amount",
            "Total Bill Amount"
        ]
    )

    # --------------------------------------------------------
    # FIND OPTIONAL PURCHASE BILL COLUMNS
    # --------------------------------------------------------

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

    unit_col = find_column(
        purchase_df,
        [
            "Unit",
            "UOM",
            "U.O.M."
        ],
        required=False
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

    bill_rate_col = find_column(
        purchase_df,
        [
            "Bill Rate",
            "Purchase Rate",
            "Rate"
        ],
        required=False
    )

    bill_no_col = find_column(
        purchase_df,
        [
            "Bill No",
            "Invoice No",
            "Purchase Bill No",
            "Supplier Bill No"
        ],
        required=False
    )

    bill_date_col = find_column(
        purchase_df,
        [
            "Bill Date",
            "Invoice Date",
            "Purchase Bill Date"
        ],
        required=False
    )

    po_date_col = find_column(
        purchase_df,
        [
            "PO Date",
            "Purchase Order Date"
        ],
        required=False
    )

    grn_date_col = find_column(
        purchase_df,
        [
            "GR Date",
            "GRN Date",
            "Goods Receipt Date"
        ],
        required=False
    )

    freight_col = find_column(
        purchase_df,
        [
            "Freight Chgs",
            "Freight Charges",
            "Freight Amt",
            "Freight Amount"
        ],
        required=False
    )

    loading_col = find_column(
        purchase_df,
        [
            "Loading / Unloading Chgs",
            "Loading Unloading Chgs",
            "Loading Charges",
            "Unloading Charges"
        ],
        required=False
    )

    other_charges_col = find_column(
        purchase_df,
        [
            "Others Chgs",
            "Other Charges",
            "Others Charges"
        ],
        required=False
    )

    # --------------------------------------------------------
    # CLEAN MATCHING FIELDS
    # --------------------------------------------------------

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

    # --------------------------------------------------------
    # CONVERT FINANCIAL AND QUANTITY FIELDS
    # --------------------------------------------------------

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

    if freight_col:
        purchase_df["PB Freight Amount"] = to_number(
            purchase_df[freight_col]
        )
    else:
        purchase_df["PB Freight Amount"] = 0.0

    if loading_col:
        purchase_df["PB Loading Unloading Amount"] = to_number(
            purchase_df[loading_col]
        )
    else:
        purchase_df["PB Loading Unloading Amount"] = 0.0

    if other_charges_col:
        purchase_df["PB Other Charges"] = to_number(
            purchase_df[other_charges_col]
        )
    else:
        purchase_df["PB Other Charges"] = 0.0

    # Assumption based on the earlier StrategicERP export:
    # Bill Item Amount contains GST.
    purchase_df["PB Principal Amount"] = (
        purchase_df["PB Total Bill Amount"]
        - purchase_df["PB GST Amount"]
    )

    purchase_df["PB Calculated Principal Rate"] = safe_divide(
        purchase_df["PB Principal Amount"],
        purchase_df["PB Received Qty"]
    )

    purchase_df["PB Calculated Total Rate"] = safe_divide(
        purchase_df["PB Total Bill Amount"],
        purchase_df["PB Received Qty"]
    )

    if bill_rate_col:
        purchase_df["PB ERP Bill Rate"] = to_number(
            purchase_df[bill_rate_col]
        )
    else:
        purchase_df["PB ERP Bill Rate"] = 0.0

    purchase_df["PB Rate Difference"] = (
        purchase_df["PB ERP Bill Rate"]
        - purchase_df["PB Calculated Principal Rate"]
    )

    # --------------------------------------------------------
    # CREATE TRACEABILITY KEY
    # --------------------------------------------------------

    purchase_df["Procurement Trace Key"] = (
        purchase_df["PR_Clean"]
        + " || "
        + purchase_df["PO_Clean"]
        + " || "
        + purchase_df["GRN_Clean"]
        + " || "
        + purchase_df["Item_Clean"]
    )

    # --------------------------------------------------------
    # DATA QUALITY CHECKS
    # --------------------------------------------------------

    purchase_df["Purchase Data Status"] = "OK"

    missing_pr = purchase_df["PR_Clean"] == ""
    missing_po = purchase_df["PO_Clean"] == ""
    missing_grn = purchase_df["GRN_Clean"] == ""
    missing_item = purchase_df["Item_Clean"] == ""

    purchase_df.loc[
        missing_pr,
        "Purchase Data Status"
    ] = "REVIEW: PR number is blank"

    purchase_df.loc[
        missing_po,
        "Purchase Data Status"
    ] = "REVIEW: PO number is blank"

    purchase_df.loc[
        missing_grn,
        "Purchase Data Status"
    ] = "REVIEW: GRN number is blank"

    purchase_df.loc[
        missing_item,
        "Purchase Data Status"
    ] = "REVIEW: Item description is blank"

    purchase_df.loc[
        purchase_df["PB Received Qty"] <= 0,
        "Purchase Data Status"
    ] = "REVIEW: Received quantity is zero or blank"

    purchase_df.loc[
        purchase_df["PB Total Bill Amount"] < 0,
        "Purchase Data Status"
    ] = "REVIEW: Negative Purchase Bill amount"

    purchase_df.loc[
        purchase_df["PB Principal Amount"] < 0,
        "Purchase Data Status"
    ] = "REVIEW: GST exceeds total Bill amount"

    # Flag duplicate procurement keys.
    purchase_df["Procurement Key Row Count"] = (
        purchase_df
        .groupby(
            "Procurement Trace Key"
        )["Procurement Trace Key"]
        .transform("count")
    )

    purchase_df["Duplicate Key Status"] = "Unique key"

    purchase_df.loc[
        purchase_df["Procurement Key Row Count"] > 1,
        "Duplicate Key Status"
    ] = (
        "INFORMATIONAL: Multiple Purchase Bill rows "
        "exist for the same PR/PO/GRN/Item"
    )

    # --------------------------------------------------------
    # PROJECT AND SUPPLIER SUMMARY
    # --------------------------------------------------------

    project_supplier_group = [
        project_col
    ]

    if company_col:
        project_supplier_group.insert(
            0,
            company_col
        )

    if supplier_col:
        project_supplier_group.append(
            supplier_col
        )

    procurement_summary = (
        purchase_df
        .groupby(
            project_supplier_group,
            as_index=False,
            dropna=False
        )
        .agg({
            "PB Received Qty": "sum",
            "PB Principal Amount": "sum",
            "PB GST Amount": "sum",
            "PB Total Bill Amount": "sum",
            "PB Freight Amount": "sum",
            "PB Loading Unloading Amount": "sum",
            "PB Other Charges": "sum",
            "PR_Clean": pd.Series.nunique,
            "PO_Clean": pd.Series.nunique,
            "GRN_Clean": pd.Series.nunique,
            "Item_Clean": pd.Series.nunique
        })
        .rename(columns={
            "PR_Clean": "Unique PR Count",
            "PO_Clean": "Unique PO Count",
            "GRN_Clean": "Unique GRN Count",
            "Item_Clean": "Unique Item Count"
        })
        .sort_values(
            "PB Principal Amount",
            ascending=False
        )
    )

    # --------------------------------------------------------
    # PROJECT AND ITEM SUMMARY
    # --------------------------------------------------------

    project_item_group = [
        project_col,
        item_col
    ]

    if unit_col:
        project_item_group.append(
            unit_col
        )

    procurement_item_summary = (
        purchase_df
        .groupby(
            project_item_group,
            as_index=False,
            dropna=False
        )
        .agg({
            "PB Received Qty": "sum",
            "PB Principal Amount": "sum",
            "PB GST Amount": "sum",
            "PB Total Bill Amount": "sum",
            "PR_Clean": pd.Series.nunique,
            "PO_Clean": pd.Series.nunique,
            "GRN_Clean": pd.Series.nunique
        })
        .rename(columns={
            "PR_Clean": "Unique PR Count",
            "PO_Clean": "Unique PO Count",
            "GRN_Clean": "Unique GRN Count"
        })
        .sort_values(
            "PB Principal Amount",
            ascending=False
        )
    )

    procurement_item_summary[
        "Weighted Principal Rate"
    ] = safe_divide(
        procurement_item_summary[
            "PB Principal Amount"
        ],
        procurement_item_summary[
            "PB Received Qty"
        ]
    )

    # --------------------------------------------------------
    # PURCHASE BILL REVIEW REPORT
    # --------------------------------------------------------

    purchase_review = purchase_df[
        purchase_df["Purchase Data Status"] != "OK"
    ].copy()

    # --------------------------------------------------------
    # COLUMN INFORMATION FOR LATER MODULES
    # --------------------------------------------------------

    detected_columns = {
        "company": company_col,
        "project": project_col,
        "supplier": supplier_col,
        "pr": pr_col,
        "po": po_col,
        "grn": grn_col,
        "item": item_col,
        "unit": unit_col,
        "received_qty": received_qty_col,
        "gst": gst_col,
        "bill_amount": bill_amount_col,
        "bill_rate": bill_rate_col,
        "bill_no": bill_no_col,
        "bill_date": bill_date_col,
        "po_date": po_date_col,
        "grn_date": grn_date_col,
        "freight": freight_col,
        "loading": loading_col,
        "other_charges": other_charges_col
    }

    return (
        purchase_df,
        procurement_summary,
        procurement_item_summary,
        purchase_review,
        detected_columns
    )


# ============================================================
# RUN MODULE 2
# ============================================================

try:
    (
        purchase_register,
        procurement_summary,
        procurement_item_summary,
        purchase_review,
        purchase_columns
    ) = process_purchase_bill(
        purchase_raw_df
    )

    st.divider()
    st.header("Module 2 — Procurement")

    total_purchase_principal = (
        purchase_register[
            "PB Principal Amount"
        ].sum()
    )

    total_purchase_gst = (
        purchase_register[
            "PB GST Amount"
        ].sum()
    )

    total_purchase_bill = (
        purchase_register[
            "PB Total Bill Amount"
        ].sum()
    )

    total_purchase_qty = (
        purchase_register[
            "PB Received Qty"
        ].sum()
    )

    procurement_c1, procurement_c2, procurement_c3, procurement_c4 = (
        st.columns(4)
    )

    show_metric(
        procurement_c1,
        "Procurement Principal Cost",
        total_purchase_principal,
        "money"
    )

    show_metric(
        procurement_c2,
        "Procurement GST",
        total_purchase_gst,
        "money"
    )

    show_metric(
        procurement_c3,
        "Total Purchase Bill Value",
        total_purchase_bill,
        "money"
    )

    show_metric(
        procurement_c4,
        "Total Received Quantity",
        total_purchase_qty,
        "quantity"
    )

    review_c1, review_c2, review_c3 = st.columns(3)

    show_metric(
        review_c1,
        "Purchase Bill Rows",
        len(purchase_register),
        "integer"
    )

    show_metric(
        review_c2,
        "Purchase Review Rows",
        len(purchase_review),
        "integer"
    )

    show_metric(
        review_c3,
        "Unique Procurement Keys",
        purchase_register[
            "Procurement Trace Key"
        ].nunique(),
        "integer"
    )

    show_dataframe(
        "Procurement Summary by Project and Supplier",
        procurement_summary,
        maximum_rows=100
    )

    show_dataframe(
        "Procurement Summary by Project and Item",
        procurement_item_summary,
        maximum_rows=100
    )

    with st.expander(
        "Preview Detailed Procurement Register",
        expanded=False
    ):
        show_dataframe(
            "Purchase Register",
            purchase_register,
            maximum_rows=100
        )

    if not purchase_review.empty:
        with st.expander(
            "Purchase Bill Review Required",
            expanded=False
        ):
            show_dataframe(
                "Purchase Review",
                purchase_review,
                maximum_rows=100
            )

    st.success(
        "Module 2 completed successfully. "
        "The detailed procurement register and summaries are ready."
    )

except Exception as error:
    st.error(
        "Module 2 could not process the Purchase Bill report."
    )

    st.exception(error)

    st.stop()
    # ============================================================
# MODULE 3 — STOCK LEDGER, CONSUMPTION AND INVENTORY ENGINE
# ============================================================


def process_stock_ledger(
    stock_raw_df,
    inventory_value_tolerance=1.00,
    inventory_qty_tolerance=0.001
):
    """
    Process the StrategicERP Stock Ledger.

    Business logic:
    - Issued Amt is treated as actual consumption cost.
    - Issue Sub Project is treated as the actual consuming subproject.
    - Received rows are treated as inventory receipts.
    - Closing inventory = Received - Issued.
    - PR allocation is not used in this module.

    Returns:
    1. Cleaned stock ledger register
    2. Actual consumption register
    3. Subproject summary
    4. Activity summary
    5. Contractor summary
    6. Item consumption summary
    7. Inventory summary
    8. Stock issue review
    9. Inventory review
    10. Detected column dictionary
    """

    stock_df = stock_raw_df.copy()

    # --------------------------------------------------------
    # FIND REQUIRED STOCK LEDGER COLUMNS
    # --------------------------------------------------------

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

    item_col = find_column(
        stock_df,
        [
            "Item Desc",
            "Item Description",
            "Item Name"
        ]
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
            "Issue Amount",
            "Consumed Amount"
        ]
    )

    # --------------------------------------------------------
    # FIND OPTIONAL STOCK LEDGER COLUMNS
    # --------------------------------------------------------

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

    date_col = find_column(
        stock_df,
        [
            "Date",
            "Transaction Date",
            "Voucher Date"
        ],
        required=False
    )

    activity_col = find_column(
        stock_df,
        [
            "Activity Code",
            "Activity",
            "Work Activity"
        ],
        required=False
    )

    contractor_col = find_column(
        stock_df,
        [
            "Contractor / Service Provider Name",
            "Contractor Name",
            "Service Provider Name",
            "Contractor",
            "Service Provider"
        ],
        required=False
    )

    godown_col = find_column(
        stock_df,
        [
            "Godown Name",
            "Store Name",
            "Warehouse Name",
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

    unit_col = find_column(
        stock_df,
        [
            "Unit",
            "UOM",
            "U.O.M."
        ],
        required=False
    )

    voucher_col = find_column(
        stock_df,
        [
            "Voucher No",
            "GIN No",
            "Issue Voucher No",
            "Goods Issue No",
            "Voucher Number"
        ],
        required=False
    )

    from_voucher_col = find_column(
        stock_df,
        [
            "From Voucher",
            "Source Voucher"
        ],
        required=False
    )

    status_col = find_column(
        stock_df,
        [
            "Status",
            "Voucher Status",
            "Transaction Status"
        ],
        required=False
    )

    pr_col = find_column(
        stock_df,
        [
            "P.RNo",
            "PRNo",
            "PR No",
            "P.R. No",
            "Purchase Requisition No"
        ],
        required=False
    )

    po_col = find_column(
        stock_df,
        [
            "P.O. No",
            "PO No",
            "PONo",
            "Purchase Order No"
        ],
        required=False
    )

    grn_col = find_column(
        stock_df,
        [
            "G.R. No",
            "GR No",
            "GRN No",
            "Goods Receipt No"
        ],
        required=False
    )

    grn_line_col = find_column(
        stock_df,
        [
            "GRN Line ID",
            "GR Line ID",
            "Receipt Line ID"
        ],
        required=False
    )

    received_qty_col = find_column(
        stock_df,
        [
            "Received Qty",
            "Receipt Qty",
            "GRN Qty",
            "Inward Qty"
        ],
        required=False
    )

    received_amount_col = find_column(
        stock_df,
        [
            "Received Amt",
            "Received Amount",
            "Receipt Amount",
            "Inward Amount"
        ],
        required=False
    )

    # --------------------------------------------------------
    # ADD SOURCE ROW NUMBER
    # --------------------------------------------------------

    stock_df["Source Stock Row No"] = range(
        1,
        len(stock_df) + 1
    )

    # --------------------------------------------------------
    # CLEAN IMPORTANT TEXT FIELDS
    # --------------------------------------------------------

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

    if activity_col:
        stock_df["Activity_Clean"] = (
            stock_df[activity_col]
            .apply(clean_text)
        )
    else:
        stock_df["Activity_Clean"] = ""

    if contractor_col:
        stock_df["Contractor_Clean"] = (
            stock_df[contractor_col]
            .apply(clean_text)
        )
    else:
        stock_df["Contractor_Clean"] = ""

    if godown_col:
        stock_df["Godown_Clean"] = (
            stock_df[godown_col]
            .apply(clean_text)
        )
    else:
        stock_df["Godown_Clean"] = ""

    if item_group_col:
        stock_df["Item_Group_Clean"] = (
            stock_df[item_group_col]
            .apply(clean_text)
        )
    else:
        stock_df["Item_Group_Clean"] = ""

    if voucher_col:
        stock_df["Voucher_Clean"] = (
            stock_df[voucher_col]
            .apply(clean_text)
        )
    else:
        stock_df["Voucher_Clean"] = ""

    # --------------------------------------------------------
    # CONVERT QUANTITY AND VALUE COLUMNS
    # --------------------------------------------------------

    stock_df["Stock Issued Qty"] = to_number(
        stock_df[issued_qty_col]
    )

    stock_df["Stock Issued Amount"] = to_number(
        stock_df[issued_amount_col]
    )

    if received_qty_col:
        stock_df["Stock Received Qty"] = to_number(
            stock_df[received_qty_col]
        )
    else:
        stock_df["Stock Received Qty"] = 0.0

    if received_amount_col:
        stock_df["Stock Received Amount"] = to_number(
            stock_df[received_amount_col]
        )
    else:
        stock_df["Stock Received Amount"] = 0.0

    # --------------------------------------------------------
    # PARSE DATE COLUMN
    # --------------------------------------------------------

    if date_col:
        stock_df["Transaction Date Parsed"] = pd.to_datetime(
            stock_df[date_col],
            errors="coerce",
            dayfirst=True
        )

        stock_df["Transaction Month"] = (
            stock_df["Transaction Date Parsed"]
            .dt.to_period("M")
            .astype(str)
        )

        stock_df.loc[
            stock_df["Transaction Date Parsed"].isna(),
            "Transaction Month"
        ] = ""

    else:
        stock_df["Transaction Date Parsed"] = pd.NaT
        stock_df["Transaction Month"] = ""

    # --------------------------------------------------------
    # CREATE TRANSACTION TYPE
    # --------------------------------------------------------

    stock_df["Stock Transaction Type"] = "No Quantity Movement"

    stock_df.loc[
        (
            stock_df["Stock Received Qty"] != 0
        )
        & (
            stock_df["Stock Issued Qty"] == 0
        ),
        "Stock Transaction Type"
    ] = "Receipt"

    stock_df.loc[
        (
            stock_df["Stock Issued Qty"] != 0
        )
        & (
            stock_df["Stock Received Qty"] == 0
        ),
        "Stock Transaction Type"
    ] = "Issue"

    stock_df.loc[
        (
            stock_df["Stock Received Qty"] != 0
        )
        & (
            stock_df["Stock Issued Qty"] != 0
        ),
        "Stock Transaction Type"
    ] = "Receipt and Issue"

    stock_df.loc[
        (
            stock_df["Stock Received Amount"] != 0
        )
        & (
            stock_df["Stock Issued Amount"] == 0
        )
        & (
            stock_df["Stock Received Qty"] == 0
        ),
        "Stock Transaction Type"
    ] = "Receipt Value Only"

    stock_df.loc[
        (
            stock_df["Stock Issued Amount"] != 0
        )
        & (
            stock_df["Stock Received Amount"] == 0
        )
        & (
            stock_df["Stock Issued Qty"] == 0
        ),
        "Stock Transaction Type"
    ] = "Issue Value Only"

    # --------------------------------------------------------
    # CALCULATE ERP ISSUE AND RECEIPT RATES
    # --------------------------------------------------------

    stock_df["ERP Issue Rate"] = safe_divide(
        stock_df["Stock Issued Amount"],
        stock_df["Stock Issued Qty"]
    )

    stock_df["ERP Receipt Rate"] = safe_divide(
        stock_df["Stock Received Amount"],
        stock_df["Stock Received Qty"]
    )

    # --------------------------------------------------------
    # CREATE TRACE KEYS
    # --------------------------------------------------------

    stock_df["Stock Trace Key"] = (
        stock_df["PR_Clean"]
        + " || "
        + stock_df["PO_Clean"]
        + " || "
        + stock_df["GRN_Clean"]
        + " || "
        + stock_df["Item_Clean"]
    )

    stock_df["Inventory Key"] = (
        stock_df["Project_Clean"]
        + " || "
        + stock_df["Godown_Clean"]
        + " || "
        + stock_df["Item_Clean"]
    )

    stock_df["Consumption Key"] = (
        stock_df["Project_Clean"]
        + " || "
        + stock_df["Subproject_Clean"]
        + " || "
        + stock_df["Activity_Clean"]
        + " || "
        + stock_df["Contractor_Clean"]
        + " || "
        + stock_df["Item_Clean"]
    )

    # --------------------------------------------------------
    # STOCK DATA QUALITY CHECKS
    # --------------------------------------------------------

    stock_df["Stock Data Status"] = "OK"

    stock_df.loc[
        stock_df["Project_Clean"] == "",
        "Stock Data Status"
    ] = "REVIEW: Project is blank"

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
            stock_df["Stock Issued Amount"] != 0
        )
        & (
            stock_df["Subproject_Clean"] == ""
        ),
        "Stock Data Status"
    ] = "REVIEW: Issued amount has blank subproject"

    stock_df.loc[
        (
            stock_df["Stock Issued Qty"] != 0
        )
        & (
            stock_df["Stock Issued Amount"] == 0
        ),
        "Stock Data Status"
    ] = "REVIEW: Issued quantity exists but issued amount is zero"

    stock_df.loc[
        (
            stock_df["Stock Issued Qty"] == 0
        )
        & (
            stock_df["Stock Issued Amount"] != 0
        ),
        "Stock Data Status"
    ] = "REVIEW: Issued amount exists but issued quantity is zero"

    stock_df.loc[
        (
            stock_df["Stock Received Qty"] != 0
        )
        & (
            stock_df["Stock Received Amount"] == 0
        ),
        "Stock Data Status"
    ] = "REVIEW: Received quantity exists but received amount is zero"

    stock_df.loc[
        (
            stock_df["Stock Received Qty"] == 0
        )
        & (
            stock_df["Stock Received Amount"] != 0
        ),
        "Stock Data Status"
    ] = "REVIEW: Received amount exists but received quantity is zero"

    stock_df.loc[
        stock_df["Stock Issued Qty"] < 0,
        "Stock Data Status"
    ] = "REVIEW: Negative issued quantity"

    stock_df.loc[
        stock_df["Stock Received Qty"] < 0,
        "Stock Data Status"
    ] = "REVIEW: Negative received quantity"

    # --------------------------------------------------------
    # IDENTIFY ACTUAL ISSUE ROWS
    # --------------------------------------------------------

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

    # --------------------------------------------------------
    # PREPARE CONSUMPTION REGISTER
    # --------------------------------------------------------

    consumption_columns = []

    for column in [
        date_col,
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
        from_voucher_col,
        status_col,
        pr_col,
        po_col,
        grn_col,
        grn_line_col
    ]:
        if column and column in issue_df.columns:
            consumption_columns.append(column)

    consumption_columns.extend([
        "Transaction Date Parsed",
        "Transaction Month",
        "Actual Issue Subproject",
        "Actual Consumption Qty",
        "Actual Consumption Rate",
        "Actual Consumption Cost",
        "PR_Clean",
        "PO_Clean",
        "GRN_Clean",
        "Item_Clean",
        "Stock Trace Key",
        "Consumption Key",
        "Stock Data Status",
        "Source Stock Row No"
    ])

    consumption_columns = list(
        dict.fromkeys(consumption_columns)
    )

    consumption_register = issue_df[
        consumption_columns
    ].copy()

    # --------------------------------------------------------
    # SUBPROJECT CONSUMPTION SUMMARY
    # --------------------------------------------------------

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
            "Item_Clean": pd.Series.nunique,
            "Activity_Clean": pd.Series.nunique,
            "Contractor_Clean": pd.Series.nunique,
            "Voucher_Clean": pd.Series.nunique
        })
        .rename(columns={
            "Item_Clean": "Unique Item Count",
            "Activity_Clean": "Unique Activity Count",
            "Contractor_Clean": "Unique Contractor Count",
            "Voucher_Clean": "Issue Voucher Count"
        })
        .sort_values(
            "Actual Consumption Cost",
            ascending=False
        )
    )

    subproject_summary[
        "Average Consumption Rate"
    ] = safe_divide(
        subproject_summary[
            "Actual Consumption Cost"
        ],
        subproject_summary[
            "Actual Consumption Qty"
        ]
    )

    # --------------------------------------------------------
    # ACTIVITY SUMMARY
    # --------------------------------------------------------

    if activity_col:

        activity_group = [
            project_col,
            "Actual Issue Subproject",
            activity_col
        ]

        if company_col:
            activity_group.insert(
                0,
                company_col
            )

        activity_summary = (
            issue_df
            .groupby(
                activity_group,
                as_index=False,
                dropna=False
            )
            .agg({
                "Actual Consumption Qty": "sum",
                "Actual Consumption Cost": "sum",
                "Item_Clean": pd.Series.nunique,
                "Voucher_Clean": pd.Series.nunique
            })
            .rename(columns={
                "Item_Clean": "Unique Item Count",
                "Voucher_Clean": "Issue Voucher Count"
            })
            .sort_values(
                "Actual Consumption Cost",
                ascending=False
            )
        )

    else:

        activity_summary = pd.DataFrame(
            columns=[
                project_col,
                "Actual Issue Subproject",
                "Actual Consumption Qty",
                "Actual Consumption Cost",
                "Unique Item Count",
                "Issue Voucher Count"
            ]
        )

    # --------------------------------------------------------
    # CONTRACTOR SUMMARY
    # --------------------------------------------------------

    if contractor_col:

        contractor_group = [
            project_col,
            "Actual Issue Subproject",
            contractor_col
        ]

        if company_col:
            contractor_group.insert(
                0,
                company_col
            )

        contractor_summary = (
            issue_df
            .groupby(
                contractor_group,
                as_index=False,
                dropna=False
            )
            .agg({
                "Actual Consumption Qty": "sum",
                "Actual Consumption Cost": "sum",
                "Item_Clean": pd.Series.nunique,
                "Voucher_Clean": pd.Series.nunique
            })
            .rename(columns={
                "Item_Clean": "Unique Item Count",
                "Voucher_Clean": "Issue Voucher Count"
            })
            .sort_values(
                "Actual Consumption Cost",
                ascending=False
            )
        )

    else:

        contractor_summary = pd.DataFrame(
            columns=[
                project_col,
                "Actual Issue Subproject",
                "Actual Consumption Qty",
                "Actual Consumption Cost",
                "Unique Item Count",
                "Issue Voucher Count"
            ]
        )

    # --------------------------------------------------------
    # ITEM CONSUMPTION SUMMARY
    # --------------------------------------------------------

    item_group_columns = [
        project_col,
        "Actual Issue Subproject",
        item_col
    ]

    if company_col:
        item_group_columns.insert(
            0,
            company_col
        )

    if unit_col:
        item_group_columns.append(
            unit_col
        )

    item_consumption_summary = (
        issue_df
        .groupby(
            item_group_columns,
            as_index=False,
            dropna=False
        )
        .agg({
            "Actual Consumption Qty": "sum",
            "Actual Consumption Cost": "sum",
            "Voucher_Clean": pd.Series.nunique,
            "Activity_Clean": pd.Series.nunique,
            "Contractor_Clean": pd.Series.nunique
        })
        .rename(columns={
            "Voucher_Clean": "Issue Voucher Count",
            "Activity_Clean": "Unique Activity Count",
            "Contractor_Clean": "Unique Contractor Count"
        })
        .sort_values(
            "Actual Consumption Cost",
            ascending=False
        )
    )

    item_consumption_summary[
        "Weighted Average Issue Rate"
    ] = safe_divide(
        item_consumption_summary[
            "Actual Consumption Cost"
        ],
        item_consumption_summary[
            "Actual Consumption Qty"
        ]
    )

    # --------------------------------------------------------
    # INVENTORY SUMMARY
    # --------------------------------------------------------

    inventory_group_columns = [
        project_col,
        item_col
    ]

    if company_col:
        inventory_group_columns.insert(
            0,
            company_col
        )

    if godown_col:
        inventory_group_columns.insert(
            1 if not company_col else 2,
            godown_col
        )

    if item_group_col:
        inventory_group_columns.append(
            item_group_col
        )

    if unit_col:
        inventory_group_columns.append(
            unit_col
        )

    inventory_summary = (
        stock_df
        .groupby(
            inventory_group_columns,
            as_index=False,
            dropna=False
        )
        .agg({
            "Stock Received Qty": "sum",
            "Stock Issued Qty": "sum",
            "Stock Received Amount": "sum",
            "Stock Issued Amount": "sum",
            "PR_Clean": pd.Series.nunique,
            "PO_Clean": pd.Series.nunique,
            "GRN_Clean": pd.Series.nunique,
            "Voucher_Clean": pd.Series.nunique
        })
        .rename(columns={
            "PR_Clean": "Unique PR Count",
            "PO_Clean": "Unique PO Count",
            "GRN_Clean": "Unique GRN Count",
            "Voucher_Clean": "Issue Voucher Count"
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

    inventory_summary["Weighted Receipt Rate"] = safe_divide(
        inventory_summary["Stock Received Amount"],
        inventory_summary["Stock Received Qty"]
    )

    inventory_summary["Weighted Issue Rate"] = safe_divide(
        inventory_summary["Stock Issued Amount"],
        inventory_summary["Stock Issued Qty"]
    )

    inventory_summary["Calculated Closing Rate"] = safe_divide(
        inventory_summary["Closing Stock Value"],
        inventory_summary["Closing Stock Qty"]
    )

    inventory_summary["Inventory Status"] = "OK"

    inventory_summary.loc[
        (
            inventory_summary["Closing Stock Qty"]
            < -abs(inventory_qty_tolerance)
        ),
        "Inventory Status"
    ] = "REVIEW: Negative closing stock quantity"

    inventory_summary.loc[
        (
            inventory_summary["Closing Stock Value"]
            < -abs(inventory_value_tolerance)
        ),
        "Inventory Status"
    ] = "REVIEW: Negative closing stock value"

    inventory_summary.loc[
        (
            inventory_summary["Closing Stock Qty"] > abs(
                inventory_qty_tolerance
            )
        )
        & (
            inventory_summary["Closing Stock Value"] == 0
        ),
        "Inventory Status"
    ] = "REVIEW: Positive stock quantity has zero value"

    inventory_summary.loc[
        (
            inventory_summary["Closing Stock Qty"] == 0
        )
        & (
            inventory_summary["Closing Stock Value"]
            > abs(inventory_value_tolerance)
        ),
        "Inventory Status"
    ] = "REVIEW: Zero stock quantity has remaining value"

    inventory_summary = inventory_summary.sort_values(
        "Closing Stock Value",
        ascending=False
    )

    # --------------------------------------------------------
    # MONTHLY CONSUMPTION SUMMARY
    # --------------------------------------------------------

    if date_col:

        monthly_consumption_summary = (
            issue_df[
                issue_df["Transaction Month"] != ""
            ]
            .groupby(
                [
                    "Transaction Month",
                    project_col
                ],
                as_index=False,
                dropna=False
            )
            .agg({
                "Actual Consumption Qty": "sum",
                "Actual Consumption Cost": "sum",
                "Voucher_Clean": pd.Series.nunique,
                "Item_Clean": pd.Series.nunique
            })
            .rename(columns={
                "Voucher_Clean": "Issue Voucher Count",
                "Item_Clean": "Unique Item Count"
            })
            .sort_values(
                "Transaction Month"
            )
        )

    else:

        monthly_consumption_summary = pd.DataFrame(
            columns=[
                "Transaction Month",
                project_col,
                "Actual Consumption Qty",
                "Actual Consumption Cost",
                "Issue Voucher Count",
                "Unique Item Count"
            ]
        )

    # --------------------------------------------------------
    # ISSUE REVIEW REPORT
    # --------------------------------------------------------

    issue_review = issue_df[
        issue_df["Stock Data Status"] != "OK"
    ].copy()

    # --------------------------------------------------------
    # INVENTORY REVIEW REPORT
    # --------------------------------------------------------

    inventory_review = inventory_summary[
        inventory_summary["Inventory Status"] != "OK"
    ].copy()

    # --------------------------------------------------------
    # DETECTED COLUMN INFORMATION
    # --------------------------------------------------------

    detected_columns = {
        "company": company_col,
        "project": project_col,
        "subproject": subproject_col,
        "activity": activity_col,
        "contractor": contractor_col,
        "godown": godown_col,
        "item_group": item_group_col,
        "item": item_col,
        "unit": unit_col,
        "date": date_col,
        "voucher": voucher_col,
        "from_voucher": from_voucher_col,
        "status": status_col,
        "pr": pr_col,
        "po": po_col,
        "grn": grn_col,
        "grn_line": grn_line_col,
        "received_qty": received_qty_col,
        "received_amount": received_amount_col,
        "issued_qty": issued_qty_col,
        "issued_amount": issued_amount_col
    }

    return (
        stock_df,
        consumption_register,
        subproject_summary,
        activity_summary,
        contractor_summary,
        item_consumption_summary,
        inventory_summary,
        monthly_consumption_summary,
        issue_review,
        inventory_review,
        detected_columns
    )


# ============================================================
# RUN MODULE 3
# ============================================================

try:

    (
        stock_register,
        consumption_register,
        subproject_consumption_summary,
        activity_consumption_summary,
        contractor_consumption_summary,
        item_consumption_summary,
        inventory_summary,
        monthly_consumption_summary,
        stock_issue_review,
        inventory_review,
        stock_columns
    ) = process_stock_ledger(
        stock_raw_df,
        inventory_value_tolerance=inventory_value_tolerance,
        inventory_qty_tolerance=inventory_qty_tolerance
    )

    st.divider()
    st.header(
        "Module 3 — Actual Consumption and Inventory"
    )

    # --------------------------------------------------------
    # CALCULATE MAIN METRICS
    # --------------------------------------------------------

    total_received_qty = (
        stock_register[
            "Stock Received Qty"
        ].sum()
    )

    total_issued_qty = (
        stock_register[
            "Stock Issued Qty"
        ].sum()
    )

    total_received_value = (
        stock_register[
            "Stock Received Amount"
        ].sum()
    )

    total_consumption_cost = (
        consumption_register[
            "Actual Consumption Cost"
        ].sum()
    )

    total_closing_stock_qty = (
        inventory_summary[
            "Closing Stock Qty"
        ].sum()
    )

    total_closing_stock_value = (
        inventory_summary[
            "Closing Stock Value"
        ].sum()
    )

    unique_subprojects = (
        consumption_register[
            "Actual Issue Subproject"
        ]
        .replace("", pd.NA)
        .dropna()
        .nunique()
    )

    unique_items_consumed = (
        consumption_register[
            "Item_Clean"
        ]
        .replace("", pd.NA)
        .dropna()
        .nunique()
    )

    # --------------------------------------------------------
    # DISPLAY PRIMARY METRICS
    # --------------------------------------------------------

    consumption_c1, consumption_c2, consumption_c3, consumption_c4 = (
        st.columns(4)
    )

    show_metric(
        consumption_c1,
        "Actual Consumption Cost",
        total_consumption_cost,
        "money"
    )

    show_metric(
        consumption_c2,
        "Stock Received Value",
        total_received_value,
        "money"
    )

    show_metric(
        consumption_c3,
        "Closing Stock Value",
        total_closing_stock_value,
        "money"
    )

    show_metric(
        consumption_c4,
        "Actual Issue Rows",
        len(consumption_register),
        "integer"
    )

    quantity_c1, quantity_c2, quantity_c3, quantity_c4 = (
        st.columns(4)
    )

    show_metric(
        quantity_c1,
        "Total Received Quantity",
        total_received_qty,
        "quantity"
    )

    show_metric(
        quantity_c2,
        "Total Issued Quantity",
        total_issued_qty,
        "quantity"
    )

    show_metric(
        quantity_c3,
        "Closing Stock Quantity",
        total_closing_stock_qty,
        "quantity"
    )

    show_metric(
        quantity_c4,
        "Unique Consuming Subprojects",
        unique_subprojects,
        "integer"
    )

    audit_c1, audit_c2, audit_c3, audit_c4 = (
        st.columns(4)
    )

    show_metric(
        audit_c1,
        "Unique Consumed Items",
        unique_items_consumed,
        "integer"
    )

    show_metric(
        audit_c2,
        "Issue Review Rows",
        len(stock_issue_review),
        "integer"
    )

    show_metric(
        audit_c3,
        "Inventory Review Rows",
        len(inventory_review),
        "integer"
    )

    show_metric(
        audit_c4,
        "Inventory Summary Rows",
        len(inventory_summary),
        "integer"
    )

    # --------------------------------------------------------
    # SHOW SUMMARIES
    # --------------------------------------------------------

    show_dataframe(
        "Subproject Consumption Summary",
        subproject_consumption_summary,
        maximum_rows=200
    )

    show_dataframe(
        "Activity-wise Consumption Summary",
        activity_consumption_summary,
        maximum_rows=200
    )

    show_dataframe(
        "Contractor-wise Consumption Summary",
        contractor_consumption_summary,
        maximum_rows=200
    )

    show_dataframe(
        "Item-wise Consumption Summary",
        item_consumption_summary,
        maximum_rows=200
    )

    show_dataframe(
        "Inventory Summary",
        inventory_summary,
        maximum_rows=200
    )

    if not monthly_consumption_summary.empty:

        show_dataframe(
            "Monthly Consumption Summary",
            monthly_consumption_summary,
            maximum_rows=200
        )

    # --------------------------------------------------------
    # DETAILED REGISTER PREVIEWS
    # --------------------------------------------------------

    with st.expander(
        "Preview Actual Consumption Register",
        expanded=False
    ):

        show_dataframe(
            "Consumption Register",
            consumption_register,
            maximum_rows=200
        )

    with st.expander(
        "Preview Full Stock Ledger Register",
        expanded=False
    ):

        show_dataframe(
            "Stock Ledger Register",
            stock_register,
            maximum_rows=200
        )

    # --------------------------------------------------------
    # REVIEW SECTIONS
    # --------------------------------------------------------

    if not stock_issue_review.empty:

        with st.expander(
            "Stock Issue Review Required",
            expanded=False
        ):

            show_dataframe(
                "Stock Issue Review",
                stock_issue_review,
                maximum_rows=200
            )

    if not inventory_review.empty:

        with st.expander(
            "Inventory Review Required",
            expanded=False
        ):

            show_dataframe(
                "Inventory Review",
                inventory_review,
                maximum_rows=200
            )

    # --------------------------------------------------------
    # STOCK MOVEMENT RECONCILIATION
    # --------------------------------------------------------

    stock_reconciliation_difference = (
        total_received_value
        - total_consumption_cost
        - total_closing_stock_value
    )

    stock_reconciliation_status = (
        "Reconciled"
        if abs(
            stock_reconciliation_difference
        ) <= 0.01
        else "REVIEW REQUIRED"
    )

    stock_reconciliation = pd.DataFrame([
        {
            "Check":
                "Stock Ledger movement reconciliation",

            "Received Value":
                total_received_value,

            "Issued / Consumed Value":
                total_consumption_cost,

            "Closing Stock Value":
                total_closing_stock_value,

            "Difference":
                stock_reconciliation_difference,

            "Status":
                stock_reconciliation_status
        }
    ])

    show_dataframe(
        "Stock Ledger Reconciliation",
        stock_reconciliation
    )

    st.success(
        "Module 3 completed successfully. "
        "Actual subproject consumption and inventory reports are ready."
    )

except Exception as error:

    st.error(
        "Module 3 could not process the Stock Ledger report."
    )

    st.exception(error)

    st.stop()
    # ============================================================
# MODULE 4 — PR REFERENCE AND VARIANCE ENGINE
# ============================================================


def process_pr_report(
    pr_raw_df,
    consumption_register
):
    """
    Process the StrategicERP Purchase Requisition report.

    Business logic:
    - PR is used only as the intended-use reference.
    - Actual consumption remains based on Stock Ledger Issued Amount.
    - Actual consuming subproject remains the Stock Ledger issue subproject.
    - A difference between PR subproject and issue subproject is informational.
    - PR differences do not change actual consumption cost.

    Returns:
    1. Detailed PR register
    2. PR reference summary
    3. Consumption register with PR validation
    4. PR versus actual issue variance report
    5. Issues with no PR reference
    6. PR data review
    7. Detected PR column dictionary
    """

    pr_df = pr_raw_df.copy()

    # --------------------------------------------------------
    # FIND REQUIRED PR COLUMNS
    # --------------------------------------------------------

    pr_number_col = find_column(
        pr_df,
        [
            "Purchase Requisition (PR) No.",
            "PRNo",
            "PR No",
            "P.RNo",
            "P.R. No",
            "Purchase Requisition No"
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
            "Requested Qty",
            "Requisition Qty"
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

    # --------------------------------------------------------
    # FIND OPTIONAL PR COLUMNS
    # --------------------------------------------------------

    company_col = find_column(
        pr_df,
        [
            "Name of Company",
            "Name Of Company",
            "Company Name",
            "Company"
        ],
        required=False
    )

    project_col = find_column(
        pr_df,
        [
            "Project Name",
            "Project"
        ],
        required=False
    )

    unit_col = find_column(
        pr_df,
        [
            "Unit",
            "UOM",
            "U.O.M."
        ],
        required=False
    )

    pr_date_col = find_column(
        pr_df,
        [
            "PR Date",
            "Purchase Requisition Date",
            "Requisition Date",
            "Date"
        ],
        required=False
    )

    required_by_date_col = find_column(
        pr_df,
        [
            "Required By Date",
            "Required Date",
            "Need By Date"
        ],
        required=False
    )

    department_col = find_column(
        pr_df,
        [
            "Department Name",
            "Department",
            "Dept Name"
        ],
        required=False
    )

    activity_col = find_column(
        pr_df,
        [
            "Activity Code",
            "Activity",
            "Work Activity"
        ],
        required=False
    )

    requestor_col = find_column(
        pr_df,
        [
            "Requested By",
            "Created By",
            "Prepared By",
            "Requestor",
            "Requisition By"
        ],
        required=False
    )

    status_col = find_column(
        pr_df,
        [
            "Status",
            "PR Status",
            "Requisition Status"
        ],
        required=False
    )

    remarks_col = find_column(
        pr_df,
        [
            "Remarks",
            "Remark",
            "Description",
            "Narration"
        ],
        required=False
    )

    # --------------------------------------------------------
    # ADD SOURCE ROW NUMBER
    # --------------------------------------------------------

    pr_df["Source PR Row No"] = range(
        1,
        len(pr_df) + 1
    )

    # --------------------------------------------------------
    # CLEAN IMPORTANT FIELDS
    # --------------------------------------------------------

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

    pr_df["PR Subproject Original"] = (
        pr_df[subproject_col]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    if project_col:
        pr_df["PR Project Clean"] = (
            pr_df[project_col]
            .apply(clean_text)
        )
    else:
        pr_df["PR Project Clean"] = ""

    if activity_col:
        pr_df["PR Activity Clean"] = (
            pr_df[activity_col]
            .apply(clean_text)
        )
    else:
        pr_df["PR Activity Clean"] = ""

    # --------------------------------------------------------
    # DATE CLEANING
    # --------------------------------------------------------

    if pr_date_col:
        pr_df["PR Date Parsed"] = pd.to_datetime(
            pr_df[pr_date_col],
            errors="coerce",
            dayfirst=True
        )
    else:
        pr_df["PR Date Parsed"] = pd.NaT

    if required_by_date_col:
        pr_df["Required By Date Parsed"] = pd.to_datetime(
            pr_df[required_by_date_col],
            errors="coerce",
            dayfirst=True
        )
    else:
        pr_df["Required By Date Parsed"] = pd.NaT

    # --------------------------------------------------------
    # PR DATA QUALITY CHECKS
    # --------------------------------------------------------

    pr_df["PR Data Status"] = "OK"

    pr_df.loc[
        pr_df["PR_Clean"] == "",
        "PR Data Status"
    ] = "REVIEW: PR number is blank"

    pr_df.loc[
        pr_df["Item_Clean"] == "",
        "PR Data Status"
    ] = "REVIEW: Item description is blank"

    pr_df.loc[
        pr_df["PR Quantity Numeric"] <= 0,
        "PR Data Status"
    ] = "REVIEW: PR quantity is zero or blank"

    pr_df.loc[
        pr_df["PR Subproject Original"] == "",
        "PR Data Status"
    ] = "REVIEW: PR subproject is blank"

    # --------------------------------------------------------
    # EXPAND MULTIPLE PR SUBPROJECTS
    # --------------------------------------------------------

    expanded_rows = []

    for _, row in pr_df.iterrows():

        pr_number = row["PR_Clean"]
        item_clean = row["Item_Clean"]
        quantity = row["PR Quantity Numeric"]

        subprojects = split_subprojects(
            row[subproject_col]
        )

        if not subprojects:
            subprojects = [""]

        # The quantity is kept as the original PR quantity.
        # It is not divided and is not used for consumption costing.
        for intended_subproject in subprojects:

            expanded_row = {
                "PR_Clean": pr_number,
                "Item_Clean": item_clean,
                "Intended PR Subproject": intended_subproject,
                "PR Reference Quantity": quantity,
                "Source PR Row No": row["Source PR Row No"],
                "PR Data Status": row["PR Data Status"]
            }

            if company_col:
                expanded_row["PR Company"] = row[company_col]

            if project_col:
                expanded_row["PR Project"] = row[project_col]

            if unit_col:
                expanded_row["PR Unit"] = row[unit_col]

            if department_col:
                expanded_row["PR Department"] = row[department_col]

            if activity_col:
                expanded_row["PR Activity Code"] = row[activity_col]

            if requestor_col:
                expanded_row["PR Requested By"] = row[requestor_col]

            if status_col:
                expanded_row["PR ERP Status"] = row[status_col]

            if remarks_col:
                expanded_row["PR Remarks"] = row[remarks_col]

            if pr_date_col:
                expanded_row["PR Date"] = row[pr_date_col]

            if required_by_date_col:
                expanded_row["Required By Date"] = row[
                    required_by_date_col
                ]

            expanded_rows.append(
                expanded_row
            )

    pr_reference_detail = pd.DataFrame(
        expanded_rows
    )

    # --------------------------------------------------------
    # CREATE PR REFERENCE SUMMARY
    # --------------------------------------------------------

    if pr_reference_detail.empty:

        pr_reference_summary = pd.DataFrame(
            columns=[
                "PR_Clean",
                "Item_Clean",
                "Intended PR Subprojects",
                "PR Reference Quantity",
                "Source PR Row Nos",
                "PR Reference Status"
            ]
        )

    else:

        aggregation_rules = {
            "Intended PR Subproject": join_unique,
            "PR Reference Quantity": "max",
            "Source PR Row No": lambda values: join_unique(
                values.astype(str)
            ),
            "PR Data Status": join_unique
        }

        optional_reference_columns = [
            "PR Company",
            "PR Project",
            "PR Unit",
            "PR Department",
            "PR Activity Code",
            "PR Requested By",
            "PR ERP Status"
        ]

        for optional_column in optional_reference_columns:
            if optional_column in pr_reference_detail.columns:
                aggregation_rules[optional_column] = join_unique

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
            .agg(aggregation_rules)
            .rename(columns={
                "Intended PR Subproject":
                    "Intended PR Subprojects",

                "Source PR Row No":
                    "Source PR Row Nos",

                "PR Data Status":
                    "PR Reference Status"
            })
        )

    # --------------------------------------------------------
    # ADD PR REFERENCE TO ACTUAL CONSUMPTION REGISTER
    # --------------------------------------------------------

    consumption_with_pr = (
        consumption_register
        .copy()
    )

    consumption_with_pr = consumption_with_pr.merge(
        pr_reference_summary,
        on=[
            "PR_Clean",
            "Item_Clean"
        ],
        how="left",
        indicator=True
    )

    consumption_with_pr[
        "Intended PR Subprojects"
    ] = (
        consumption_with_pr[
            "Intended PR Subprojects"
        ]
        .fillna("")
    )

    consumption_with_pr[
        "PR Reference Quantity"
    ] = to_number(
        consumption_with_pr[
            "PR Reference Quantity"
        ]
    )

    consumption_with_pr[
        "PR Reference Match Status"
    ] = "PR + Item reference found"

    consumption_with_pr.loc[
        consumption_with_pr["_merge"] != "both",
        "PR Reference Match Status"
    ] = "INFORMATIONAL: PR + Item reference not found"

    consumption_with_pr = (
        consumption_with_pr
        .drop(columns=["_merge"])
    )

    # --------------------------------------------------------
    # COMPARE ACTUAL ISSUE SUBPROJECT WITH PR REFERENCE
    # --------------------------------------------------------

    def compare_actual_with_pr(row):
        """
        Compare the actual issue subproject with the PR reference.

        This is informational only.
        It does not alter or reallocate consumption cost.
        """

        actual_subproject = clean_subproject(
            row.get(
                "Actual Issue Subproject",
                ""
            )
        )

        intended_text = str(
            row.get(
                "Intended PR Subprojects",
                ""
            )
        ).strip()

        if not intended_text:
            return (
                "INFORMATIONAL: "
                "PR reference unavailable"
            )

        intended_subprojects = [
            clean_subproject(value)
            for value in intended_text.split(" | ")
            if clean_subproject(value)
        ]

        if actual_subproject in intended_subprojects:
            return (
                "MATCHED: Actual issue agrees "
                "with PR reference"
            )

        return (
            "INFORMATIONAL: Actual issue "
            "subproject differs from PR reference"
        )

    if consumption_with_pr.empty:

        consumption_with_pr[
            "PR vs Actual Issue Result"
        ] = pd.Series(dtype=str)

    else:

        consumption_with_pr[
            "PR vs Actual Issue Result"
        ] = consumption_with_pr.apply(
            compare_actual_with_pr,
            axis=1
        )

    # --------------------------------------------------------
    # PR VARIANCE REPORT
    # --------------------------------------------------------

    pr_variance = consumption_with_pr[
        consumption_with_pr[
            "PR vs Actual Issue Result"
        ].str.contains(
            "differs",
            case=False,
            na=False
        )
    ].copy()

    # PR variance is intentionally not classified as an error.
    pr_variance[
        "Variance Classification"
    ] = (
        "INFORMATIONAL: Actual consuming subproject "
        "differs from intended PR subproject"
    )

    # --------------------------------------------------------
    # ISSUES WITH NO PR REFERENCE
    # --------------------------------------------------------

    no_pr_reference = consumption_with_pr[
        consumption_with_pr[
            "PR Reference Match Status"
        ].str.contains(
            "not found",
            case=False,
            na=False
        )
    ].copy()

    # --------------------------------------------------------
    # ACTUAL CONSUMPTION SUMMARY BY PR VALIDATION RESULT
    # --------------------------------------------------------

    pr_validation_summary = (
        consumption_with_pr
        .groupby(
            "PR vs Actual Issue Result",
            as_index=False,
            dropna=False
        )
        .agg({
            "Actual Consumption Qty": "sum",
            "Actual Consumption Cost": "sum",
            "Source Stock Row No": "count"
        })
        .rename(columns={
            "Source Stock Row No":
                "Consumption Row Count"
        })
        .sort_values(
            "Actual Consumption Cost",
            ascending=False
        )
    )

    # --------------------------------------------------------
    # PR VARIANCE SUMMARY BY ACTUAL SUBPROJECT
    # --------------------------------------------------------

    if pr_variance.empty:

        pr_variance_summary = pd.DataFrame(
            columns=[
                "Actual Issue Subproject",
                "Actual Consumption Qty",
                "Actual Consumption Cost",
                "Variance Row Count"
            ]
        )

    else:

        pr_variance_summary = (
            pr_variance
            .groupby(
                "Actual Issue Subproject",
                as_index=False,
                dropna=False
            )
            .agg({
                "Actual Consumption Qty": "sum",
                "Actual Consumption Cost": "sum",
                "Source Stock Row No": "count"
            })
            .rename(columns={
                "Source Stock Row No":
                    "Variance Row Count"
            })
            .sort_values(
                "Actual Consumption Cost",
                ascending=False
            )
        )

    # --------------------------------------------------------
    # PR DATA REVIEW REPORT
    # --------------------------------------------------------

    pr_review = pr_df[
        pr_df["PR Data Status"] != "OK"
    ].copy()

    # --------------------------------------------------------
    # DETECTED PR COLUMN INFORMATION
    # --------------------------------------------------------

    detected_columns = {
        "company": company_col,
        "project": project_col,
        "pr_number": pr_number_col,
        "pr_date": pr_date_col,
        "required_by_date": required_by_date_col,
        "item": item_col,
        "unit": unit_col,
        "quantity": quantity_col,
        "subproject": subproject_col,
        "department": department_col,
        "activity": activity_col,
        "requestor": requestor_col,
        "status": status_col,
        "remarks": remarks_col
    }

    return (
        pr_df,
        pr_reference_detail,
        pr_reference_summary,
        consumption_with_pr,
        pr_validation_summary,
        pr_variance,
        pr_variance_summary,
        no_pr_reference,
        pr_review,
        detected_columns
    )


# ============================================================
# RUN MODULE 4
# ============================================================

try:

    (
        pr_register,
        pr_reference_detail,
        pr_reference_summary,
        consumption_register_with_pr,
        pr_validation_summary,
        pr_variance,
        pr_variance_summary,
        no_pr_reference,
        pr_review,
        pr_columns
    ) = process_pr_report(
        pr_raw_df,
        consumption_register
    )

    st.divider()

    st.header(
        "Module 4 — PR Reference and Variance"
    )

    # --------------------------------------------------------
    # MODULE 4 METRICS
    # --------------------------------------------------------

    matched_pr_rows = int(
        consumption_register_with_pr[
            "PR vs Actual Issue Result"
        ]
        .str.contains(
            "MATCHED",
            case=False,
            na=False
        )
        .sum()
    )

    variance_rows = len(
        pr_variance
    )

    no_reference_rows = len(
        no_pr_reference
    )

    unique_pr_numbers = (
        pr_register[
            "PR_Clean"
        ]
        .replace("", pd.NA)
        .dropna()
        .nunique()
    )

    pr_c1, pr_c2, pr_c3, pr_c4 = st.columns(4)

    show_metric(
        pr_c1,
        "Unique PR Numbers",
        unique_pr_numbers,
        "integer"
    )

    show_metric(
        pr_c2,
        "Matched PR Reference Rows",
        matched_pr_rows,
        "integer"
    )

    show_metric(
        pr_c3,
        "PR Variance Rows",
        variance_rows,
        "integer"
    )

    show_metric(
        pr_c4,
        "No PR Reference Rows",
        no_reference_rows,
        "integer"
    )

    # --------------------------------------------------------
    # DISPLAY PR SUMMARIES
    # --------------------------------------------------------

    show_dataframe(
        "PR Validation Summary",
        pr_validation_summary,
        maximum_rows=100
    )

    show_dataframe(
        "PR Variance Summary by Actual Subproject",
        pr_variance_summary,
        maximum_rows=200
    )

    # --------------------------------------------------------
    # DISPLAY REFERENCE TABLES
    # --------------------------------------------------------

    with st.expander(
        "Preview PR Reference Summary",
        expanded=False
    ):

        show_dataframe(
            "PR Reference Summary",
            pr_reference_summary,
            maximum_rows=200
        )

    with st.expander(
        "Preview Consumption Register with PR Reference",
        expanded=False
    ):

        show_dataframe(
            "Consumption Register with PR Reference",
            consumption_register_with_pr,
            maximum_rows=200
        )

    # --------------------------------------------------------
    # DISPLAY INFORMATIONAL VARIANCES
    # --------------------------------------------------------

    if not pr_variance.empty:

        with st.expander(
            "PR versus Actual Issue Variance — Informational",
            expanded=False
        ):

            st.info(
                "These rows are not costing errors. "
                "The Stock Ledger issue subproject remains "
                "the actual consuming subproject."
            )

            show_dataframe(
                "PR Variance Detail",
                pr_variance,
                maximum_rows=300
            )

    if not no_pr_reference.empty:

        with st.expander(
            "Issues Without PR + Item Reference",
            expanded=False
        ):

            st.info(
                "Consumption remains valid because it is based "
                "on the Stock Ledger issued amount. "
                "This report only identifies missing PR traceability."
            )

            show_dataframe(
                "No PR Reference",
                no_pr_reference,
                maximum_rows=300
            )

    # --------------------------------------------------------
    # DISPLAY GENUINE PR DATA ISSUES
    # --------------------------------------------------------

    if not pr_review.empty:

        with st.expander(
            "PR Source Data Review Required",
            expanded=False
        ):

            show_dataframe(
                "PR Review",
                pr_review,
                maximum_rows=200
            )

    st.success(
        "Module 4 completed successfully. "
        "PR references and informational variances are ready."
    )

except Exception as error:

    st.error(
        "Module 4 could not process the PR report."
    )

    st.exception(error)

    st.stop()
    # ============================================================
# MODULE 5 — FINAL DASHBOARD, AUDIT AND EXCEL EXPORT
# ============================================================


def build_final_audit_report(
    purchase_review,
    stock_issue_review,
    inventory_review,
    pr_review
):
    """
    Combine only genuine data-quality issues.

    PR variance is intentionally excluded because it is informational.
    """

    audit_frames = []

    if purchase_review is not None and not purchase_review.empty:
        purchase_audit = purchase_review.copy()
        purchase_audit.insert(
            0,
            "Audit Source",
            "Purchase Bill"
        )
        audit_frames.append(
            purchase_audit
        )

    if stock_issue_review is not None and not stock_issue_review.empty:
        stock_audit = stock_issue_review.copy()
        stock_audit.insert(
            0,
            "Audit Source",
            "Stock Ledger Issue"
        )
        audit_frames.append(
            stock_audit
        )

    if inventory_review is not None and not inventory_review.empty:
        inventory_audit = inventory_review.copy()
        inventory_audit.insert(
            0,
            "Audit Source",
            "Inventory"
        )
        audit_frames.append(
            inventory_audit
        )

    if pr_review is not None and not pr_review.empty:
        pr_audit = pr_review.copy()
        pr_audit.insert(
            0,
            "Audit Source",
            "PR Source Data"
        )
        audit_frames.append(
            pr_audit
        )

    if not audit_frames:
        return pd.DataFrame(
            columns=[
                "Audit Source",
                "Audit Status"
            ]
        )

    all_columns = []

    for dataframe in audit_frames:
        for column in dataframe.columns:
            if column not in all_columns:
                all_columns.append(column)

    aligned_frames = []

    for dataframe in audit_frames:
        aligned_dataframe = dataframe.reindex(
            columns=all_columns
        )
        aligned_frames.append(
            aligned_dataframe
        )

    final_audit = pd.concat(
        aligned_frames,
        ignore_index=True
    )

    return final_audit


def build_dashboard_summary(
    purchase_register,
    consumption_register_with_pr,
    inventory_summary,
    procurement_summary,
    subproject_consumption_summary,
    activity_consumption_summary,
    contractor_consumption_summary,
    item_consumption_summary,
    purchase_review,
    stock_issue_review,
    inventory_review,
    pr_review,
    pr_variance
):
    """
    Create a one-row dashboard summary for Excel export.
    """

    procurement_principal = purchase_register[
        "PB Principal Amount"
    ].sum()

    procurement_gst = purchase_register[
        "PB GST Amount"
    ].sum()

    procurement_total = purchase_register[
        "PB Total Bill Amount"
    ].sum()

    actual_consumption = consumption_register_with_pr[
        "Actual Consumption Cost"
    ].sum()

    closing_stock_value = inventory_summary[
        "Closing Stock Value"
    ].sum()

    closing_stock_qty = inventory_summary[
        "Closing Stock Qty"
    ].sum()

    received_stock_value = inventory_summary[
        "Stock Received Amount"
    ].sum()

    issued_stock_value = inventory_summary[
        "Stock Issued Amount"
    ].sum()

    return pd.DataFrame([
        {
            "Report Generated On":
                datetime.now().strftime(
                    "%d-%m-%Y %H:%M"
                ),

            "Procurement Principal Cost":
                procurement_principal,

            "Procurement GST":
                procurement_gst,

            "Total Purchase Bill Value":
                procurement_total,

            "Stock Ledger Received Value":
                received_stock_value,

            "Actual Consumption Cost":
                actual_consumption,

            "Stock Ledger Issued Value":
                issued_stock_value,

            "Closing Stock Value":
                closing_stock_value,

            "Closing Stock Quantity":
                closing_stock_qty,

            "Purchase Bill Rows":
                len(purchase_register),

            "Consumption Rows":
                len(consumption_register_with_pr),

            "Inventory Summary Rows":
                len(inventory_summary),

            "Procurement Summary Rows":
                len(procurement_summary),

            "Subproject Summary Rows":
                len(subproject_consumption_summary),

            "Activity Summary Rows":
                len(activity_consumption_summary),

            "Contractor Summary Rows":
                len(contractor_consumption_summary),

            "Item Summary Rows":
                len(item_consumption_summary),

            "Purchase Review Rows":
                len(purchase_review),

            "Stock Issue Review Rows":
                len(stock_issue_review),

            "Inventory Review Rows":
                len(inventory_review),

            "PR Source Review Rows":
                len(pr_review),

            "PR Variance Rows":
                len(pr_variance)
        }
    ])


def build_data_quality_summary(
    purchase_register,
    consumption_register_with_pr,
    inventory_summary,
    pr_register,
    purchase_review,
    stock_issue_review,
    inventory_review,
    pr_review,
    pr_variance,
    no_pr_reference
):
    """
    Create a clean data-quality summary.

    PR variance and missing PR reference are informational.
    """

    return pd.DataFrame([
        {
            "Check":
                "Purchase Bill source data",

            "Total Rows":
                len(purchase_register),

            "Review Rows":
                len(purchase_review),

            "Classification":
                "ERROR CHECK",

            "Status":
                (
                    "OK"
                    if len(purchase_review) == 0
                    else "REVIEW REQUIRED"
                )
        },
        {
            "Check":
                "Stock issue source data",

            "Total Rows":
                len(consumption_register_with_pr),

            "Review Rows":
                len(stock_issue_review),

            "Classification":
                "ERROR CHECK",

            "Status":
                (
                    "OK"
                    if len(stock_issue_review) == 0
                    else "REVIEW REQUIRED"
                )
        },
        {
            "Check":
                "Inventory balances",

            "Total Rows":
                len(inventory_summary),

            "Review Rows":
                len(inventory_review),

            "Classification":
                "ERROR CHECK",

            "Status":
                (
                    "OK"
                    if len(inventory_review) == 0
                    else "REVIEW REQUIRED"
                )
        },
        {
            "Check":
                "PR source data",

            "Total Rows":
                len(pr_register),

            "Review Rows":
                len(pr_review),

            "Classification":
                "ERROR CHECK",

            "Status":
                (
                    "OK"
                    if len(pr_review) == 0
                    else "REVIEW REQUIRED"
                )
        },
        {
            "Check":
                "PR vs actual issue variance",

            "Total Rows":
                len(consumption_register_with_pr),

            "Review Rows":
                len(pr_variance),

            "Classification":
                "INFORMATIONAL",

            "Status":
                "INFORMATIONAL"
        },
        {
            "Check":
                "Consumption rows without PR reference",

            "Total Rows":
                len(consumption_register_with_pr),

            "Review Rows":
                len(no_pr_reference),

            "Classification":
                "INFORMATIONAL",

            "Status":
                "INFORMATIONAL"
        }
    ])


def build_procurement_consumption_comparison(
    purchase_register,
    consumption_register_with_pr
):
    """
    Compare procurement and consumption at project level.

    This is informational because Purchase Bill and Stock Ledger
    may not cover the same date range or valuation basis.
    """

    purchase_project_col = purchase_columns[
        "project"
    ]

    stock_project_col = stock_columns[
        "project"
    ]

    procurement_by_project = (
        purchase_register
        .groupby(
            purchase_project_col,
            as_index=False,
            dropna=False
        )
        .agg({
            "PB Principal Amount": "sum",
            "PB GST Amount": "sum",
            "PB Total Bill Amount": "sum"
        })
        .rename(columns={
            purchase_project_col:
                "Project",

            "PB Principal Amount":
                "Procurement Principal Cost",

            "PB GST Amount":
                "Procurement GST",

            "PB Total Bill Amount":
                "Procurement Total Including GST"
        })
    )

    consumption_by_project = (
        consumption_register_with_pr
        .groupby(
            stock_project_col,
            as_index=False,
            dropna=False
        )
        .agg({
            "Actual Consumption Cost": "sum"
        })
        .rename(columns={
            stock_project_col:
                "Project",

            "Actual Consumption Cost":
                "Actual Consumption Cost"
        })
    )

    comparison = procurement_by_project.merge(
        consumption_by_project,
        on="Project",
        how="outer"
    )

    numeric_columns = [
        "Procurement Principal Cost",
        "Procurement GST",
        "Procurement Total Including GST",
        "Actual Consumption Cost"
    ]

    for column in numeric_columns:
        comparison[column] = to_number(
            comparison[column]
        )

    comparison[
        "Procurement Principal Less Consumption"
    ] = (
        comparison[
            "Procurement Principal Cost"
        ]
        - comparison[
            "Actual Consumption Cost"
        ]
    )

    comparison[
        "Comparison Status"
    ] = (
        "INFORMATIONAL: Compare only when "
        "both reports cover the same period and scope"
    )

    return comparison.sort_values(
        "Actual Consumption Cost",
        ascending=False
    )


try:

    st.divider()

    st.header(
        "Module 5 — Final Dashboard and Export"
    )

    # --------------------------------------------------------
    # BUILD FINAL REPORTS
    # --------------------------------------------------------

    final_audit_report = build_final_audit_report(
        purchase_review,
        stock_issue_review,
        inventory_review,
        pr_review
    )

    dashboard_summary = build_dashboard_summary(
        purchase_register,
        consumption_register_with_pr,
        inventory_summary,
        procurement_summary,
        subproject_consumption_summary,
        activity_consumption_summary,
        contractor_consumption_summary,
        item_consumption_summary,
        purchase_review,
        stock_issue_review,
        inventory_review,
        pr_review,
        pr_variance
    )

    final_data_quality = build_data_quality_summary(
        purchase_register,
        consumption_register_with_pr,
        inventory_summary,
        pr_register,
        purchase_review,
        stock_issue_review,
        inventory_review,
        pr_review,
        pr_variance,
        no_pr_reference
    )

    procurement_consumption_comparison = (
        build_procurement_consumption_comparison(
            purchase_register,
            consumption_register_with_pr
        )
    )

    # --------------------------------------------------------
    # FINAL KPI VALUES
    # --------------------------------------------------------

    final_procurement_principal = (
        purchase_register[
            "PB Principal Amount"
        ].sum()
    )

    final_procurement_total = (
        purchase_register[
            "PB Total Bill Amount"
        ].sum()
    )

    final_consumption_cost = (
        consumption_register_with_pr[
            "Actual Consumption Cost"
        ].sum()
    )

    final_closing_stock_value = (
        inventory_summary[
            "Closing Stock Value"
        ].sum()
    )

    final_received_stock_value = (
        inventory_summary[
            "Stock Received Amount"
        ].sum()
    )

    final_issue_count = len(
        consumption_register_with_pr
    )

    final_subproject_count = (
        consumption_register_with_pr[
            "Actual Issue Subproject"
        ]
        .replace("", pd.NA)
        .dropna()
        .nunique()
    )

    genuine_audit_count = (
        len(purchase_review)
        + len(stock_issue_review)
        + len(inventory_review)
        + len(pr_review)
    )

    # --------------------------------------------------------
    # FINAL DASHBOARD METRICS
    # --------------------------------------------------------

    dashboard_c1, dashboard_c2, dashboard_c3, dashboard_c4 = (
        st.columns(4)
    )

    show_metric(
        dashboard_c1,
        "Procurement Principal Cost",
        final_procurement_principal,
        "money"
    )

    show_metric(
        dashboard_c2,
        "Actual Consumption Cost",
        final_consumption_cost,
        "money"
    )

    show_metric(
        dashboard_c3,
        "Closing Stock Value",
        final_closing_stock_value,
        "money"
    )

    show_metric(
        dashboard_c4,
        "Stock Ledger Received Value",
        final_received_stock_value,
        "money"
    )

    dashboard_c5, dashboard_c6, dashboard_c7, dashboard_c8 = (
        st.columns(4)
    )

    show_metric(
        dashboard_c5,
        "Total Purchase Bill Value",
        final_procurement_total,
        "money"
    )

    show_metric(
        dashboard_c6,
        "Actual Issue Rows",
        final_issue_count,
        "integer"
    )

    show_metric(
        dashboard_c7,
        "Consuming Subprojects",
        final_subproject_count,
        "integer"
    )

    show_metric(
        dashboard_c8,
        "Genuine Audit Rows",
        genuine_audit_count,
        "integer"
    )

    # --------------------------------------------------------
    # FINAL DASHBOARD TABLES
    # --------------------------------------------------------

    show_dataframe(
        "Final Data Quality Summary",
        final_data_quality,
        maximum_rows=100
    )

    show_dataframe(
        "Procurement versus Consumption — Informational",
        procurement_consumption_comparison,
        maximum_rows=100
    )

    # --------------------------------------------------------
    # TOP COST TABLES
    # --------------------------------------------------------

    top_subprojects = (
        subproject_consumption_summary
        .sort_values(
            "Actual Consumption Cost",
            ascending=False
        )
        .head(20)
    )

    top_items = (
        item_consumption_summary
        .sort_values(
            "Actual Consumption Cost",
            ascending=False
        )
        .head(20)
    )

    top_dashboard_c1, top_dashboard_c2 = (
        st.columns(2)
    )

    with top_dashboard_c1:
        st.subheader(
            "Top 20 Subprojects by Consumption"
        )

        st.dataframe(
            top_subprojects,
            use_container_width=True,
            hide_index=True
        )

    with top_dashboard_c2:
        st.subheader(
            "Top 20 Items by Consumption"
        )

        st.dataframe(
            top_items,
            use_container_width=True,
            hide_index=True
        )

    if not contractor_consumption_summary.empty:

        top_contractors = (
            contractor_consumption_summary
            .sort_values(
                "Actual Consumption Cost",
                ascending=False
            )
            .head(20)
        )

        show_dataframe(
            "Top 20 Contractors by Material Consumption",
            top_contractors,
            maximum_rows=20
        )

    if not activity_consumption_summary.empty:

        top_activities = (
            activity_consumption_summary
            .sort_values(
                "Actual Consumption Cost",
                ascending=False
            )
            .head(20)
        )

        show_dataframe(
            "Top 20 Activities by Material Consumption",
            top_activities,
            maximum_rows=20
        )

    # --------------------------------------------------------
    # FINAL AUDIT DISPLAY
    # --------------------------------------------------------

    if final_audit_report.empty:

        st.success(
            "No genuine audit issues were found."
        )

    else:

        with st.expander(
            "Final Genuine Audit Report",
            expanded=False
        ):

            show_dataframe(
                "Audit Report",
                final_audit_report,
                maximum_rows=500
            )

  --------------------------------------------------
    # --------------------------------------------------------
    # FINAL EXCEL WORKBOOK — 9 SHEETS ONLY
    # --------------------------------------------------------

    final_output_sheets = {
        "Dashboard":
            dashboard_summary,

        "Subproject Cost Summary":
            subproject_consumption_summary,

        "Activity Summary":
            activity_consumption_summary,

        "Contractor Summary":
            contractor_consumption_summary,

        "Item Summary":
            item_consumption_summary,

        "Inventory Summary":
            inventory_summary,

        "Consumption Register":
            consumption_register_with_pr,

        "Procurement Register":
            purchase_register,

        "Audit Report":
            final_audit_report
    }

    final_excel_output = create_excel_workbook(
        final_output_sheets
    )

    final_filename = (
        "StrategicERP_Cost_Intelligence_"
        f"{datetime.now():%Y%m%d_%H%M}.xlsx"
    )

    st.download_button(
        label="Download Final Cost Intelligence Excel",
        data=final_excel_output,
        file_name=final_filename,
        mime=(
            "application/vnd.openxmlformats-"
            "officedocument.spreadsheetml.sheet"
        )
    )

    st.success(
        "Final report created successfully with only 9 Excel sheets."
    )
