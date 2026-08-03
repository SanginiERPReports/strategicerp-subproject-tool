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
