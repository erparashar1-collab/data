import io
import json
from pathlib import Path

import pandas as pd
import streamlit as st


# ============================================================
# PAGE CONFIG
# ============================================================

st.set_page_config(
    page_title="Metal Pallet Stock – Palwal",
    page_icon="📦",
    layout="wide",
)


# ============================================================
# CONFIGURATION
# ============================================================

APP_DIR = Path(__file__).parent

STATE_FILE = APP_DIR / "stock_state.json"

DEFAULT_XLSX = (
    APP_DIR
    / "FY  25-26 Metal Pallet Stock -Palwal 13-08-2026(1).xlsx"
)

EXPECTED_SHEETS = [
    "Fixed Metal Pallet",
    "Pipe-Profile ( width )  ",
    " U-Profile (Center)",
    "L--Members (Top)",
    "Length Middle (LM)",
]

DISPLAY_NAMES = {
    "Fixed Metal Pallet": "Fixed Metal Pallet",
    "Pipe-Profile ( width )  ": "Pipe-Profile (Width)",
    " U-Profile (Center)": "U-Profile (Center)",
    "L--Members (Top)": "L-Members (Top)",
    "Length Middle (LM)": "Length Middle (LM)",
}


# ============================================================
# CSS
# ============================================================

st.markdown(
    """
    <style>
    .block-container {
        padding-top: 1.5rem;
        padding-bottom: 2rem;
    }

    .stock-card {
        padding: 1rem;
        border-radius: 12px;
        border: 1px solid rgba(128,128,128,.25);
        background: rgba(128,128,128,.06);
        text-align: center;
    }

    .stock-number {
        font-size: 2rem;
        font-weight: 700;
    }

    .stock-label {
        font-size: .85rem;
        opacity: .75;
    }

    .small-note {
        font-size: .85rem;
        opacity: .70;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# TITLE
# ============================================================

st.title("📦 Metal Pallet Stock – Palwal")

st.caption(
    "TOTAL STOCK (Nos) = original stock from Excel. "
    "AVAILABLE = current live stock. "
    "NOT AVAILABLE = TOTAL STOCK − AVAILABLE."
)


# ============================================================
# FIND EXCEL FILE
# ============================================================

def find_workbook():
    if DEFAULT_XLSX.exists():
        return DEFAULT_XLSX

    candidates = sorted(APP_DIR.glob("*.xlsx"))

    if candidates:
        return candidates[0]

    return None


xlsx_path = find_workbook()

if not xlsx_path:
    st.error(
        "Excel workbook not found. Upload the original .xlsx file "
        "to the same GitHub repository as app.py."
    )
    st.stop()


# ============================================================
# HELPER FUNCTIONS
# ============================================================

def clean_text(value):
    """Safely convert an Excel cell to text."""
    if pd.isna(value):
        return ""
    return str(value).strip()


def number_or_zero(value):
    """Safely convert an Excel value to an integer."""
    if pd.isna(value):
        return 0

    try:
        return int(float(value))
    except (ValueError, TypeError):
        return 0


def find_header_row(raw):
    """
    Find the header row without assuming every Excel cell is text.
    This prevents:
    TypeError: argument of type 'float' is not a container or iterable
    """

    for i in range(len(raw)):
        row = [
            clean_text(value).upper()
            for value in raw.iloc[i].tolist()
        ]

        has_total = any(
            "TOTAL STOCK" in value
            for value in row
        )

        has_available = any(
            "AVAILABLE" in value
            for value in row
        )

        has_nos = "NOS" in row

        if has_total:
            return i

        if has_nos and has_available:
            return i

    return 0


def find_columns(raw, header_row):
    """Find TOTAL STOCK and AVAILABLE columns."""

    headers = [
        clean_text(value)
        for value in raw.iloc[header_row].tolist()
    ]

    total_col = None
    available_col = None

    for index, header in enumerate(headers):
        upper_header = header.upper()

        if (
            "TOTAL STOCK" in upper_header
            or upper_header == "NOS"
            or upper_header == "QTY"
            or "QTY (NOS)" in upper_header
        ):
            if total_col is None:
                total_col = index

        if "AVAILABLE" in upper_header:
            if available_col is None:
                available_col = index

    return total_col, available_col


def find_serial_and_item(values):
    """
    Find serial number and item column.

    The workbook has sheets with slightly different layouts,
    so we inspect the first few columns instead of hard-coding
    one layout.
    """

    serial = None
    serial_col = None

    for col_index, value in enumerate(values[:5]):

        if pd.isna(value):
            continue

        try:
            number = float(value)

            if number.is_integer() and number >= 1:
                serial = int(number)
                serial_col = col_index
                break

        except (ValueError, TypeError):
            continue

    if serial_col is None:
        return None, None, None

    item_col = serial_col + 1

    item = ""
    if item_col < len(values):
        item = clean_text(values[item_col])

    return serial, serial_col, item


# ============================================================
# LOAD EXCEL
# ============================================================

@st.cache_data
def load_workbook_data(path_string):

    excel = pd.ExcelFile(path_string)

    result = {}

    for sheet_name in excel.sheet_names:

        raw = pd.read_excel(
            path_string,
            sheet_name=sheet_name,
            header=None,
        )

        if raw.empty:
            result[sheet_name] = pd.DataFrame(
                columns=[
                    "Row ID",
                    "Sr No",
                    "Item / Size",
                    "TOTAL STOCK (Nos)",
                    "Initial Available",
                ]
            )
            continue

        header_row = find_header_row(raw)

        total_col, available_col = find_columns(
            raw,
            header_row,
        )

        # Fallbacks based on the workbook structure.
        if total_col is None:
            if raw.shape[1] >= 4:
                total_col = 2
            else:
                total_col = max(0, raw.shape[1] - 2)

        if available_col is None:
            available_col = min(
                total_col + 1,
                raw.shape[1] - 1,
            )

        rows = []

        for excel_row_number in range(
            header_row + 1,
            len(raw),
        ):

            values = raw.iloc[excel_row_number].tolist()

            if not values:
                continue

            # Ignore completely blank rows.
            if all(
                clean_text(value) == ""
                for value in values
            ):
                continue

            # Ignore obvious total/formula rows.
            row_text = [
                clean_text(value).lower()
                for value in values
            ]

            if any(
                value == "total"
                or value.startswith("=sum(")
                for value in row_text
            ):
                continue

            serial, serial_col, item = find_serial_and_item(
                values
            )

            if serial is None:
                continue

            # Unique ID is based on the actual Excel row.
            # This is deliberately NOT Sr No.
            row_id = f"{sheet_name}__excelrow_{excel_row_number}"

            if total_col < len(values):
                total_stock = number_or_zero(
                    values[total_col]
                )
            else:
                total_stock = 0

            if available_col < len(values):
                available = number_or_zero(
                    values[available_col]
                )
            else:
                available = 0

            # Keep availability inside valid bounds.
            available = max(
                0,
                min(
                    available,
                    total_stock,
                ),
            )

            rows.append(
                {
                    "Row ID": row_id,
                    "Sr No": serial,
                    "Item / Size": item,
                    "TOTAL STOCK (Nos)": total_stock,
                    "Initial Available": available,
                }
            )

        result[sheet_name] = pd.DataFrame(
            rows,
            columns=[
                "Row ID",
                "Sr No",
                "Item / Size",
                "TOTAL STOCK (Nos)",
                "Initial Available",
            ],
        )

    return result


# ============================================================
# LOAD DATA
# ============================================================

try:
    data = load_workbook_data(
        str(xlsx_path)
    )

except Exception as error:
    st.error("There was a problem reading the Excel workbook.")
    st.exception(error)
    st.stop()


if not data:
    st.error("No worksheets were found in the workbook.")
    st.stop()


# ============================================================
# STATE
# ============================================================

def state_key(row_id):
    """
    Every physical Excel row gets its own stock quantity.
    This prevents duplicate Sr No values from sharing stock.
    """
    return f"stock__{row_id}"


def load_saved_state():

    if not STATE_FILE.exists():
        return {}

    try:
        with open(
            STATE_FILE,
            "r",
            encoding="utf-8",
        ) as file:
            saved = json.load(file)

        if isinstance(saved, dict):
            return saved

        return {}

    except Exception:
        return {}


if "stock_state" not in st.session_state:

    saved_state = load_saved_state()

    st.session_state.stock_state = saved_state

    # Seed from Excel.
    for sheet_name, df in data.items():

        for _, row in df.iterrows():

            row_id = str(row["Row ID"])

            key = state_key(row_id)

            if key not in st.session_state.stock_state:
                st.session_state.stock_state[key] = int(
                    row["Initial Available"]
                )


def save_state():

    try:
        with open(
            STATE_FILE,
            "w",
            encoding="utf-8",
        ) as file:

            json.dump(
                st.session_state.stock_state,
                file,
                indent=2,
            )

    except Exception:
        # Streamlit Cloud may use an ephemeral filesystem.
        # The current session will still work.
        pass


def get_available(row_id):

    return int(
        st.session_state.stock_state.get(
            state_key(row_id),
            0,
        )
    )


def set_available(
    row_id,
    value,
    total_stock,
):

    value = max(
        0,
        int(value),
    )

    value = min(
        value,
        int(total_stock),
    )

    st.session_state.stock_state[
        state_key(row_id)
    ] = value

    save_state()


# ============================================================
# OVERALL TOTALS
# ============================================================

overall_total = 0
overall_available = 0

for sheet_name, df in data.items():

    overall_total += int(
        df["TOTAL STOCK (Nos)"].sum()
    )

    for _, row in df.iterrows():

        overall_available += get_available(
            str(row["Row ID"])
        )


overall_not_available = (
    overall_total - overall_available
)


# ============================================================
# TOP DASHBOARD
# ============================================================

col1, col2, col3 = st.columns(3)

with col1:
    st.markdown(
        f"""
        <div class="stock-card">
            <div class="stock-number">
                {overall_total:,}
            </div>
            <div class="stock-label">
                TOTAL STOCK (Nos)
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

with col2:
    st.markdown(
        f"""
        <div class="stock-card">
            <div class="stock-number">
                {overall_available:,}
            </div>
            <div class="stock-label">
                AVAILABLE NOW
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

with col3:
    st.markdown(
        f"""
        <div class="stock-card">
            <div class="stock-number">
                {overall_not_available:,}
            </div>
            <div class="stock-label">
                NOT AVAILABLE
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


st.divider()


# ============================================================
# SHEET TOGGLES
# ============================================================

available_sheets = []

for sheet_name in EXPECTED_SHEETS:

    if sheet_name in data:
        available_sheets.append(sheet_name)

# Add any extra sheets if the workbook changes later.
for sheet_name in data:

    if sheet_name not in available_sheets:
        available_sheets.append(sheet_name)


selected_display = st.radio(
    "Select stock category",
    [
        DISPLAY_NAMES.get(
            sheet_name,
            sheet_name,
        )
        for sheet_name in available_sheets
    ],
    horizontal=True,
)


selected_sheet = next(
    sheet_name
    for sheet_name in available_sheets
    if DISPLAY_NAMES.get(
        sheet_name,
        sheet_name,
    ) == selected_display
)


# ============================================================
# SELECTED SHEET SUMMARY
# ============================================================

source_df = data[selected_sheet]

sheet_total = int(
    source_df["TOTAL STOCK (Nos)"].sum()
)

sheet_available = sum(
    get_available(str(row["Row ID"]))
    for _, row in source_df.iterrows()
)

sheet_not_available = (
    sheet_total - sheet_available
)


col1, col2, col3 = st.columns(3)

with col1:
    st.metric(
        "TOTAL STOCK (Nos)",
        f"{sheet_total:,}",
    )

with col2:
    st.metric(
        "AVAILABLE NOW",
        f"{sheet_available:,}",
    )

with col3:
    st.metric(
        "NOT AVAILABLE",
        f"{sheet_not_available:,}",
    )


st.subheader(
    DISPLAY_NAMES.get(
        selected_sheet,
        selected_sheet,
    )
)

st.info(
    "Use − / + to change AVAILABLE. "
    "TOTAL STOCK (Nos) never changes."
)


# ============================================================
# SEARCH
# ============================================================

search = st.text_input(
    "🔎 Search item / size / length",
    placeholder="e.g. 1000, 875 x 1035, 1200",
)

display_df = source_df.copy()

if search.strip():

    query = search.strip().lower()

    display_df = display_df[
        display_df["Item / Size"]
        .astype(str)
        .str.lower()
        .str.contains(
            query,
            na=False,
        )
    ].copy()


# ============================================================
# INVENTORY TABLE HEADER
# ============================================================

header = st.columns(
    [
        0.7,
        3.0,
        1.5,
        1.5,
        1.5,
        1.3,
    ]
)

header[0].markdown("**Sr No**")
header[1].markdown("**Item / Size**")
header[2].markdown("**TOTAL STOCK (Nos)**")
header[3].markdown("**AVAILABLE**")
header[4].markdown("**NOT AVAILABLE**")
header[5].markdown("**CHANGE**")


# ============================================================
# INVENTORY ROWS
# ============================================================

for display_position, (_, row) in enumerate(
    display_df.iterrows()
):

    row_id = str(row["Row ID"])

    serial_number = int(
        row["Sr No"]
    )

    item = row["Item / Size"]

    total_stock = int(
        row["TOTAL STOCK (Nos)"]
    )

    available = get_available(
        row_id
    )

    not_available = (
        total_stock - available
    )

    cols = st.columns(
        [
            0.7,
            3.0,
            1.5,
            1.5,
            1.5,
            1.3,
        ]
    )

    cols[0].write(
        serial_number
    )

    cols[1].write(
        item
    )

    cols[2].write(
        f"{total_stock:,}"
    )

    cols[3].markdown(
        f"**{available:,}**"
    )

    cols[4].write(
        f"{not_available:,}"
    )

    minus_col, plus_col = cols[5].columns(2)

    # --------------------------------------------------------
    # IMPORTANT:
    # Button keys use Row ID, not Sr No.
    # This prevents duplicate element keys.
    # --------------------------------------------------------

    minus_key = (
        f"minus__{row_id}"
    )

    plus_key = (
        f"plus__{row_id}"
    )

    if minus_col.button(
        "−",
        key=minus_key,
        use_container_width=True,
    ):

        set_available(
            row_id,
            available - 1,
            total_stock,
        )

        st.rerun()

    if plus_col.button(
        "+",
        key=plus_key,
        use_container_width=True,
    ):

        set_available(
            row_id,
            available + 1,
            total_stock,
        )

        st.rerun()


# ============================================================
# EXPORT SECTION
# ============================================================

st.divider()

st.subheader(
    "📥 Download Current Stock"
)

st.write(
    "The Excel download contains all 5 stock categories as separate "
    "worksheets, matching the 5 toggles above."
)


# ============================================================
# CREATE MULTI-SHEET EXCEL FILE
# ============================================================

def create_excel_download():

    output = io.BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl",
    ) as writer:

        for sheet_name in available_sheets:

            source = data[sheet_name]

            export_rows = []

            for _, row in source.iterrows():

                row_id = str(
                    row["Row ID"]
                )

                serial_number = int(
                    row["Sr No"]
                )

                total_stock = int(
                    row["TOTAL STOCK (Nos)"]
                )

                available = get_available(
                    row_id
                )

                not_available = (
                    total_stock - available
                )

                export_rows.append(
                    {
                        "Sr No": serial_number,
                        "Item / Size": row[
                            "Item / Size"
                        ],
                        "TOTAL STOCK (Nos)": total_stock,
                        "AVAILABLE": available,
                        "NOT AVAILABLE": not_available,
                    }
                )

            export_df = pd.DataFrame(
                export_rows
            )

            # Excel sheet names cannot be longer than 31 characters.
            safe_sheet_name = DISPLAY_NAMES.get(
                sheet_name,
                sheet_name,
            )[:31]

            export_df.to_excel(
                writer,
                sheet_name=safe_sheet_name,
                index=False,
            )

            # Basic Excel formatting.
            worksheet = writer.sheets[
                safe_sheet_name
            ]

            worksheet.freeze_panes = "A2"

            for column_cells in worksheet.columns:

                max_length = 0

                column_letter = (
                    column_cells[0].column_letter
                )

                for cell in column_cells:

                    try:
                        value_length = len(
                            str(cell.value)
                        )

                        max_length = max(
                            max_length,
                            value_length,
                        )

                    except Exception:
                        pass

                worksheet.column_dimensions[
                    column_letter
                ].width = min(
                    max_length + 2,
                    45,
                )

    output.seek(0)

    return output.getvalue()


excel_download = create_excel_download()


st.download_button(
    label="⬇️ Download Current Stock – Excel with 5 Sheets",
    data=excel_download,
    file_name="metal_pallet_current_stock.xlsx",
    mime=(
        "application/vnd.openxmlformats-officedocument."
        "spreadsheetml.sheet"
    ),
    use_container_width=True,
)


st.caption(
    "The downloaded Excel file contains one worksheet for each "
    "stock toggle/category."
)


# ============================================================
# OPTIONAL: DOWNLOAD CURRENT TOGGLE AS CSV
# ============================================================

st.subheader(
    "CSV for Current Toggle"
)

current_export = []

for _, row in source_df.iterrows():

    row_id = str(
        row["Row ID"]
    )

    total_stock = int(
        row["TOTAL STOCK (Nos)"]
    )

    available = get_available(
        row_id
    )

    current_export.append(
        {
            "Sr No": int(row["Sr No"]),
            "Item / Size": row["Item / Size"],
            "TOTAL STOCK (Nos)": total_stock,
            "AVAILABLE": available,
            "NOT AVAILABLE": (
                total_stock - available
            ),
        }
    )


current_csv_df = pd.DataFrame(
    current_export
)


st.download_button(
    label=(
        f"⬇️ Download {selected_display} as CSV"
    ),
    data=current_csv_df.to_csv(
        index=False
    ).encode("utf-8"),
    file_name=(
        "metal_pallet_"
        + selected_display.replace(" ", "_")
        + ".csv"
    ),
    mime="text/csv",
)


# ============================================================
# RESET
# ============================================================

st.divider()

st.subheader(
    "⚙️ Stock Controls"
)

if st.button(
    "↩️ Reset ALL AVAILABLE quantities to Excel values",
    use_container_width=True,
):

    for sheet_name, df in data.items():

        for _, row in df.iterrows():

            row_id = str(
                row["Row ID"]
            )

            st.session_state.stock_state[
                state_key(row_id)
            ] = int(
                row["Initial Available"]
            )

    save_state()

    st.success(
        "All AVAILABLE quantities have been reset to the original Excel values."
    )

    st.rerun()


# ============================================================
# FOOTER
# ============================================================

st.divider()

st.caption(
    "TOTAL STOCK (Nos) is never changed by the + / − controls. "
    "AVAILABLE can only move between 0 and TOTAL STOCK."
)
