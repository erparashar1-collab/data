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

# Your Excel file should be uploaded to the same GitHub repository
# as this app.py file.
DEFAULT_XLSX = (
    APP_DIR
    / "FY  25-26 Metal Pallet Stock -Palwal 13-08-2026(1).xlsx"
)


# These are the 5 sheets in your workbook.
SHEET_NAMES = [
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

    </style>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# TITLE
# ============================================================

st.title("📦 Metal Pallet Stock – Palwal")

st.caption(
    "Inventory dashboard • TOTAL STOCK (Nos) is the original stock. "
    "AVAILABLE is the current live quantity."
)


# ============================================================
# FIND EXCEL FILE
# ============================================================

def find_workbook():
    """
    Find the Excel workbook.

    First looks for the exact expected filename.
    If it doesn't exist, looks for the first .xlsx file
    in the application folder.
    """

    if DEFAULT_XLSX.exists():
        return DEFAULT_XLSX

    candidates = sorted(APP_DIR.glob("*.xlsx"))

    if candidates:
        return candidates[0]

    return None


xlsx_path = find_workbook()

if not xlsx_path:

    st.error(
        """
        Excel workbook not found.

        Please upload the Excel file to the same GitHub repository
        as app.py.
        """
    )

    st.stop()


# ============================================================
# HELPER FUNCTIONS
# ============================================================

def clean_text(value):
    """
    Safely convert any Excel value to text.

    This fixes the error:

    TypeError: argument of type 'float' is not a container or iterable
    """

    if pd.isna(value):
        return ""

    return str(value).strip()


def number_or_zero(value):
    """
    Convert Excel number/value into integer.

    Blank or invalid values become 0.
    """

    if pd.isna(value):
        return 0

    try:
        return int(float(value))

    except (ValueError, TypeError):
        return 0


def find_header_row(raw):
    """
    Find the row containing the stock headers.

    Works safely even when cells contain numbers or blanks.
    """

    for i in range(len(raw)):

        row = [
            clean_text(x).upper()
            for x in raw.iloc[i].tolist()
        ]

        has_total_stock = any(
            "TOTAL STOCK" in x
            for x in row
        )

        has_available = any(
            "AVAILABLE" in x
            for x in row
        )

        has_nos = "NOS" in row

        if has_total_stock:
            return i

        if has_nos and has_available:
            return i

    # If no header is detected, use first row.
    return 0


def find_columns(raw, header_row):
    """
    Determine TOTAL STOCK and AVAILABLE columns.
    """

    headers = [
        clean_text(x)
        for x in raw.iloc[header_row].tolist()
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


# ============================================================
# LOAD EXCEL WORKBOOK
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
                    "Sr No",
                    "Item / Size",
                    "TOTAL STOCK (Nos)",
                    "Initial Available",
                ]
            )

            continue

        # ----------------------------------------------------
        # Find header
        # ----------------------------------------------------

        header_row = find_header_row(raw)

        # ----------------------------------------------------
        # Find stock columns
        # ----------------------------------------------------

        total_col, available_col = find_columns(
            raw,
            header_row,
        )

        # ----------------------------------------------------
        # Fallback column detection
        # ----------------------------------------------------

        if total_col is None:

            # Based on the workbook structure:
            # Fixed Metal Pallet normally has total stock
            # around column D.
            #
            # Other sheets normally have it around column C.

            if raw.shape[1] >= 4:
                total_col = 2

            else:
                total_col = min(
                    1,
                    raw.shape[1] - 1
                )

        if available_col is None:

            available_col = min(
                total_col + 1,
                raw.shape[1] - 1
            )

        # ----------------------------------------------------
        # Read data rows
        # ----------------------------------------------------

        rows = []

        for i in range(header_row + 1, len(raw)):

            values = raw.iloc[i].tolist()

            if not values:
                continue

            # -----------------------------------------------
            # Ignore completely blank rows
            # -----------------------------------------------

            if all(
                clean_text(v) == ""
                for v in values
            ):
                continue

            # -----------------------------------------------
            # Find serial number
            # -----------------------------------------------

            serial = None
            serial_col = None

            # Search first 3 columns for serial number.
            for col_index, value in enumerate(
                values[:3]
            ):

                if pd.isna(value):
                    continue

                try:

                    number = float(value)

                    if (
                        number.is_integer()
                        and number >= 1
                    ):

                        serial = int(number)
                        serial_col = col_index

                        break

                except (
                    ValueError,
                    TypeError,
                ):
                    pass

            # No serial number = not an inventory row.
            if serial is None:
                continue

            # -----------------------------------------------
            # Determine item column
            # -----------------------------------------------

            if serial_col is not None:

                item_col = serial_col + 1

            else:

                item_col = 1

            if item_col >= len(values):
                item = ""

            else:
                item = clean_text(values[item_col])

            # -----------------------------------------------
            # Total stock
            # -----------------------------------------------

            if total_col < len(values):

                total_stock = number_or_zero(
                    values[total_col]
                )

            else:

                total_stock = 0

            # -----------------------------------------------
            # Available stock
            # -----------------------------------------------

            if available_col < len(values):

                available = number_or_zero(
                    values[available_col]
                )

            else:

                available = 0

            # -----------------------------------------------
            # Make sure available is valid
            # -----------------------------------------------

            available = max(
                0,
                min(
                    available,
                    total_stock,
                ),
            )

            rows.append(
                {
                    "Sr No": serial,
                    "Item / Size": item,
                    "TOTAL STOCK (Nos)": total_stock,
                    "Initial Available": available,
                }
            )

        # ----------------------------------------------------
        # Create dataframe
        # ----------------------------------------------------

        df = pd.DataFrame(rows)

        if df.empty:

            df = pd.DataFrame(
                columns=[
                    "Sr No",
                    "Item / Size",
                    "TOTAL STOCK (Nos)",
                    "Initial Available",
                ]
            )

        result[sheet_name] = df

    return result


# ============================================================
# LOAD DATA
# ============================================================

try:

    data = load_workbook_data(
        str(xlsx_path)
    )

except Exception as e:

    st.error(
        "There was a problem reading the Excel workbook."
    )

    st.exception(e)

    st.stop()


# ============================================================
# STATE
# ============================================================

def state_key(sheet, serial_number):

    return f"{sheet}__{serial_number}"


def load_saved_state():

    if not STATE_FILE.exists():
        return {}

    try:

        with open(
            STATE_FILE,
            "r",
            encoding="utf-8",
        ) as file:

            return json.load(file)

    except Exception:

        return {}


# ------------------------------------------------------------
# Initialize Streamlit state
# ------------------------------------------------------------

if "stock_state" not in st.session_state:

    saved_state = load_saved_state()

    st.session_state.stock_state = saved_state

    # Seed quantities from Excel.
    for sheet_name, df in data.items():

        for _, row in df.iterrows():

            serial_number = int(
                row["Sr No"]
            )

            key = state_key(
                sheet_name,
                serial_number,
            )

            if key not in st.session_state.stock_state:

                st.session_state.stock_state[key] = int(
                    row["Initial Available"]
                )


# ============================================================
# SAVE STATE
# ============================================================

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
        # Some Streamlit hosting environments
        # don't allow permanent local file writes.
        pass


# ============================================================
# GET AVAILABLE
# ============================================================

def get_available(
    sheet_name,
    serial_number,
):

    key = state_key(
        sheet_name,
        serial_number,
    )

    return int(
        st.session_state.stock_state.get(
            key,
            0,
        )
    )


# ============================================================
# SET AVAILABLE
# ============================================================

def set_available(
    sheet_name,
    serial_number,
    value,
):

    df = data[sheet_name]

    row = df.loc[
        df["Sr No"] == serial_number
    ]

    if row.empty:
        return

    total_stock = int(
        row.iloc[0]["TOTAL STOCK (Nos)"]
    )

    # Never allow negative stock.
    value = max(
        0,
        int(value),
    )

    # Never allow available > total stock.
    value = min(
        value,
        total_stock,
    )

    key = state_key(
        sheet_name,
        serial_number,
    )

    st.session_state.stock_state[key] = value

    save_state()


# ============================================================
# OVERALL STOCK
# ============================================================

overall_total = 0
overall_available = 0

for sheet_name, df in data.items():

    overall_total += int(
        df["TOTAL STOCK (Nos)"].sum()
    )

    for serial_number in df["Sr No"].tolist():

        overall_available += get_available(
            sheet_name,
            int(serial_number),
        )


overall_not_available = (
    overall_total
    - overall_available
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
# 5 SHEET TOGGLE
# ============================================================

available_sheet_names = [
    sheet
    for sheet in SHEET_NAMES
    if sheet in data
]

# Add any unexpected workbook sheets as well.
for sheet in data.keys():

    if sheet not in available_sheet_names:

        available_sheet_names.append(sheet)


selected_display = st.radio(
    "Select stock category",
    [
        DISPLAY_NAMES.get(
            sheet,
            sheet,
        )
        for sheet in available_sheet_names
    ],
    horizontal=True,
)


selected_sheet = next(
    sheet
    for sheet in available_sheet_names
    if DISPLAY_NAMES.get(
        sheet,
        sheet,
    ) == selected_display
)


# ============================================================
# SELECTED SHEET
# ============================================================

source_df = data[selected_sheet]

sheet_total = int(
    source_df["TOTAL STOCK (Nos)"].sum()
)

sheet_available = sum(
    get_available(
        selected_sheet,
        int(serial_number),
    )
    for serial_number in source_df["Sr No"].tolist()
)

sheet_not_available = (
    sheet_total
    - sheet_available
)


# ============================================================
# SHEET SUMMARY
# ============================================================

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
    "Use − / + to change AVAILABLE quantity. "
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
# TABLE HEADER
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

for _, row in display_df.iterrows():

    serial_number = int(
        row["Sr No"]
    )

    item = row["Item / Size"]

    total_stock = int(
        row["TOTAL STOCK (Nos)"]
    )

    available = get_available(
        selected_sheet,
        serial_number,
    )

    not_available = (
        total_stock
        - available
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

    # Sr No
    cols[0].write(
        serial_number
    )

    # Item
    cols[1].write(
        item
    )

    # Total
    cols[2].write(
        f"{total_stock:,}"
    )

    # Available
    cols[3].markdown(
        f"**{available:,}**"
    )

    # Not Available
    cols[4].write(
        f"{not_available:,}"
    )

    # Buttons
    minus_col, plus_col = cols[5].columns(2)

    # --------------------------------------------------------
    # MINUS
    # --------------------------------------------------------

    if minus_col.button(
        "−",
        key=f"minus_{selected_sheet}_{serial_number}",
        use_container_width=True,
    ):

        set_available(
            selected_sheet,
            serial_number,
            available - 1,
        )

        st.rerun()

    # --------------------------------------------------------
    # PLUS
    # --------------------------------------------------------

    if plus_col.button(
        "+",
        key=f"plus_{selected_sheet}_{serial_number}",
        use_container_width=True,
    ):

        set_available(
            selected_sheet,
            serial_number,
            available + 1,
        )

        st.rerun()


# ============================================================
# EXPORT
# ============================================================

st.divider()

st.subheader(
    "📥 Export Current Stock"
)


export_rows = []


for sheet_name, df in data.items():

    for _, row in df.iterrows():

        serial_number = int(
            row["Sr No"]
        )

        total_stock = int(
            row["TOTAL STOCK (Nos)"]
        )

        available = get_available(
            sheet_name,
            serial_number,
        )

        not_available = (
            total_stock
            - available
        )

        export_rows.append(
            {
                "Sheet": DISPLAY_NAMES.get(
                    sheet_name,
                    sheet_name,
                ),
                "Sr No": serial_number,
                "Item / Size": row[
                    "Item / Size"
                ],
                "TOTAL STOCK (Nos)": total_stock,
                "AVAILABLE QTY": available,
                "NOT AVAILABLE": not_available,
            }
        )


export_df = pd.DataFrame(
    export_rows
)


st.download_button(
    label="⬇️ Download Current Stock as CSV",
    data=export_df.to_csv(
        index=False
    ).encode("utf-8"),
    file_name="metal_pallet_current_stock.csv",
    mime="text/csv",
)


# ============================================================
# RESET
# ============================================================

st.divider()


if st.button(
    "↩️ Reset All AVAILABLE Quantities to Excel Values"
):

    for sheet_name, df in data.items():

        for _, row in df.iterrows():

            serial_number = int(
                row["Sr No"]
            )

            key = state_key(
                sheet_name,
                serial_number,
            )

            st.session_state.stock_state[
                key
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
    "TOTAL STOCK (Nos) = original stock from Excel. "
    "AVAILABLE = live stock quantity. "
    "NOT AVAILABLE = TOTAL STOCK − AVAILABLE."
)
