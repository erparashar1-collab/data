
import json
from pathlib import Path

import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Metal Pallet Stock – Palwal",
    page_icon="📦",
    layout="wide",
)

# ------------------------------------------------------------
# Configuration
# ------------------------------------------------------------
APP_DIR = Path(__file__).parent
STATE_FILE = APP_DIR / "stock_state.json"

# Put your Excel workbook in the same GitHub repository as app.py.
# The code automatically looks for an .xlsx file if this exact name
# is not present.
DEFAULT_XLSX = APP_DIR / "FY  25-26 Metal Pallet Stock -Palwal 13-08-2026(1).xlsx"

SHEET_NAMES = [
    "Fixed Metal Pallet",
    "Pipe-Profile ( width )  ",
    " U-Profile (Center)",
    "L--Members (Top)",
    "Length Middle (LM)",
]

# ------------------------------------------------------------
# Styling
# ------------------------------------------------------------
st.markdown(
    """
    <style>
    .block-container {padding-top: 1.5rem; padding-bottom: 2rem;}
    .stock-card {
        padding: 1rem 1.2rem;
        border-radius: 12px;
        border: 1px solid rgba(128,128,128,.25);
        background: rgba(128,128,128,.06);
        text-align: center;
    }
    .stock-number {font-size: 2rem; font-weight: 700;}
    .stock-label {font-size: .85rem; opacity: .75;}
    div[data-testid="stDataEditor"] {font-size: 14px;}
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("📦 Metal Pallet Stock – Palwal")
st.caption("Inventory dashboard • TOTAL STOCK is the original/initial stock; AVAILABLE is the live quantity.")

# ------------------------------------------------------------
# Excel loading
# ------------------------------------------------------------
def find_workbook():
    if DEFAULT_XLSX.exists():
        return DEFAULT_XLSX

    candidates = sorted(APP_DIR.glob("*.xlsx"))
    if candidates:
        return candidates[0]

    return None


@st.cache_data
def load_workbook_data(path_str):
    xls = pd.ExcelFile(path_str)
    result = {}

    for sheet in xls.sheet_names:
        raw = pd.read_excel(path_str, sheet_name=sheet, header=None)

        # Find the row containing "TOTAL STOCK" (or equivalent "Nos")
        # and the row containing the first serial number.
        header_row = None
        data_start = None

        for i in range(len(raw)):
            row = raw.iloc[i].astype(str).str.upper().tolist()
            if any("TOTAL STOCK" in x for x in row) or (
                "NOS" in row and any("AVAILABLE" in x for x in row)
            ):
                header_row = i
                break

        if header_row is None:
            # Fallback: assume first row is header.
            header_row = 0

        # Determine columns from the header.
        headers = raw.iloc[header_row].tolist()
        headers = [str(x).strip() if pd.notna(x) else "" for x in headers]

        total_col = None
        available_col = None

        for j, h in enumerate(headers):
            hu = h.upper()
            if "TOTAL STOCK" in hu or hu in {"NOS", "QTY", "QTY (NOS)"}:
                # For these workbooks, TOTAL STOCK is the numeric stock column.
                if total_col is None:
                    total_col = j
            if "AVAILABLE" in hu:
                available_col = j

        # Some sheets have "Nos" on the next row and "AVAILABLE" in another column.
        # Known workbook layout uses:
        # Fixed: C = item, D = total, E = available
        # Other sheets: B = item, C = total, D/E = available.
        if total_col is None:
            total_col = 2 if raw.shape[1] >= 3 else 1

        if available_col is None:
            available_col = min(total_col + 1, raw.shape[1] - 1)

        # Find data rows: rows whose first/second/third cell contains a serial number.
        rows = []
        for i in range(header_row + 1, len(raw)):
            vals = raw.iloc[i].tolist()

            # Ignore total/formula rows and empty rows.
            text_vals = [str(v).strip().lower() if pd.notna(v) else "" for v in vals]
            if any(v == "total" or v.startswith("=sum(") for v in text_vals):
                continue

            # Determine likely serial-number column.
            serial = None
            for v in vals[:3]:
                if pd.notna(v):
                    try:
                        fv = float(v)
                        if fv.is_integer() and fv >= 1:
                            serial = int(fv)
                            break
                    except (ValueError, TypeError):
                        pass

            if serial is None:
                continue

            # Item is generally the column immediately after serial number.
            # For Fixed Metal Pallet: serial B, item C.
            # For the other sheets: serial A, item B.
            serial_col = next(
                (j for j, v in enumerate(vals[:3])
                 if pd.notna(v) and str(v).strip().isdigit()),
                0
            )
            item_col = serial_col + 1

            item = vals[item_col] if item_col < len(vals) else ""
            total = vals[total_col] if total_col < len(vals) else 0
            available = vals[available_col] if available_col < len(vals) else 0

            # Convert numeric values safely.
            def number_or_zero(x):
                if pd.isna(x):
                    return 0
                try:
                    return int(float(x))
                except (ValueError, TypeError):
                    return 0

            rows.append(
                {
                    "Sr No": serial,
                    "Item / Size": str(item).strip() if pd.notna(item) else "",
                    "TOTAL STOCK (Nos)": number_or_zero(total),
                    "Initial Available": number_or_zero(available),
                }
            )

        df = pd.DataFrame(rows)

        if df.empty:
            df = pd.DataFrame(
                columns=["Sr No", "Item / Size", "TOTAL STOCK (Nos)", "Initial Available"]
            )

        result[sheet] = df

    return result


xlsx_path = find_workbook()

if not xlsx_path:
    st.error(
        "No Excel workbook was found. Put the original .xlsx file in the same "
        "GitHub folder as app.py and redeploy."
    )
    st.stop()

data = load_workbook_data(str(xlsx_path))

# ------------------------------------------------------------
# State handling
# ------------------------------------------------------------
def state_key(sheet, sr_no):
    return f"{sheet}__{sr_no}"


def load_saved_state():
    if not STATE_FILE.exists():
        return {}

    try:
        with open(STATE_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


if "stock_state" not in st.session_state:
    saved = load_saved_state()
    st.session_state.stock_state = saved

    # Seed editable availability from the workbook only when no saved
    # value exists for that row.
    for sheet, df in data.items():
        for _, row in df.iterrows():
            key = state_key(sheet, int(row["Sr No"]))
            if key not in st.session_state.stock_state:
                st.session_state.stock_state[key] = int(row["Initial Available"])


def save_state():
    try:
        with open(STATE_FILE, "w", encoding="utf-8") as f:
            json.dump(st.session_state.stock_state, f, indent=2)
    except Exception:
        # Streamlit Cloud/local read-only environments should not break
        # the dashboard. Session state still works for the current session.
        pass


def get_available(sheet, sr_no):
    return int(st.session_state.stock_state.get(state_key(sheet, sr_no), 0))


def set_available(sheet, sr_no, value):
    df = data[sheet]
    row = df.loc[df["Sr No"] == sr_no]
    if row.empty:
        return

    total = int(row.iloc[0]["TOTAL STOCK (Nos)"])
    value = max(0, min(int(value), total))

    st.session_state.stock_state[state_key(sheet, sr_no)] = value
    save_state()


# ------------------------------------------------------------
# Overall summary
# ------------------------------------------------------------
overall_total = sum(
    int(df["TOTAL STOCK (Nos)"].sum()) for df in data.values()
)
overall_available = sum(
    get_available(sheet, int(sr))
    for sheet, df in data.items()
    for sr in df["Sr No"].tolist()
)

c1, c2, c3 = st.columns(3)
with c1:
    st.markdown(
        f'<div class="stock-card"><div class="stock-number">{overall_total:,}</div>'
        f'<div class="stock-label">TOTAL STOCK (Nos)</div></div>',
        unsafe_allow_html=True,
    )
with c2:
    st.markdown(
        f'<div class="stock-card"><div class="stock-number">{overall_available:,}</div>'
        f'<div class="stock-label">AVAILABLE NOW</div></div>',
        unsafe_allow_html=True,
    )
with c3:
    used = overall_total - overall_available
    st.markdown(
        f'<div class="stock-card"><div class="stock-number">{used:,}</div>'
        f'<div class="stock-label">USED / NOT AVAILABLE</div></div>',
        unsafe_allow_html=True,
    )

st.divider()

# ------------------------------------------------------------
# Sheet toggle
# ------------------------------------------------------------
display_names = {
    "Fixed Metal Pallet": "Fixed Metal Pallet",
    "Pipe-Profile ( width )  ": "Pipe-Profile (Width)",
    " U-Profile (Center)": "U-Profile (Center)",
    "L--Members (Top)": "L-Members (Top)",
    "Length Middle (LM)": "Length Middle (LM)",
}

selected_display = st.radio(
    "Select stock category",
    [display_names[s] for s in SHEET_NAMES],
    horizontal=True,
)

selected_sheet = next(
    s for s in SHEET_NAMES if display_names[s] == selected_display
)

df = data[selected_sheet].copy()

# Build display dataframe with live availability.
df["AVAILABLE QTY"] = [
    get_available(selected_sheet, int(sr)) for sr in df["Sr No"]
]
df["BALANCE"] = df["AVAILABLE QTY"]

# Search/filter
search = st.text_input(
    "🔎 Search item / size / length",
    placeholder="e.g. 1000, 875 x 1035, 1200",
)

if search.strip():
    q = search.strip().lower()
    mask = df["Item / Size"].astype(str).str.lower().str.contains(q, na=False)
    df = df[mask].copy()

# Category totals
sheet_total = int(data[selected_sheet]["TOTAL STOCK (Nos)"].sum())
sheet_available = sum(
    get_available(selected_sheet, int(sr))
    for sr in data[selected_sheet]["Sr No"]
)

a, b, c = st.columns(3)
with a:
    st.metric("TOTAL STOCK (Nos)", f"{sheet_total:,}")
with b:
    st.metric("AVAILABLE NOW", f"{sheet_available:,}")
with c:
    st.metric("NOT AVAILABLE", f"{sheet_total - sheet_available:,}")

st.subheader(display_names[selected_sheet])

st.info(
    "Use − / + to change available quantity. AVAILABLE can never go below 0 "
    "or above TOTAL STOCK. TOTAL STOCK is never changed."
)

# ------------------------------------------------------------
# Editable rows
# ------------------------------------------------------------
# Streamlit columns make reliable + / - buttons much easier than editing
# a data-editor cell. Each item gets its own controls.
header = st.columns([0.8, 2.8, 1.5, 1.5, 1.5, 1.2])
header[0].markdown("**Sr No**")
header[1].markdown("**Item / Size**")
header[2].markdown("**TOTAL STOCK (Nos)**")
header[3].markdown("**AVAILABLE**")
header[4].markdown("**NOT AVAILABLE**")
header[5].markdown("**Change**")

for _, row in df.iterrows():
    sr = int(row["Sr No"])
    item = row["Item / Size"]
    total = int(row["TOTAL STOCK (Nos)"])
    available = get_available(selected_sheet, sr)
    unavailable = total - available

    cols = st.columns([0.8, 2.8, 1.5, 1.5, 1.5, 1.2])

    cols[0].write(sr)
    cols[1].write(item)
    cols[2].write(f"{total:,}")
    cols[3].write(f"**{available:,}**")
    cols[4].write(f"{unavailable:,}")

    bminus, bplus = cols[5].columns(2)

    if bminus.button("−", key=f"minus_{selected_sheet}_{sr}", use_container_width=True):
        set_available(selected_sheet, sr, available - 1)
        st.rerun()

    if bplus.button("+", key=f"plus_{selected_sheet}_{sr}", use_container_width=True):
        set_available(selected_sheet, sr, available + 1)
        st.rerun()

st.divider()

# ------------------------------------------------------------
# Export current inventory
# ------------------------------------------------------------
st.subheader("Export current stock")

export_rows = []
for sheet, source_df in data.items():
    for _, row in source_df.iterrows():
        sr = int(row["Sr No"])
        total = int(row["TOTAL STOCK (Nos)"])
        available = get_available(sheet, sr)

        export_rows.append(
            {
                "Sheet": display_names.get(sheet, sheet),
                "Sr No": sr,
                "Item / Size": row["Item / Size"],
                "TOTAL STOCK (Nos)": total,
                "AVAILABLE QTY": available,
                "NOT AVAILABLE": total - available,
            }
        )

export_df = pd.DataFrame(export_rows)

st.download_button(
    "⬇️ Download current stock as CSV",
    data=export_df.to_csv(index=False).encode("utf-8"),
    file_name="metal_pallet_current_stock.csv",
    mime="text/csv",
)

if st.button("↩️ Reset all AVAILABLE quantities to Excel values"):
    for sheet, source_df in data.items():
        for _, row in source_df.iterrows():
            key = state_key(sheet, int(row["Sr No"]))
            st.session_state.stock_state[key] = int(row["Initial Available"])
    save_state()
    st.success("All available quantities were reset to the original Excel values.")
    st.rerun()

st.caption(
    "Note: Streamlit Cloud's local filesystem is not a permanent multi-user database. "
    "The app saves state locally when possible, but for permanent shared stock data "
    "across redeployments/users, connect the app to a database such as Supabase or "
    "Google Sheets."
)
