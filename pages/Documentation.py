"""
Jim Daws Trucking — Settlement Processor Documentation
"""

import streamlit as st

st.set_page_config(
    page_title="Documentation — Jim Daws Trucking",
    page_icon="📖",
    layout="centered",
)

st.title("📖 Documentation")
st.caption("How the Settlement Processor works")
st.divider()

# ── Overview ──────────────────────────────────────────────────────────────────

st.header("Overview")
st.markdown("""
The Settlement Processor reads a weekly driver settlement PDF exported from the trucking 
management system and converts it into an Excel file formatted for direct import into QuickBooks.

Each driver's settlement becomes a **Bill** in QuickBooks, with individual line items mapped 
to the correct accounts. Any line items that could not be automatically mapped are included 
directly in the Excel sheet with their original description so they can be categorized manually 
before importing.
""")

st.divider()

# ── How to use ────────────────────────────────────────────────────────────────

st.header("How to Use")
st.markdown("""
1. Launch the app by double-clicking the launcher file
2. Your browser opens automatically at `http://localhost:8501`
3. Click **Browse files** and select the weekly settlement PDF
4. Click **Process Settlements**
5. Review the processing summary and driver breakdown
6. Click **Download Excel File**
7. Open the Excel file and fill in the Category column for any unmapped rows
8. Import into QuickBooks
""")

st.divider()

# ── PDF structure ─────────────────────────────────────────────────────────────

st.header("How the PDF is Read")
st.markdown("""
The script reads the PDF line by line and works through three sections for each driver in order.
A driver must exist in the `DRIVERS` dictionary in the script to be processed. Unknown driver 
codes are skipped with a warning printed to the terminal.
""")

st.subheader("1. OTHER PAY/DEDUCTIONS")
st.markdown("""
Contains insurance deductions, garnishments, reserve contributions, advances, fuel charges, 
and scales. The script scans each line with a dollar amount and maps it if it matches a known 
pattern. Lines that are not recognized are included in the Excel output with their original 
description and a category of `OTHER PAY/DEDUCTIONS` so they can be filled in manually.
""")

st.subheader("2. TOTAL SETTLEMENT")
st.markdown("""
Contains the driver's gross pay and total trip expenses. Gross pay becomes the **Settlements** 
line item and trip expenses become the **Diesel** line item. Any fuel or DEF charges found in 
OTHER PAY/DEDUCTIONS that were not already baked into the trip expenses total are added on top.
""")

st.subheader("3. RESERVES")
st.markdown("""
Shows running balances for each reserve account — Escrow, Licensing, Maintenance, and Loan. 
The script ignores opening balances and "Addition to Reserve" lines since those amounts are 
already captured in OTHER PAY/DEDUCTIONS. Dated transaction lines within each reserve 
(maintenance charges, repairs, parts) are captured and routed to their correct QB account 
based on which reserve they belong to. Unrecognized reserve types are included in the Excel 
output with a category of `RESERVE UNKNOWN`.
""")

st.divider()

# ── Mappings ──────────────────────────────────────────────────────────────────

st.header("Current Mappings")
st.markdown("These are the line items the script currently maps automatically:")

st.subheader("From TOTAL SETTLEMENT")
st.table({
    "PDF Line Item": [
        "PERCENTAGE PAY(...)",
        "TOTAL TRIP EXPENSES + fuel/DEF from OTHER PAY",
    ],
    "Excel Description": [
        "Settlements",
        "Diesel",
    ],
    "QuickBooks Category": [
        "53400.3500 Driver Pay:I/C Settlements",
        "54100.1500 Truck Expense:Diesel",
    ]
})

st.subheader("From OTHER PAY/DEDUCTIONS")
st.table({
    "PDF Line Item": [
        "PHYSICAL DAMAGE INS / PHYS DAM INS",
        "LIAB/CARGO INS / LIAB & CARGO INS / 50% LIAB & CARGO INS",
        "OCC ACC INS / OCC ACC INS & membership dues",
        "TRUCK NOTE / TRUCK NOTE PAYMENT / TRUCK PAYMENT",
        "KS CHILD SUPPORT / PA CHILD SUPPORT",
        "MM/DD/YY SCALES / TOLLS",
        "MM/DD/YY ADVANCE (EFS Mastercard)",
        "MM/DD/YY [fuel Gals]",
        "MM/DD/YY DEF (bulk)",
        "RESERVE 2025 Licensing",
        "RESERVE 2026 Licensing",
        "RESERVE Owner/Operator Escrow",
        "RESERVE [driver] - Maint Fund - Tr #[truck]",
        "RESERVE [driver] Loan",
    ],
    "Excel Description": [
        "Physical Damage Ins - Trucks",
        "Liability & Cargo Ins",
        "Occupational Accident Insurance",
        "Driver-specific N/R account",
        "Garnishments",
        "Scales",
        "EFS Mastercard Advance",
        "Added to Diesel total",
        "Added to Diesel total",
        "2025 Licensing Accrual",
        "2026 Licensing Accrual",
        "Owner/Operator Escrow",
        "Driver-specific Maint Fund account",
        "Loan",
    ],
    "QuickBooks Category": [
        "54580.1500 Truck Expense:Physical Damage Ins",
        "54520.1500 Truck Expense:Liability & Cargo Ins",
        "53710.3500 Occupational Accident Insurance",
        "19xxx.9000 Notes Receivable:N/R - Truck #",
        "21700.9000 Garnishments",
        "55500.1500 Truck Expense:Scales & Tolls",
        "54090.1500 Truck Expense:EFS Mastercard Transaction",
        "54100.1500 Truck Expense:Diesel",
        "54100.1500 Truck Expense:Diesel",
        "22397.9000 2025 Licensing Accrual",
        "22398.9000 2026 Licensing Accrual",
        "22418.9000 Owner/Operator Escrow",
        "22xxx.9000 Dr Maint Funds - Tr #",
        "10051.9000 Truck Expense:Comdata Comcheck",
    ]
})

st.subheader("From RESERVES")
st.table({
    "Reserve Type": [
        "Escrow — dated transaction",
        "Licensing 2025 — dated transaction",
        "Licensing 2026 — dated transaction",
        "Maintenance — dated transaction",
        "Loan — dated transaction",
        "Unknown — dated transaction",
    ],
    "QuickBooks Category": [
        "22418.9000 Owner/Operator Escrow",
        "22397.9000 2025 Licensing Accrual",
        "22398.9000 2026 Licensing Accrual",
        "Driver-specific Maint Fund account",
        "10051.9000 Truck Expense:Comdata Comcheck",
        "RESERVE UNKNOWN — must be filled in manually",
    ]
})

st.divider()

# ── Unmapped lines ────────────────────────────────────────────────────────────

st.header("Unmapped Lines")
st.markdown("""
Any line item the script could not map is still written to the Excel file with:

- The original description from the PDF in the **Description** column
- The dollar amount in the **Unit Price** column
- A placeholder in the **Category** column (`OTHER PAY/DEDUCTIONS` or `RESERVE UNKNOWN`)

These rows need to have their Category filled in manually before importing into QuickBooks. 
The processing summary in the app shows how many unmapped lines exist per driver.
""")

st.divider()

# ── Unmapped account warning ──────────────────────────────────────────────────

st.header("Unmapped Account Warning")
st.markdown("""
If a driver has a **Truck Note** or **Maintenance Fund** deduction but their account is 
not configured in the `DRIVERS` dictionary, the app shows an ⚠️ warning after processing 
and the Excel row will contain `UNMAPPED` in the Category column.

To fix this permanently, the driver's Notes Receivable or Maintenance Fund account needs 
to be added to the `DRIVERS` dictionary in `process_settlements.py`.
""")

st.divider()

# ── Excel output structure ────────────────────────────────────────────────────

st.header("Excel Output Structure")
st.markdown("""
Each row represents one line item on a driver's bill. Rows sharing the same 
Invoice/Bill Number belong to the same bill in QuickBooks.

| Column | Value |
|---|---|
| Post? | Always "No" |
| Invoice/Bill Date | Settlement date |
| Due Date | Settlement date + 1 day |
| Invoice / Bill Number | AP invoice number (e.g. ROLBRE/250602) |
| Transaction Type | Always "Bill" |
| Vendor | Driver full name |
| Currency Code | Always "USD" |
| Description | Line item description |
| Unit Price | Dollar amount |
| Category | QuickBooks account |
| Class | Driver's permanent truck/class number |

One blank row is added after each driver's entries, pre-filled with the driver's bill 
information, for any additional manual entries needed.
""")

st.divider()

# ── Files ─────────────────────────────────────────────────────────────────────

st.header("Files")
st.markdown("""
All files must stay in the same folder:

| File | Purpose |
|---|---|
| `process_settlements.py` | Core script — PDF parsing, account mappings, Excel generation |
| `App.py` | Streamlit UI — the browser interface |
| `pages/Documentation.py` | This documentation page |
| `launch_mac.command` | Launcher for Mac — double-click to start |
| `launch_windows.bat` | Launcher for Windows — double-click to start |
""")

st.divider()
st.caption("Jim Daws Trucking LLC — Settlement Processor — Internal Use Only")