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
to the correct accounts. The output Excel file has two tabs:

- **Settlements** — the QuickBooks-ready data
- **Skipped Lines** — any line items that could not be automatically mapped
""")

st.divider()

# ── How to use ────────────────────────────────────────────────────────────────

st.header("How to Use")
st.markdown("""
1. Launch the app by double-clicking the launcher file
2. Your browser opens automatically at `http://localhost:8501`
3. Click **Processor** in the left sidebar
4. Click **Browse files** and select the weekly settlement PDF
5. Click **Process Settlements**
6. Review the processing summary
7. Click **Download Excel File**
8. Open the Excel file, review the **Skipped Lines** tab, and manually fill in any missing entries using the two blank rows provided after each driver
9. Import the **Settlements** tab into QuickBooks
""")

st.divider()

# ── PDF structure ─────────────────────────────────────────────────────────────

st.header("How the PDF is Read")
st.markdown("""
The script reads the PDF line by line and works through three sections for each driver 
in order:
""")

st.subheader("1. OTHER PAY/DEDUCTIONS")
st.markdown("""
This section contains insurance deductions, garnishments, reserve contributions, and 
scales charges. The script scans each line and maps it if it matches a known pattern. 
Any line with a dollar amount that is not recognized is logged to the **Skipped Lines** report.
""")

st.subheader("2. TOTAL SETTLEMENT")
st.markdown("""
This section contains the driver's gross pay and total trip expenses. These are the two 
most important numbers — gross pay becomes the **Settlements** line item and trip expenses 
become the **Driver Reimbursements** line item in QuickBooks.
""")

st.subheader("3. RESERVES")
st.markdown("""
This section shows running balances for each reserve account. The script ignores balance 
lines and "Addition to Reserve" lines since those amounts are already captured in the 
OTHER PAY/DEDUCTIONS section. Any dated lines (maintenance charges, repairs, parts) are 
logged to the **Skipped Lines** report.
""")

st.divider()

# ── Mappings ──────────────────────────────────────────────────────────────────

st.header("Current Mappings")
st.markdown("These are the line items the script currently maps automatically:")

st.subheader("From TOTAL SETTLEMENT")
st.table({
    "PDF Line Item": [
        "PERCENTAGE PAY(...)",
        "TOTAL TRIP EXPENSES",
    ],
    "Excel Description": [
        "Settlements",
        "Driver Reimbursements",
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
        "RESERVE 2025 Licensing",
        "RESERVE 2026 Licensing",
        "RESERVE Owner/Operator Escrow",
        "RESERVE [driver] - Maint Fund - Tr #[truck]",
    ],
    "Excel Description": [
        "Physical Damage Ins - Trucks",
        "Liability & Cargo Ins",
        "Occupational Accident Insurance",
        "N/R - Truck # [number]",
        "Garnishments",
        "Scales",
        "2025 Licensing Accrual",
        "2026 Licensing Accrual",
        "Owner/Operator Escrow",
        "Dr Maint Funds - Tr # [number]",
    ],
    "QuickBooks Category": [
        "54580.1500 Truck Expense:Physical Damage Ins",
        "54520.1500 Truck Expense:Liability & Cargo Ins",
        "53710.3500 Occupational Accident Insurance",
        "19xxx.9000 Notes Receivable:N/R - Truck #",
        "21700.9000 Garnishments",
        "55500.1500 Truck Expense:Scales & Tolls",
        "22397.9000 2025 Licensing Accrual",
        "22398.9000 2026 Licensing Accrual",
        "22418.9000 Owner/Operator Escrow",
        "22xxx.9000 Dr Maint Funds - Tr #",
    ]
})

st.divider()

# ── Skipped lines ─────────────────────────────────────────────────────────────

st.header("Skipped Lines Tab")
st.markdown("""
Any line item with a dollar amount that the script could not map automatically is written 
to the **Skipped Lines** tab in the Excel output. Each row shows:

- **Driver** — the driver's full name
- **Invoice #** — the AP invoice number for that settlement
- **Section** — which section of the PDF the line came from (OTHER PAY/DEDUCTIONS, TOTAL SETTLEMENT, or RESERVES)
- **Skipped Line** — the exact text from the PDF

Common items found here include cash advances, driver loans, truck washes, and 
maintenance charges from the RESERVES section. These need to be entered manually 
using the two blank rows provided after each driver's mapped lines in the Settlements tab.
""")

st.divider()

# ── Unmapped accounts ─────────────────────────────────────────────────────────

st.header("Unmapped Account Warning")
st.markdown("""
If a driver has a **Truck Note** or **Maintenance Fund** deduction but their truck number 
is not in the script's account lookup tables, the app will show an ⚠️ warning after 
processing. The Excel row for that line item will contain `UNMAPPED` in the Category 
column so it is easy to find and correct before importing into QuickBooks.

To fix a permanently unmapped account, the truck number and its corresponding QuickBooks 
account need to be added to the script by a developer.
""")

st.divider()

# ── Excel output structure ────────────────────────────────────────────────────

st.header("Excel Output Structure")
st.markdown("""
Each row in the **Settlements** tab represents one line item on a driver's bill. 
Rows sharing the same Invoice/Bill Number belong to the same bill in QuickBooks. 
The columns are:

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
| Class | Truck number |

Two blank rows are added after each driver's entries. These are pre-filled with the 
driver's bill information so you only need to add the Description, Unit Price, and 
Category for any manually entered items.
""")

st.divider()

# ── Files ─────────────────────────────────────────────────────────────────────

st.header("Files")
st.markdown("""
The app consists of three files that must always stay in the same folder:

| File | Purpose |
|---|---|
| `process_settlements.py` | Core script — handles all PDF parsing and Excel generation |
| `app.py` | Streamlit UI — the browser interface |
| `launch_mac.command` / `launch_windows.bat` | Launcher — double-click to start the app |
""")

st.divider()
st.caption("Jim Daws Trucking LLC — Settlement Processor — Internal Use Only")
