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
directly in the Excel sheet with placeholder descriptions so they can be categorized manually 
before importing.
""")

st.divider()

# ── How to use ────────────────────────────────────────────────────────────────

st.header("How to Use")
st.markdown("""
1. Click **Browse files** and select the weekly settlement PDF
2. Click **Process Settlements**
3. Review the processing summary and driver breakdown
4. Click **Download Excel File**
5. Open the Excel file and fill in the Category column for any unmapped/uncertain rows
6. Import into QuickBooks
""")

st.divider()

# ── PDF structure ─────────────────────────────────────────────────────────────

st.header("How the PDF is Read")
st.markdown("""
The script reads the PDF line by line and for each driver extracts transactions 
from three sections: OTHER PAY/DEDUCTIONS, TOTAL SETTLEMENT, and RESERVES. Each relevant transaction line 
that **must** be accounted for, will be extracted. A driver must exist in the `DRIVERS` dictionary 
in the script in order to be processed. Unknown driver codes are skipped with a warning printed to the terminal.
""")

st.subheader("1. OTHER PAY/DEDUCTIONS")
st.markdown("""
Contains insurance deductions, garnishments, reserve contributions, advances, fuel charges, 
and scales. The script scans each line with a dollar amount and maps the line if it matches 1 to 1 with an account in QuickBooks.
Lines that do not have simple mappings are STILL included in the Excel output with their original 
PDF line description and a category of `OTHER PAY/DEDUCTIONS` so they can be handled manually.
""")

st.subheader("2. TOTAL SETTLEMENT")
st.markdown("""
Contains the driver's gross pay and total trip expenses. PERCENTAGE PAY becomes the **Settlements** 
line in Excel and TOTAL TRIP EXPENSE become the **Diesel** line. Any **fuel or DEF charges** found in 
OTHER PAY/DEDUCTIONS are **added** to the TOTAL TRIP EXPENSE dollar value.
""")

st.subheader("3. RESERVES")
st.markdown("""
Shows running balances for each reserve account — Escrow, Licensing, Maintenance, and Loan. 
Within each reserve, any transaction lines (maintenance charges, repairs, parts)
between the **initial balance** and the **end balance** are captured and routed to their reserve QB account.
The lines "Addition to Reserve" are not captured since those amounts are already extracted from the section OTHER PAY/DEDUCTIONS.
If an **unrecognized reserve** appears, the transactions that took place
are STILL included in the Excel output with a category of `RESERVE UNKNOWN`.
""")

st.divider()

# ── Mappings ──────────────────────────────────────────────────────────────────

st.header("Current Mappings")
st.markdown("These are the line items the script currently maps automatically:")

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
        "No Description, simply added to Diesel total",
        "No Description, simply added to Diesel total",
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

st.subheader("From TOTAL SETTLEMENT")
st.table({
    "PDF Line Item": [
        "PERCENTAGE PAY(...)",
        "TOTAL TRIP EXPENSES + fuel/DEF from OTHER PAY/DEDUCTIONS",
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

st.divider()

# ── Unmapped lines ────────────────────────────────────────────────────────────

st.header("Unmapped Lines")
st.markdown("""
Any line item the script could not map cleanly is still written to the Excel file with:
"""
)           

st.subheader("From OTHER PAY/DEDUCTIONS")
st.table({
    "PDF Line Item": [
        "PDF line the script cannot map cleanly",
    ],
    "Excel Description": [
        "PDF line description is used as placeholder",
    ],
    "QuickBooks Category": [
        "OTHER PAY/DEDUCTIONS used as placeholder",
    ]
})

st.subheader("From RESERVES")
st.table({
    "PDF Line Item": [
        "Transaction lines between Escrow initial balance & end balance, NOT including 'Addition to Reserve:' lines",
        "Transaction lines between Licensing initial balance & end balance, NOT including 'Addition to Reserve:' lines",
        "Transaction lines between Maintenance initial balance & end balance, NOT including 'Addition to Reserve:' lines",
        "Transaction lines between Loan initial balance & end balance, NOT including 'Addition to Reserve:' lines",
        "Transaction lines between the initial balance & end balance, NOT including 'Addition to Reserve:' lines",
    ],
    "Excel Description": [
        "PDF line description is used as placeholder",
        "[year] Licensing Accrual",
        "PDF line description is used as placeholder",
        "Driver-specific Maint Fund",
        "PDF line description is used as placeholder",
        "PDF line description is used as placeholder",
    ],
    "QuickBooks Category": [
        "22418.9000 Owner/Operator Escrow",
        "Driver-specific Maint Fund account",
        "2239x.9000 [year] Licensing Accrual",
        "10051.9000 Truck Expense:Comdata Comcheck",
        "RESERVE UNKNOWN used as placeholder",
    ]
})

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

# ── Unknown drivers warning ──────────────────────────────────────────────────

st.header("Unknown Drivers Warning")
st.markdown("""
Drivers not found in the script's driver dictionary will be SKIPPED entirely.
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
| Class | Driver's QuickBooks class number |

One blank row is added after each driver's entries, pre-filled with the driver's bill 
information, for any additional manual entries needed.
""")

st.divider()

st.caption("Jim Daws Trucking LLC — Settlement Processor — Internal Use Only")