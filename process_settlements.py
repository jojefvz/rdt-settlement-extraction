#!/usr/bin/env python3
"""
Jim Daws Trucking - Driver Settlement PDF to QuickBooks Excel
Usage: python process_settlements.py <settlement.pdf> [output.xlsx]

PARSING FLOW:
  1. OTHER PAY/DEDUCTIONS section
  2. TOTAL SETTLEMENT section
  3. RESERVES section
"""

import sys
import re
from datetime import datetime, timedelta
import pdfplumber
import openpyxl

# ── Account mappings ──────────────────────────────────────────────────────────

MAINT_FUND_ACCOUNTS = {
    "075C": ("Dr Maint Funds - Tr #79C - Burt",           "22467.3500 Dr Maint Funds - Tr #79C - Burt"),
    "126E": ("Dr Maint Funds - Tr #126E - Swett",         "22496.9000 Dr Maint Funds - Tr #126E - Swett"),
    "084C": ("Dr Maint Funds - Tr #84C - Saliba",         "22552.9000 Dr Maint Funds - Tr #84C - Saliba"),
    "096C": ("Dr Maint Funds - Tr #96C - May",            "22553.9000 Dr Maint Funds - Tr #96C - May"),
    "181A": ("Dr Maint Funds - Tr # 181A - Hill",         "22561.9000 Dr Maint Funds - Tr # 181A - Hill"),
    "182":  ("Dr Maint Funds - Tr # 182 - Miller",        "22562.9000 Dr Maint Funds - Tr # 182 - Miller"),
    "105C": ("Dr Maint Funds - Tr # 105C - Pooler",       "22564.9000 Dr Maint Funds - Tr # 105C - Pooler"),
    "184":  ("Dr Maint Funds - Tr # 184 - T Hall",        "22565.9000 Dr Maint Funds - Tr # 184 - T Hall"),
    "184A": ("Dr Maint Funds - Tr # 184 - T Hall",        "22565.9000 Dr Maint Funds - Tr # 184 - T Hall"),
    "183A": ("Dr Maint Funds - Tr # 183A - Parry",        "22566.9000 Dr Maint Funds - Tr # 183A - Parry"),
    "187":  ("Dr Maint Funds - Tr # 187 - M Buttjer",     "22568.9000 Dr Maint Funds - Tr # 187 - M Buttjer"),
    "071C": ("Dr Maint Funds - Tr # 71C -R Hemenway",     "22757.9000 Dr Maint Funds - Tr # 71C -R Hemenway"),
    "094C": ("Dr Maint Funds - Tr # 94C -Wiltrout",       "22760.9000 Dr Maint Funds - Tr # 94C -Wiltrout"),
    "095C": ("Dr Maint Funds - Tr # 95C -Williams",       "22763.9000 Dr Maint Funds - Tr # 95C -Williams"),
    "035C": ("Dr Maint Funds - Tr # 35C - Laws, M",       "22768.9000 Dr Maint Funds - Tr # 35C - Laws, M"),
    "190":  ("Dr Maint Funds - Tr # 190 - Green",         "22770.9000 Dr Maint Funds - Tr # 190 - Green"),
    "192":  ("Dr Maint Funds - Tr # 192 - McLeod",        "22771.9000 Dr Maint Funds - Tr # 192 - McLeod"),
    "017C": ("Dr Maint Fund - Tr #17C - Roach",           "22774.9000 Dr Maint Fund - Tr #17C - Roach"),
    "073":  ("Dr Maint Funds - Tr #73 - Evans",           "22775.9000 Dr Maint Funds - Tr #73 - Evans"),
    "110B": ("Dr Maint Funds - Tr #110B - Spell",         "22776.9000 Dr Maint Funds - Tr #110B - Spell"),
    "124B": ("Dr Maint Funds - Tr #110B - Spell",         "22776.9000 Dr Maint Funds - Tr #110B - Spell"),
    "047":  ("Dr Maint Funds - Tr# 47 Schulte",           "22777.9000 Dr Maint Funds - Tr# 47 Schulte"),
    "047G": ("Dr Maint Funds - Tr# 47 Schulte",           "22777.9000 Dr Maint Funds - Tr# 47 Schulte"),
    "195":  ("Dr Maint Funds - Tr# 195 Sourignavong",     "22779.9000 Dr Maint Funds - Tr# 195 Sourignavong"),
    "195C": ("Dr Maint Funds - Tr# 195 Sourignavong",     "22779.9000 Dr Maint Funds - Tr# 195 Sourignavong"),
    "194":  ("Dr Maint Funds - Tr# 194 N Lee",            "22780.9000 Dr Maint Funds - Tr# 194 N Lee"),
    "194C": ("Dr Maint Funds - Tr# 194 N Lee",            "22780.9000 Dr Maint Funds - Tr# 194 N Lee"),
    "083E": ("Dr Maint Funds - Tr# 083E Godsey",          "22781.9000 Dr Maint Funds - Tr# 083E Godsey"),
}

NOTES_RECEIVABLE = {
    "105C": ("N/R - Truck # 105C (Pooler)",       "19881.9000 Notes Receivable:N/R - Truck # 105C (Pooler)"),
    "071C": ("N/R - Truck # 71C (Hemenway, R)",   "19886.9000 Notes Receivable:N/R - Truck # 71C (Hemenway, R)"),
    "094C": ("N/R - Truck # 94C (Wiltrout)",      "19888.9000 Notes Receivable:N/R - Truck # 94C (Wiltrout)"),
    "095C": ("N/R - Truck # 95C (Williams)",      "19891.9000 Notes Receivable:N/R - Truck # 95C (Williams)"),
    "035C": ("N/R - Truck # 35C (Laws)",          "19893.9000 Notes Receivable:N/R - Truck # 35C (Laws)"),
    "017C": ("N/R - Truck # 17C (Roach)",         "19896.9000 Notes Receivable:N/R - Truck # 17C (Roach)"),
    "073":  ("N/R - Truck #73 (Evans)",           "19897.9000 Notes Receivable:N/R - Truck #73 (Evans)"),
    "110B": ("N/R - Truck #110B (Spell)",         "19898.9000 Notes Receivable:N/R - Truck #110B (Spell)"),
    "124B": ("N/R - Truck #110B (Spell)",         "19898.9000 Notes Receivable:N/R - Truck #110B (Spell)"),
    "083E": ("N/R - Truck #083E (Godsey)",        "19899.9000 Notes Receivable:N/R - Truck #083E (Godsey)"),
    "075C": ("N/R - Truck #75C (Burt)",           "19900.9000 Notes Receivable:N/R - Truck #75C (Burt)"),
}

DRIVER_NAMES = {
    "ROLBRE": "Rollin Brenneman",
    "ERIBUR": "Eric Burt",
    "MICBUT": "Michael Buttjer",
    "STEGRE": "Steve Green",
    "WILHAL": "William Haley",
    "ANTHAL": "Anthony Hall",
    "ROBHEM": "Robert Hemenway",
    "HEMFU":  "Robert Hemenway",
    "MICHIL": "Michael Hill",
    "MATLAW": "Matthew Laws",
    "NICLEE": "Nicolle Lee",
    "JOSMAY": "Josh May",
    "CURMCL": "Curtis McLeod",
    "KENMIL": "Kendall Miller",
    "JOSPOO": "Joshua Pooler",
    "ROAMAT": "Matthew Roach",
    "JONSAL": "Jonathan Saliba",
    "SCHCON": "Conrad Schulte",
    "CONSCH": "Conrad Schulte",
    "LADSOU": "Laddavanh Sourignavong",
    "SPEDAV": "David Spell",
    "TIMSWE": "Tim Swett",
    "SCOVAN": "Scott VanBuskirk",
    "OTHWIL": "Otho Williams",
    "GEOWIL": "George Wiltrout",
    "GODLEE": "Lee Godsey",
    "DARPAR": "Darl Parry",
}

# ── Helpers ───────────────────────────────────────────────────────────────────

def parse_amount(text):
    text = re.sub(r'[,$]', '', text).strip()
    try:
        return float(text)
    except ValueError:
        return None

def new_settlement(settlement_date, driver_code, driver_name):
    return {
        'settlement_date': settlement_date,
        'due_date': settlement_date + timedelta(days=1),
        'driver_code': driver_code,
        'driver_name': driver_name,
        'truck': None,
        'invoice_number': None,
        'gross_pay': None,
        'total_trip_expenses': None,
        'phys_dam_ins': None,
        'liab_cargo_ins': None,
        'occ_acc_ins': None,
        'truck_note': None,
        'maint_fund': None,
        'licensing_2025': None,
        'licensing_2026': None,
        'owner_operator_escrow': None,
        'garnishments': [],
        'scales': [],
        'skipped_other_pay': [],
        'skipped_total_settlement': [],
        'skipped_reserves': [],
    }

# ── PDF parsing ───────────────────────────────────────────────────────────────

SEC_NONE             = 0
SEC_OTHER_PAY        = 1
SEC_TOTAL_SETTLEMENT = 2
SEC_RESERVES         = 3

def extract_settlements(pdf_path):
    settlements = []
    current = None
    section = SEC_NONE
    settlement_date = None

    all_lines = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                all_lines.extend(text.split('\n'))

    for line in all_lines:
        line = line.strip()
        if not line:
            continue

        # ── Settlement date from page header ──────────────────────────────────
        m = re.search(r'FINAL TRUCKING SETTLEMENTS as of (\d{2}/\d{2}/\d{4})', line)
        if m:
            settlement_date = datetime.strptime(m.group(1), '%m/%d/%Y')
            continue

        # ── New driver ─────────────────────────────────────────────────────────
        m = re.search(r'Payables:\s+.+?\s+\(([A-Z]+)\)', line)
        if m:
            code = m.group(1)
            if current is None or code != current['driver_code']:
                if current:
                    settlements.append(current)
                name = DRIVER_NAMES.get(code, code)
                current = new_settlement(settlement_date, code, name)
                section = SEC_NONE
            continue

        if current is None:
            continue

        # ── Truck number (first occurrence wins) ───────────────────────────────
        if current['truck'] is None:
            m = re.search(r'TRK:(\S+)', line)
            if m:
                current['truck'] = m.group(1)

        # ── AP Invoice number ──────────────────────────────────────────────────
        m = re.search(r'NET PAY \(AP invo#\s+(.+?)\)', line)
        if m:
            current['invoice_number'] = m.group(1)

        # ── Section detection ──────────────────────────────────────────────────
        if re.match(r'OTHER PAY/DEDUCTIONS', line) and 'TOTAL' not in line:
            section = SEC_OTHER_PAY
            continue

        if re.match(r'TOTAL SETTLEMENT', line):
            section = SEC_TOTAL_SETTLEMENT
            continue

        if re.match(r'RESERVES$', line):
            section = SEC_RESERVES
            continue

        # ═════════════════════════════════════════════════════════════════════
        # SECTION 1 — OTHER PAY/DEDUCTIONS
        # ═════════════════════════════════════════════════════════════════════
        if section == SEC_OTHER_PAY:

            if re.match(r'TOTAL OTHER PAY/DEDUCTIONS', line):
                section = SEC_NONE
                continue

            has_amount = bool(re.search(r'[\$]?[\d,]+\.\d{2}\s*$', line))
            mapped = False

            # Physical Damage Ins
            if re.search(r'PHYS(ICAL)?\s*DAM(AGE)?\s*INS', line, re.I):
                m = re.search(r'(-[\$]?[\d,]+\.\d{2})\s*$', line)
                if m:
                    current['phys_dam_ins'] = parse_amount(m.group(1))
                mapped = True

            # Liability & Cargo Ins
            elif re.search(r'LIAB', line, re.I) and re.search(r'CARGO|INS', line, re.I):
                m = re.search(r'(-[\$]?[\d,]+\.\d{2})\s*$', line)
                if m:
                    current['liab_cargo_ins'] = parse_amount(m.group(1))
                mapped = True

            # OCC ACC INS
            elif re.search(r'OCC\s*ACC\s*INS', line, re.I):
                m = re.search(r'(-[\$]?[\d,]+\.\d{2})\s*$', line)
                if m:
                    current['occ_acc_ins'] = parse_amount(m.group(1))
                mapped = True

            # Truck Note / Payment
            elif re.search(r'TRUCK\s+(NOTE|PAYMENT)', line, re.I):
                m = re.search(r'(-[\$]?[\d,]+\.\d{2})\s*$', line)
                if m:
                    current['truck_note'] = parse_amount(m.group(1))
                mapped = True

            # Child support / garnishments
            elif re.search(r'CHILD\s+SUPPORT', line, re.I):
                m = re.search(r'(-[\$]?[\d,]+\.\d{2})\s*$', line)
                if m:
                    current['garnishments'].append(parse_amount(m.group(1)))
                mapped = True

            # Scales & Tolls
            elif re.match(r'^\d{2}/\d{2}/\d{2}', line) and re.search(r'SCALE|TOLL', line, re.I):
                m = re.search(r'(-[\$]?[\d,]+\.\d{2})\s*$', line)
                if m:
                    current['scales'].append(parse_amount(m.group(1)))
                mapped = True

            # RESERVE lines
            elif line.startswith('RESERVE'):
                m = re.search(r'(-[\$]?[\d,]+\.\d{2})\s*$', line)
                if m:
                    amt = parse_amount(m.group(1))
                    if re.search(r'2025\s*Licens', line, re.I):
                        current['licensing_2025'] = amt
                        mapped = True
                    elif re.search(r'2026\s*Licens', line, re.I):
                        current['licensing_2026'] = amt
                        mapped = True
                    elif re.search(r'Owner/Operator\s*Escrow', line, re.I):
                        current['owner_operator_escrow'] = amt
                        mapped = True
                    elif re.search(r'Maint', line, re.I):
                        current['maint_fund'] = amt
                        mapped = True

            if not mapped and has_amount:
                current['skipped_other_pay'].append(line)

        # ═════════════════════════════════════════════════════════════════════
        # SECTION 2 — TOTAL SETTLEMENT
        # ═════════════════════════════════════════════════════════════════════
        elif section == SEC_TOTAL_SETTLEMENT:

            matched = False

            if re.search(r'PERCENTAGE PAY', line, re.I):
                m = re.search(r'\$?([\d,]+\.\d{2})\s*$', line)
                if m:
                    val = float(m.group(1).replace(',', ''))
                    current['gross_pay'] = (current['gross_pay'] or 0) + val
                matched = True

            elif re.search(r'TOTAL TRIP EXPENSES', line, re.I):
                amounts = re.findall(r'-[\$]?([\d,]+\.\d{2})', line)
                if amounts:
                    current['total_trip_expenses'] = -float(amounts[-1].replace(',', ''))
                matched = True

            if not matched:
                has_amount = bool(re.search(r'[\$]?[\d,]+\.\d{2}\s*$', line))
                ignore = any(re.search(p, line, re.I) for p in [
                    r'^TOTAL', r'^YOUR TOTAL', r'^NET PAY', r'^TOTAL OTHER',
                    r'^Run Date', r'^Run Time', r'^Page No', r'^JIM DAWS',
                    r'^FINAL TRUCKING', r'^Payables:'
                ])
                if has_amount and not ignore:
                    current['skipped_total_settlement'].append(line)

        # ═════════════════════════════════════════════════════════════════════
        # SECTION 3 — RESERVES
        # ═════════════════════════════════════════════════════════════════════
        elif section == SEC_RESERVES:

            # End of reserves — next driver or end of file will close this
            # Ignore: "Addition to Reserve:" lines (already captured via RESERVE in OTHER PAY)
            if re.match(r'Addition to Reserve:', line, re.I):
                continue

            # Ignore: "End Reserve Balance" lines
            if re.match(r'End Reserve Balance', line, re.I):
                continue

            # Dated lines → skipped report
            if re.match(r'^\d{2}/\d{2}/\d{2}', line):
                has_amount = bool(re.search(r'[\$]?[\d,]+\.\d{2}\s*$', line))
                if has_amount:
                    current['skipped_reserves'].append(line)

    if current:
        settlements.append(current)

    return settlements

# ── Excel output ──────────────────────────────────────────────────────────────

HEADERS = [
    'Post?', 'Invoice/Bill Date', 'Due Date', 'Invoice / Bill Number',
    'Transaction Type', 'Customer', 'Vendor', 'Currency Code',
    'Product/Services', 'Description', 'Qty', 'Discount %',
    'Unit Price', 'Category', 'Location', 'Class'
]

def make_row(s, description, unit_price, category):
    return [
        'No',
        s['settlement_date'],
        s['due_date'],
        s['invoice_number'],
        'Bill',
        None,
        s['driver_name'],
        'USD',
        None,
        description,
        None,
        None,
        unit_price,
        category,
        None,
        s['truck'],
    ]

def settlement_to_rows(s):
    rows = []
    truck = s.get('truck', '') or ''

    # ── TOTAL SETTLEMENT section ───────────────────────────────────────────────
    if s['gross_pay'] is not None:
        rows.append(make_row(s, 'Settlements', s['gross_pay'],
                             '53400.3500 Driver Pay:I/C Settlements'))

    if s['total_trip_expenses'] is not None and s['total_trip_expenses'] != 0:
        rows.append(make_row(s, 'Driver Reimbursements', s['total_trip_expenses'],
                             '54100.1500 Truck Expense:Diesel'))

    # ── OTHER PAY/DEDUCTIONS section ──────────────────────────────────────────
    if s['phys_dam_ins'] is not None:
        rows.append(make_row(s, 'Physical Damage Ins - Trucks', s['phys_dam_ins'],
                             '54580.1500 Truck Expense:Physical Damage Ins'))

    if s['liab_cargo_ins'] is not None:
        rows.append(make_row(s, 'Liability & Cargo Ins', s['liab_cargo_ins'],
                             '54520.1500 Truck Expense:Liability & Cargo Ins'))

    if s['occ_acc_ins'] is not None:
        rows.append(make_row(s, 'Occupational Accident Insurance', s['occ_acc_ins'],
                             '53710.3500 Occupational Accident Insurance'))

    if s['truck_note'] is not None:
        nr = NOTES_RECEIVABLE.get(truck)
        if nr:
            rows.append(make_row(s, nr[0], s['truck_note'], nr[1]))
        else:
            rows.append(make_row(s, f'N/R - Truck # {truck}',
                                 s['truck_note'], f'UNMAPPED - Notes Receivable Truck {truck}'))

    if s['maint_fund'] is not None:
        mf = MAINT_FUND_ACCOUNTS.get(truck)
        if mf:
            rows.append(make_row(s, mf[0], s['maint_fund'], mf[1]))
        else:
            rows.append(make_row(s, f'Dr Maint Funds - Tr # {truck}',
                                 s['maint_fund'], f'UNMAPPED - Maint Fund Truck {truck}'))

    if s['licensing_2025'] is not None:
        rows.append(make_row(s, '2025 Licensing Accrual', s['licensing_2025'],
                             '22397.9000 2025 Licensing Accrual'))

    if s['licensing_2026'] is not None:
        rows.append(make_row(s, '2026 Licensing Accrual', s['licensing_2026'],
                             '22398.9000 2026 Licensing Accrual'))

    if s['owner_operator_escrow'] is not None:
        rows.append(make_row(s, 'Owner/Operator Escrow', s['owner_operator_escrow'],
                             '22418.9000 Owner/Operator Escrow'))

    for amt in s['garnishments']:
        rows.append(make_row(s, 'Garnishments', amt, '21700.9000 Garnishments'))

    for amt in s['scales']:
        rows.append(make_row(s, 'Scales', amt, '55500.1500 Truck Expense:Scales & Tolls'))

    rows.append(make_row(s, '', '', ''))
    rows.append(make_row(s, '', '', ''))

    return rows

def write_excel(settlements, output_path):
    wb = openpyxl.Workbook()

    # ── Sheet 1: QB Import data ───────────────────────────────────────────────
    ws = wb.active
    ws.title = 'Settlements'
    ws.append(HEADERS)

    for s in settlements:
        for row in settlement_to_rows(s):
            ws.append(row)

    for row in ws.iter_rows(min_row=2):
        for cell in row:
            if isinstance(cell.value, datetime):
                cell.number_format = 'MM/DD/YYYY'

    # ── Sheet 2: Skipped lines report ─────────────────────────────────────────
    ws2 = wb.create_sheet(title='Skipped Lines')
    ws2.append(['Driver', 'Invoice #', 'Section', 'Skipped Line'])

    any_skipped = False
    for s in settlements:
        for line in s['skipped_other_pay']:
            ws2.append([s['driver_name'], s['invoice_number'],
                        'OTHER PAY/DEDUCTIONS', line])
            any_skipped = True
        for line in s['skipped_total_settlement']:
            ws2.append([s['driver_name'], s['invoice_number'],
                        'TOTAL SETTLEMENT', line])
            any_skipped = True
        for line in s['skipped_reserves']:
            ws2.append([s['driver_name'], s['invoice_number'],
                        'RESERVES', line])
            any_skipped = True

    if not any_skipped:
        ws2.append(['No skipped lines — all items mapped successfully.'])

    ws2.column_dimensions['A'].width = 28
    ws2.column_dimensions['B'].width = 20
    ws2.column_dimensions['C'].width = 24
    ws2.column_dimensions['D'].width = 80

    wb.save(output_path)

    total_rows = sum(len(settlement_to_rows(s)) for s in settlements)
    total_skipped = sum(
        len(s['skipped_other_pay']) + len(s['skipped_total_settlement']) + len(s['skipped_reserves'])
        for s in settlements
    )
    print(f"Saved: {output_path}")
    print(f"  {len(settlements)} driver settlements → {total_rows} QB rows")
    print(f"  {total_skipped} skipped lines logged to 'Skipped Lines' tab")

# ── Main ──────────────────────────────────────────────────────────────────────

if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("Usage: python process_settlements.py <settlement.pdf> [output.xlsx]")
        sys.exit(1)

    pdf_path = sys.argv[1]
    output_path = sys.argv[2] if len(sys.argv) > 2 else pdf_path.replace('.pdf', '_QB.xlsx')

    print(f"Processing: {pdf_path}")
    settlements = extract_settlements(pdf_path)
    print(f"Found: {len(settlements)} driver settlements")

    for s in settlements:
        flag = " ⚠ UNMAPPED" if (
            (s['truck_note'] is not None and s['truck'] not in NOTES_RECEIVABLE) or
            (s['maint_fund'] is not None and s['truck'] not in MAINT_FUND_ACCOUNTS)
        ) else ""
        print(f"  {s['driver_name']:<30} TRK:{s['truck']:<6} INV:{s['invoice_number']:<16} "
              f"Gross:{s['gross_pay']}{flag}")

    write_excel(settlements, output_path)
