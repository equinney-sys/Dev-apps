#!/usr/bin/env python3
import csv
import json
import sys
import os

CSV_PATH = sys.argv[1] if len(sys.argv) > 1 else os.path.expanduser('~/Downloads/Quinney Bills - 2026.csv')
OUT_FILE = 'debts.json'
MONTHS_TO_INCLUDE = ['January','February','March','April','May','June']


def parse_csv(path):
    with open(path, newline='', encoding='utf-8') as f:
        reader = list(csv.reader(f))

    if not reader:
        raise SystemExit('CSV empty')

    # Find the row that contains month names (like 'January') and the following row that contains 'Payment'/'Remainder'
    month_row = None
    label_row = None
    for i in range(len(reader)-1):
        if any(cell and cell.strip() in MONTHS_TO_INCLUDE for cell in reader[i]):
            # candidate month row
            if any(cell and cell.strip() == 'Payment' for cell in reader[i+1]):
                month_row = reader[i]
                label_row = reader[i+1]
                header_index = i+2
                break

    if month_row is None or label_row is None:
        raise SystemExit('Could not find month/label header rows')

    # Determine which columns are "Remainder" for months Jan-Jun.
    # The CSV uses a sparse month row (months repeated with blank separators),
    # so we find every column whose label is 'Remainder' and then look left
    # to find the nearest non-empty month header for that column.
    remainder_cols = []  # list of (col_idx, month)
    for col_idx, lcell in enumerate(label_row):
        label = (lcell or '').strip()
        if label != 'Remainder':
            continue
        # search left for nearest month name
        month = None
        for j in range(col_idx, -1, -1):
            candidate = (month_row[j] or '').strip() if j < len(month_row) else ''
            if candidate:
                month = candidate
                break
        if month in MONTHS_TO_INCLUDE:
            remainder_cols.append((col_idx, month))

    if not remainder_cols:
        raise SystemExit('No remainder columns found for Jan-Jun')

    # Parse rows after header_index
    accounts = []
    for row in reader[header_index:]:
        if not any(cell.strip() for cell in row if cell):
            # skip empty rows
            continue
        name = (row[0] or '').strip()
        if not name:
            continue

        # For each remainder column (in chronological order), pick the last non-empty value
        latest_val = None
        for col_idx, month in remainder_cols:
            val = ''
            if col_idx < len(row):
                val = (row[col_idx] or '').strip()
            if val:
                # try parse number, may contain commas
                try:
                    num = float(val.replace(',',''))
                except ValueError:
                    # sometimes the "Remainder" field may be in the Payment column or have extra text; try stripping non-numeric
                    cleaned = ''.join(ch for ch in val if (ch.isdigit() or ch in '.-'))
                    try:
                        num = float(cleaned) if cleaned else 0.0
                    except ValueError:
                        num = 0.0
                latest_val = num

        if latest_val is None:
            # No remainder found in Jan-Jun; skip
            continue

        accounts.append({'name': name, 'balance': round(latest_val, 2)})

    return accounts


if __name__ == '__main__':
    print('Reading', CSV_PATH)
    accounts = parse_csv(CSV_PATH)
    data = {'accounts': accounts}
    with open(OUT_FILE, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=2)
    print(f'Wrote {len(accounts)} accounts to {OUT_FILE}')
