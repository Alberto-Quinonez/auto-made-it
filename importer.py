import csv
import openpyxl


def import_csv_to_excel(csv_path: str, excel_path: str) -> str:
    rows = _parse_csv(csv_path)

    wb = openpyxl.load_workbook(excel_path)
    ws = wb["BMO - CASH Plex"]

    last_row = _find_last_data_row(ws)
    existing = _build_duplicate_set(ws, last_row)
    last_balance = _get_last_balance(ws, last_row)

    added = 0
    skipped = 0
    for card, tx_type, date, amount, description in rows:
        key = (date, amount, description)
        if key in existing:
            skipped += 1
            continue

        amt = float(amount)
        dr = amount if tx_type == "CREDIT" else ""
        cr = amount if tx_type == "DEBIT" else ""
        last_balance = round(last_balance + amt, 2)

        ws.append([card, tx_type, date, amount, description, dr, cr, last_balance, "", ""])
        existing.add(key)
        added += 1

    wb.save(excel_path)
    return f"Done — {added} row(s) added, {skipped} skipped."


def _parse_csv(csv_path: str) -> list:
    rows = []
    header_found = False
    with open(csv_path, newline="", encoding="utf-8-sig") as f:
        reader = csv.reader(f)
        for row in reader:
            if not header_found:
                if len(row) > 1 and row[1].strip() == "Type de transaction":
                    header_found = True
                continue
            if len(row) < 5 or not row[2].strip():
                continue
            card = row[0].strip()
            tx_type = row[1].strip()
            date = row[2].strip()
            amount = row[3].strip()
            description = row[4].strip()
            rows.append((card, tx_type, date, amount, description))
    return rows


def _find_last_data_row(ws) -> int:
    for row in range(ws.max_row, 0, -1):
        if ws.cell(row=row, column=1).value or ws.cell(row=row, column=4).value:
            return row
    return 1


def _build_duplicate_set(ws, last_row: int) -> set:
    seen = set()
    for row in range(1, last_row + 1):
        date = ws.cell(row=row, column=3).value
        amount = ws.cell(row=row, column=4).value
        desc = ws.cell(row=row, column=5).value
        if date and amount is not None and desc:
            seen.add((str(date).strip(), str(amount).strip(), str(desc).strip()))
    return seen


def _get_last_balance(ws, last_row: int) -> float:
    for row in range(last_row, 0, -1):
        val = ws.cell(row=row, column=8).value
        if val is not None:
            try:
                return float(val)
            except (ValueError, TypeError):
                continue
    return 0.0
