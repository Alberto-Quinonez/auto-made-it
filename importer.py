import csv
from copy import copy

import openpyxl

from config import BankConfig


def import_csv_to_excel(csv_path: str, excel_path: str, bank: BankConfig) -> str:
    csv_rows = _parse_csv(csv_path, bank)

    wb_ro = openpyxl.load_workbook(excel_path, data_only=True)
    ws_ro = wb_ro[bank.excel_sheet]
    insert_row = _find_insert_row(ws_ro, bank)
    existing = _build_duplicate_set(ws_ro, insert_row, bank)
    last_balance = _get_last_balance(ws_ro, insert_row - 1, bank)
    wb_ro.close()

    new_rows, skipped = [], 0
    for card, tx_type, date_str, amount_str, description in csv_rows:
        try:
            key = (int(date_str), float(amount_str), description)
        except ValueError:
            continue
        if key in existing:
            skipped += 1
        else:
            new_rows.append((card, tx_type, date_str, amount_str, description))

    wb = openpyxl.load_workbook(excel_path)
    ws = wb[bank.excel_sheet]

    if new_rows:
        insert_row = _find_insert_row(ws, bank)
        ws.insert_rows(insert_row, amount=len(new_rows))

        template_row = insert_row - 1
        max_col = ws.max_column

        for i, (card, tx_type, date_str, amount_str, description) in enumerate(new_rows):
            r = insert_row + i
            amt = float(amount_str)
            last_balance = round(last_balance + amt, 2)

            # Copy cell formatting from the last existing data row
            for col in range(1, max_col + 1):
                src = ws.cell(row=template_row, column=col)
                dst = ws.cell(row=r, column=col)
                if src.has_style:
                    dst.font = copy(src.font)
                    dst.fill = copy(src.fill)
                    dst.border = copy(src.border)
                    dst.alignment = copy(src.alignment)
                    dst.number_format = src.number_format

            ws.cell(row=r, column=bank.excel_card).value = card
            ws.cell(row=r, column=bank.excel_type).value = tx_type
            ws.cell(row=r, column=bank.excel_date).value = int(date_str)
            ws.cell(row=r, column=bank.excel_amount).value = amt
            ws.cell(row=r, column=bank.excel_desc).value = description
            ws.cell(row=r, column=bank.excel_dr).value = abs(amt) if tx_type == "CREDIT" else None
            ws.cell(row=r, column=bank.excel_cr).value = amt if tx_type == "DEBIT" else None
            ws.cell(row=r, column=bank.excel_balance).value = last_balance

    wb.save(excel_path)
    return f"Done — {len(new_rows)} row(s) added, {skipped} skipped."


def _parse_csv(csv_path: str, bank: BankConfig) -> list:
    rows = []
    header_found = False
    with open(csv_path, newline="", encoding="utf-8-sig") as f:
        reader = csv.reader(f)
        for row in reader:
            if not header_found:
                if len(row) > bank.header_col_index and row[bank.header_col_index].strip() == bank.header_col_value:
                    header_found = True
                continue
            if len(row) < 5 or not row[bank.csv_date].strip():
                continue
            rows.append((
                row[bank.csv_card].strip(),
                row[bank.csv_type].strip(),
                row[bank.csv_date].strip(),
                row[bank.csv_amount].strip(),
                row[bank.csv_desc].strip(),
            ))
    return rows


def _find_insert_row(ws, bank: BankConfig) -> int:
    for row in range(bank.excel_data_start_row, ws.max_row + 2):
        if ws.cell(row=row, column=bank.excel_card).value is None:
            return row
    return bank.excel_data_start_row


def _build_duplicate_set(ws, insert_row: int, bank: BankConfig) -> set:
    seen = set()
    for row in range(bank.excel_data_start_row, insert_row):
        date = ws.cell(row=row, column=bank.excel_date).value
        amt = ws.cell(row=row, column=bank.excel_amount).value
        desc = ws.cell(row=row, column=bank.excel_desc).value
        if date is not None and amt is not None and desc:
            try:
                seen.add((int(date), float(amt), str(desc).strip()))
            except (ValueError, TypeError):
                pass
    return seen


def _get_last_balance(ws, last_row: int, bank: BankConfig) -> float:
    # Try cached formula values first (works if user saved the file from Excel)
    for row in range(last_row, bank.excel_data_start_row - 1, -1):
        val = ws.cell(row=row, column=bank.excel_balance).value
        if val is not None:
            try:
                return float(val)
            except (ValueError, TypeError):
                continue

    # Formulas not cached — find opening balance then sum all DR/CR
    opening = 0.0
    for row in range(bank.excel_data_start_row - 1, 0, -1):
        val = ws.cell(row=row, column=bank.excel_balance).value
        if val is not None:
            try:
                opening = float(val)
                break
            except (ValueError, TypeError):
                continue

    balance = opening
    for row in range(bank.excel_data_start_row, last_row + 1):
        for col in (bank.excel_dr, bank.excel_cr):
            val = ws.cell(row=row, column=col).value
            if val is not None:
                try:
                    balance += float(val)
                except (ValueError, TypeError):
                    pass

    return round(balance, 2)
