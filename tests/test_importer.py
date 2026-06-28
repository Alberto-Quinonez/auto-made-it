"""Unit tests for importer.py helper functions."""
import pytest
from datetime import datetime
from openpyxl import Workbook

from importer import (
    _bmo_duplicate_set,
    _desj_duplicate_set,
    _find_insert_row,
    _get_balance_simple,
    _parse_bmo_csv,
    _parse_desj_csv,
    _to_float,
)
from config import BMO, DESJ


# ── helpers ───────────────────────────────────────────────────────────────────

def _write_csv(tmp_path, lines, encoding="latin-1", name="test.csv"):
    p = tmp_path / name
    p.write_text("\n".join(lines) + "\n", encoding=encoding)
    return str(p)


def _ws_with(data: dict) -> object:
    """Return a worksheet pre-populated with {(row, col): value} entries."""
    wb = Workbook()
    ws = wb.active
    for (r, c), v in data.items():
        ws.cell(row=r, column=c).value = v
    return ws


# ── _to_float ────────────────────────────────────────────────────────────────

@pytest.mark.parametrize("val,expected", [
    ("", None),
    ("   ", None),
    ("4532.00", 4532.0),
    ("4,531.26", 4531.26),
    ("-100.50", -100.5),
    ("0", 0.0),
    ("abc", None),
])
def test_to_float(val, expected):
    assert _to_float(val) == expected


# ── _parse_desj_csv ──────────────────────────────────────────────────────────

# Canonical test rows matching the actual bank CSV format
_DESJ_PCA_DEPOSIT = (
    "L'\xeele-des-Soeurs--Verdun,057992,PCA,2026/07/11,00001,"
    "Money transfer received from /BANK OF MONTREAL,,,4532.00,,,,,7699.95"
)
_DESJ_PCA_WITHDRAWAL = (
    "L'\xeele-des-Soeurs--Verdun,057992,PCA,2026/07/15,00002,"
    "Payment transfer /to LN 01,,4531.26,,,,,,3168.69"
)
# Description contains an unquoted comma — the bank sends it this way
_DESJ_LN1_UNQUOTED = (
    "L'\xeele-des-Soeurs--Verdun,057992,LN1,2026/07/15,00001,"
    "Automatic payment /from PCA:4,531.26$,,,,2326.48,,,2204.78,803280.34"
)
# Same row but with CSV quoting (also valid input if the bank ever quotes it)
_DESJ_LN1_QUOTED = (
    '"L\'\xeele-des-Soeurs--Verdun",057992,LN1,2026/07/15,00001,'
    '"Automatic payment /from PCA:4,531.26$",,,,2326.48,,,2204.78,803280.34'
)


def test_desj_pca_deposit_parsed(tmp_path):
    path = _write_csv(tmp_path, [_DESJ_PCA_DEPOSIT])
    rows = _parse_desj_csv(path, DESJ)

    assert len(rows) == 1
    account_type, date, code, desc, withdrawal, deposit, interest, capital = rows[0]
    assert account_type == "PCA"
    assert date == datetime(2026, 7, 11)
    assert deposit == 4532.0
    assert withdrawal is None
    assert interest is None
    assert capital is None


def test_desj_pca_withdrawal_parsed(tmp_path):
    path = _write_csv(tmp_path, [_DESJ_PCA_WITHDRAWAL])
    rows = _parse_desj_csv(path, DESJ)

    assert len(rows) == 1
    account_type, date, _, desc, withdrawal, deposit, interest, capital = rows[0]
    assert account_type == "PCA"
    assert date == datetime(2026, 7, 15)
    assert withdrawal == 4531.26
    assert deposit is None


def test_desj_ln1_unquoted_comma_reconstructed(tmp_path):
    """Bank sends description with an unquoted comma; parser must reconstruct it."""
    path = _write_csv(tmp_path, [_DESJ_LN1_UNQUOTED])
    rows = _parse_desj_csv(path, DESJ)

    assert len(rows) == 1
    account_type, date, _, desc, withdrawal, deposit, interest, capital = rows[0]
    assert account_type == "LN1"
    assert date == datetime(2026, 7, 15)
    assert desc == "Automatic payment /from PCA:4,531.26$"
    assert interest == 2326.48
    assert capital == 2204.78
    assert withdrawal is None
    assert deposit is None


def test_desj_ln1_quoted_description(tmp_path):
    """Properly CSV-quoted description with comma is handled via csv.reader directly."""
    path = _write_csv(tmp_path, [_DESJ_LN1_QUOTED])
    rows = _parse_desj_csv(path, DESJ)

    assert len(rows) == 1
    _, _, _, desc, _, _, interest, capital = rows[0]
    assert desc == "Automatic payment /from PCA:4,531.26$"
    assert interest == 2326.48
    assert capital == 2204.78


def test_desj_row_too_short_skipped(tmp_path):
    path = _write_csv(tmp_path, ["L'ile,057992,PCA,2026/07/11"])
    assert _parse_desj_csv(path, DESJ) == []


def test_desj_invalid_date_skipped(tmp_path):
    bad = _DESJ_PCA_DEPOSIT.replace("2026/07/11", "not-a-date")
    path = _write_csv(tmp_path, [bad])
    assert _parse_desj_csv(path, DESJ) == []


def test_desj_empty_date_skipped(tmp_path):
    # Replacing the date with empty shifts fields but triggers the empty-date guard
    parts = _DESJ_PCA_DEPOSIT.split(",")
    parts[3] = ""
    path = _write_csv(tmp_path, [",".join(parts)])
    assert _parse_desj_csv(path, DESJ) == []


def test_desj_multiple_rows_all_parsed(tmp_path):
    path = _write_csv(tmp_path, [_DESJ_PCA_DEPOSIT, _DESJ_PCA_WITHDRAWAL, _DESJ_LN1_UNQUOTED])
    assert len(_parse_desj_csv(path, DESJ)) == 3


# ── _parse_bmo_csv ───────────────────────────────────────────────────────────

_BMO_HEADER = "Carte,Type de transaction,Date,Amount,Description"
_BMO_ROW = "'5510290002861002',DEBIT,20260501,-240.26,Grocery Store"


def _bmo_file(tmp_path, lines):
    p = tmp_path / "bmo.csv"
    p.write_text("\n".join(lines) + "\n", encoding="utf-8-sig")
    return str(p)


def test_bmo_parses_row_after_header(tmp_path):
    path = _bmo_file(tmp_path, [_BMO_HEADER, _BMO_ROW])
    rows = _parse_bmo_csv(path, BMO)

    assert len(rows) == 1
    card, tx_type, date_str, amount_str, desc = rows[0]
    assert tx_type == "DEBIT"
    assert date_str == "20260501"
    assert float(amount_str) == -240.26
    assert desc == "Grocery Store"


def test_bmo_skips_rows_before_header(tmp_path):
    path = _bmo_file(tmp_path, ["noise,noise,noise,noise,noise", _BMO_HEADER, _BMO_ROW])
    assert len(_parse_bmo_csv(path, BMO)) == 1


def test_bmo_empty_date_skipped(tmp_path):
    bad = _BMO_ROW.replace("20260501", "")
    path = _bmo_file(tmp_path, [_BMO_HEADER, bad])
    assert _parse_bmo_csv(path, BMO) == []


def test_bmo_multiple_rows(tmp_path):
    row2 = "'5510290002861002',CREDIT,20260502,1000.00,Payroll"
    path = _bmo_file(tmp_path, [_BMO_HEADER, _BMO_ROW, row2])
    assert len(_parse_bmo_csv(path, BMO)) == 2


# ── _find_insert_row ─────────────────────────────────────────────────────────

def test_find_insert_row_first_empty_in_sequence():
    ws = _ws_with({(6, 2): 20260101, (7, 2): 20260201, (8, 2): None})
    assert _find_insert_row(ws, date_col=2, data_start_row=6) == 8


def test_find_insert_row_empty_sheet():
    ws = _ws_with({})
    assert _find_insert_row(ws, date_col=2, data_start_row=6) == 6


def test_find_insert_row_all_rows_filled():
    ws = _ws_with({(6, 2): 20260101, (7, 2): 20260201, (8, 2): 20260301})
    assert _find_insert_row(ws, date_col=2, data_start_row=6) == 9


def test_find_insert_row_respects_data_start_row():
    # Row 4 has data but data_start_row=6, so scan begins at 6
    ws = _ws_with({(4, 2): 20260101, (6, 2): None})
    assert _find_insert_row(ws, date_col=2, data_start_row=6) == 6


# ── _get_balance_simple ───────────────────────────────────────────────────────

def test_get_balance_simple_returns_last_non_none():
    ws = _ws_with({(6, 9): 1000.0, (7, 9): 1250.0, (8, 9): None})
    assert _get_balance_simple(ws, last_row=8, data_start_row=6, balance_col=9) == 1250.0


def test_get_balance_simple_returns_zero_when_all_none():
    ws = _ws_with({})
    assert _get_balance_simple(ws, last_row=7, data_start_row=6, balance_col=9) == 0.0


def test_get_balance_simple_includes_start_row():
    ws = _ws_with({(13, 9): 3164.98})
    assert _get_balance_simple(ws, last_row=13, data_start_row=13, balance_col=9) == 3164.98


def test_get_balance_simple_scans_backwards():
    # Row 8 is None, row 7 has value — should return row 7's value
    ws = _ws_with({(6, 9): 500.0, (7, 9): 750.0, (8, 9): None})
    assert _get_balance_simple(ws, last_row=8, data_start_row=6, balance_col=9) == 750.0


# ── _desj_duplicate_set ───────────────────────────────────────────────────────

def test_desj_duplicate_set_collects_existing_rows():
    d = datetime(2026, 5, 11)
    ws = _ws_with({(13, 2): d, (13, 4): "Money transfer"})
    dupes = _desj_duplicate_set(ws, insert_row=14, bank=DESJ)
    assert (d, "Money transfer") in dupes


def test_desj_duplicate_set_skips_none_date():
    ws = _ws_with({(13, 2): None, (13, 4): "Something"})
    dupes = _desj_duplicate_set(ws, insert_row=14, bank=DESJ)
    assert len(dupes) == 0


def test_desj_duplicate_set_multiple_rows():
    d1, d2 = datetime(2026, 5, 11), datetime(2026, 5, 15)
    ws = _ws_with({
        (13, 2): d1, (13, 4): "Deposit",
        (14, 2): d2, (14, 4): "Withdrawal",
    })
    dupes = _desj_duplicate_set(ws, insert_row=15, bank=DESJ)
    assert (d1, "Deposit") in dupes
    assert (d2, "Withdrawal") in dupes


# ── _bmo_duplicate_set ────────────────────────────────────────────────────────

def test_bmo_duplicate_set_collects_existing_rows():
    ws = _ws_with({(6, 4): 20260501, (6, 5): -240.26, (6, 6): "Grocery Store"})
    dupes = _bmo_duplicate_set(ws, insert_row=7, bank=BMO)
    assert (20260501, -240.26, "Grocery Store") in dupes


def test_bmo_duplicate_set_skips_row_with_none_amount():
    ws = _ws_with({(6, 4): 20260501, (6, 5): None, (6, 6): "Grocery Store"})
    dupes = _bmo_duplicate_set(ws, insert_row=7, bank=BMO)
    assert len(dupes) == 0


def test_bmo_duplicate_set_skips_row_with_none_date():
    ws = _ws_with({(6, 4): None, (6, 5): -240.26, (6, 6): "Grocery Store"})
    dupes = _bmo_duplicate_set(ws, insert_row=7, bank=BMO)
    assert len(dupes) == 0
