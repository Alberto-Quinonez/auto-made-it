# Auto Made It

## What does this app do?

Auto Made It is a desktop app that imports bank statement CSV files directly into a specific tab of your Excel workbook. It supports two banks — **BMO** and **Desjardins (DESJ)** — each with their own column layout and rules.

When you import a file, the app:

- Reads the CSV exported from your bank
- Finds where new data should be inserted in the correct Excel tab
- Skips any transactions that are already in the sheet (duplicate detection)
- Writes each new transaction into the right columns with the correct formatting
- Updates running balances automatically

For **BMO**, the app calculates debits, credits, and the running cash balance.

For **Desjardins**, the app handles two account types in the same file:
- **PCA** (chequing) rows — updates withdrawals, deposits, and the running cash balance
- **LN1** (mortgage) rows — updates interest paid, capital payment, and the running mortgage balance

---

## Before you start

You will need **Python** installed on your computer. If you are not sure whether you have it, open a terminal and type:

```
python --version
```

If you see a version number (e.g. `Python 3.12.0`), you are good to go. If not, download Python from [python.org](https://www.python.org/downloads/) and install it.

You will also need to install the required packages. In a terminal, navigate to the app folder and run:

```
pip install -r requirements.txt
```

---

## How to run the app

1. Open a terminal (search for **Command Prompt** or **PowerShell** in the Start menu)
2. Navigate to the app folder:

```
cd C:\Users\test\dev\auto-made-it
```

3. Start the app:

```
python main.py
```

A window will open — that's the app.

---

## How to use the app

1. Use the **Bank** dropdown to select either **BMO** or **DESJ**
2. Click **Browse…** next to "CSV file" and select the bank statement you downloaded
3. Click **Browse…** next to "Excel file" and select your workbook
4. Click **Import to Excel**
5. A spinner will appear while the import runs
6. When it finishes, a message will show how many rows were added and how many were skipped as duplicates

> **Important:** Make sure the Excel file is **closed** before importing. If Excel has the file open, the import will fail with a permission error.

---

## Supported banks

### BMO
- CSV exported from BMO online banking
- Data is written into the **BMO - CASH Plex** tab
- Each row gets: card number, transaction type, date, amount, description, debit, credit, and running balance

### Desjardins (DESJ)
- CSV exported from Desjardins online banking
- Data is written into the **DESJ - CASH Plex** tab
- PCA rows: date, code, description, withdrawal, deposit, and running cash balance
- LN1 rows: date, code, description, mortgage interest, capital payment, and running mortgage balance

---

## Adding or modifying a bank configuration

All column mappings live in `config.py`. You never need to touch `importer.py` just to adjust which column a value lives in.

### Column numbering

- **CSV columns** are **0-based** (the first column is `0`)
- **Excel columns** are **1-based** (column A = `1`, column B = `2`, etc.)

### Modifying an existing config

Open `config.py` and find the instance at the bottom of the file (`BMO = BMOConfig(...)` or `DESJ = DESJConfig(...)`). Change the number next to the field you need to update.

For example, if Desjardins moves the deposit column from column F (6) to column G (7) in your Excel sheet:

```python
DESJ = DESJConfig(
    ...
    excel_deposit=7,   # was 6
    ...
)
```

After changing a column number, run the tests to catch any obvious breakage:

```
python -m pytest tests/ -v
```

### Adding a new bank

Adding a new bank requires three steps.

**Step 1 — Add a config class in `config.py`**

If the new bank has a simple structure similar to BMO (one account type, flat rows), subclass `BMOConfig`. If it has multiple account types or more complex columns, subclass `BankConfig` directly and define your own fields:

```python
@dataclass
class MyBankConfig(BankConfig):
    # CSV column indices (0-based)
    csv_date: int = 0
    csv_desc: int = 0
    csv_amount: int = 0
    # Excel column indices (1-based)
    excel_data_start_row: int = 0
    excel_date: int = 0
    excel_desc: int = 0
    excel_amount: int = 0
    excel_balance: int = 0
```

Then add an instance below the class:

```python
MYBANK = MyBankConfig(
    name="MYBANK",
    excel_sheet="MyBank - CASH Plex",
    csv_date=2,
    csv_desc=4,
    csv_amount=3,
    excel_data_start_row=6,
    excel_date=2,
    excel_desc=4,
    excel_amount=5,
    excel_balance=6,
)
```

Finally, register it in the `BANKS` dictionary so it appears in the app dropdown:

```python
BANKS: dict[str, BankConfig] = {
    "BMO": BMO,
    "DESJ": DESJ,
    "MYBANK": MYBANK,
}
```

**Step 2 — Add an import function in `importer.py`**

Add a function `_import_mybank(csv_path, excel_path, bank)` modelled after `_import_bmo` or `_import_desj`, then wire it into the dispatcher at the top:

```python
def import_csv_to_excel(csv_path, excel_path, bank):
    if isinstance(bank, BMOConfig):
        return _import_bmo(csv_path, excel_path, bank)
    elif isinstance(bank, DESJConfig):
        return _import_desj(csv_path, excel_path, bank)
    elif isinstance(bank, MyBankConfig):
        return _import_mybank(csv_path, excel_path, bank)
    raise ValueError(f"Unknown bank config type: {type(bank)}")
```

**Step 3 — Add tests**

Add a test CSV string and at least one parsing test in `tests/test_importer.py` to lock in the column mapping. This prevents a future column-number typo from going unnoticed.

---

## Running the tests

Unit tests cover the CSV parsing logic, duplicate detection, balance calculation, and row insertion. To run them:

```
python -m pytest tests/ -v
```

All 33 tests should pass.

---

## Project structure

```
auto-made-it/
├── main.py          # Desktop GUI (tkinter)
├── importer.py      # Import logic for BMO and DESJ
├── config.py        # Column mappings for each bank
├── tests/
│   └── test_importer.py   # Unit tests
├── requirements.txt
└── README.md
```
