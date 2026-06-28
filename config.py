from dataclasses import dataclass


@dataclass
class BankConfig:
    name: str
    excel_sheet: str


@dataclass
class BMOConfig(BankConfig):
    # CSV
    header_col_index: int = 0
    header_col_value: str = ""
    csv_card: int = 0
    csv_type: int = 0
    csv_date: int = 0
    csv_amount: int = 0
    csv_desc: int = 0
    # Excel
    excel_data_start_row: int = 0
    excel_card: int = 0
    excel_type: int = 0
    excel_date: int = 0
    excel_amount: int = 0
    excel_desc: int = 0
    excel_dr: int = 0
    excel_cr: int = 0
    excel_balance: int = 0
    excel_je: int = 0


@dataclass
class DESJConfig(BankConfig):
    # CSV (0-based)
    csv_account_type: int = 0
    csv_date: int = 0
    csv_code: int = 0
    csv_desc: int = 0
    csv_withdrawal: int = 0
    csv_deposit: int = 0
    csv_interest: int = 0
    csv_capital: int = 0
    # Excel (1-based)
    excel_data_start_row: int = 0
    excel_date: int = 0
    excel_code: int = 0
    excel_desc: int = 0
    excel_withdrawal: int = 0
    excel_deposit: int = 0
    excel_interest: int = 0
    excel_capital: int = 0
    excel_balance: int = 0
    excel_mortgage_balance: int = 0
    excel_je: int = 0


BMO = BMOConfig(
    name="BMO",
    excel_sheet="BMO - CASH Plex",
    header_col_index=1,
    header_col_value="Type de transaction",
    csv_card=0,
    csv_type=1,
    csv_date=2,
    csv_amount=3,
    csv_desc=4,
    excel_data_start_row=6,
    excel_card=2,
    excel_type=3,
    excel_date=4,
    excel_amount=5,
    excel_desc=6,
    excel_dr=7,
    excel_cr=8,
    excel_balance=9,
    excel_je=10,
)

DESJ = DESJConfig(
    name="DESJ",
    excel_sheet="DESJ - CASH Plex",
    csv_account_type=2,
    csv_date=3,
    csv_code=4,
    csv_desc=5,
    csv_withdrawal=7,
    csv_deposit=8,
    csv_interest=9,
    csv_capital=12,
    excel_data_start_row=13,
    excel_date=2,
    excel_code=3,
    excel_desc=4,
    excel_withdrawal=5,
    excel_deposit=6,
    excel_interest=7,
    excel_capital=8,
    excel_balance=9,
    excel_mortgage_balance=10,
    excel_je=11,
)

BANKS: dict[str, BankConfig] = {
    "BMO": BMO,
    "DESJ": DESJ,
}
