from dataclasses import dataclass


@dataclass
class BankConfig:
    name: str
    # CSV: which column/value identifies the header row
    header_col_index: int
    header_col_value: str
    # CSV column indices (0-based)
    csv_card: int
    csv_type: int
    csv_date: int
    csv_amount: int
    csv_desc: int
    # Excel target sheet
    excel_sheet: str
    excel_data_start_row: int
    # Excel column indices (1-based)
    excel_card: int
    excel_type: int
    excel_date: int
    excel_amount: int
    excel_desc: int
    excel_dr: int
    excel_cr: int
    excel_balance: int
    excel_je: int


BMO = BankConfig(
    name="BMO",
    header_col_index=1,
    header_col_value="Type de transaction",
    csv_card=0,
    csv_type=1,
    csv_date=2,
    csv_amount=3,
    csv_desc=4,
    excel_sheet="BMO - CASH Plex",
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

# Add new bank configs here
BANKS: dict[str, BankConfig] = {
    "BMO": BMO,
}
