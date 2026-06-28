import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from config import BANKS
from importer import import_csv_to_excel


def browse_csv():
    path = filedialog.askopenfilename(
        title="Select CSV file",
        filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
    )
    if path:
        csv_path_var.set(path)


def browse_excel():
    path = filedialog.askopenfilename(
        title="Select Excel file",
        filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
    )
    if path:
        excel_path_var.set(path)


def run_import():
    csv_path = csv_path_var.get().strip()
    excel_path = excel_path_var.get().strip()
    bank_name = bank_var.get()
    if not csv_path:
        messagebox.showwarning("Missing file", "Please select a CSV file.")
        return
    if not excel_path:
        messagebox.showwarning("Missing file", "Please select an Excel file.")
        return

    import_btn.config(state="disabled")
    status_var.set("")
    spinner.grid()
    spinner.start(10)

    def worker():
        try:
            bank = BANKS[bank_name]
            result = import_csv_to_excel(csv_path, excel_path, bank)
            root.after(0, lambda: on_done(result, None))
        except Exception as e:
            root.after(0, lambda err=e: on_done(None, err))

    threading.Thread(target=worker, daemon=True).start()


def on_done(result, error):
    spinner.stop()
    spinner.grid_remove()
    import_btn.config(state="normal")
    if error:
        messagebox.showerror("Error", str(error))
    else:
        status_var.set(result)


root = tk.Tk()
root.title("Auto Made It")
root.resizable(False, False)

try:
    icon = tk.PhotoImage(file="logo.png")
    root.iconphoto(True, icon)
except Exception:
    pass

csv_path_var = tk.StringVar()
excel_path_var = tk.StringVar()
bank_var = tk.StringVar(value=list(BANKS.keys())[0])
status_var = tk.StringVar()

frame = tk.Frame(root, padx=12, pady=12)
frame.pack()

try:
    logo = tk.PhotoImage(file="logo.png")
    tk.Label(frame, image=logo).grid(row=0, column=0, columnspan=3, pady=(0, 10))
    frame._logo = logo
    row_offset = 1
except Exception:
    row_offset = 0

tk.Label(frame, text="Select CSV file to process", font=("Helvetica", 10)).grid(
    row=row_offset, column=0, columnspan=3, pady=(0, 8)
)

tk.Label(frame, text="Bank:").grid(row=row_offset + 1, column=0, sticky="w")
ttk.Combobox(frame, textvariable=bank_var, values=list(BANKS.keys()), state="readonly", width=47).grid(
    row=row_offset + 1, column=1, columnspan=2, padx=(4, 0), sticky="w"
)

tk.Label(frame, text="CSV file:").grid(row=row_offset + 2, column=0, sticky="w", pady=(6, 0))
tk.Entry(frame, textvariable=csv_path_var, width=50).grid(row=row_offset + 2, column=1, padx=(4, 4), pady=(6, 0))
tk.Button(frame, text="Browse…", command=browse_csv).grid(row=row_offset + 2, column=2, pady=(6, 0))

tk.Label(frame, text="Excel file:").grid(row=row_offset + 3, column=0, sticky="w", pady=(6, 0))
tk.Entry(frame, textvariable=excel_path_var, width=50).grid(row=row_offset + 3, column=1, padx=(4, 4), pady=(6, 0))
tk.Button(frame, text="Browse…", command=browse_excel).grid(row=row_offset + 3, column=2, pady=(6, 0))

import_btn = tk.Button(frame, text="Import to Excel", command=run_import, width=14)
import_btn.grid(row=row_offset + 4, column=0, columnspan=3, pady=(12, 4))

spinner = ttk.Progressbar(frame, mode="indeterminate", length=200)
spinner.grid(row=row_offset + 5, column=0, columnspan=3, pady=(0, 4))
spinner.grid_remove()

tk.Label(frame, textvariable=status_var, fg="green").grid(
    row=row_offset + 6, column=0, columnspan=3
)

root.mainloop()
