import tkinter as tk
from tkinter import filedialog, messagebox

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
    if not csv_path:
        messagebox.showwarning("Missing file", "Please select a CSV file.")
        return
    if not excel_path:
        messagebox.showwarning("Missing file", "Please select an Excel file.")
        return
    try:
        result = import_csv_to_excel(csv_path, excel_path)
        status_var.set(result)
    except FileNotFoundError as e:
        messagebox.showerror("File not found", str(e))
    except Exception as e:
        messagebox.showerror("Error", str(e))


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

tk.Label(frame, text="CSV file:").grid(row=row_offset + 1, column=0, sticky="w")
tk.Entry(frame, textvariable=csv_path_var, width=50).grid(row=row_offset + 1, column=1, padx=(4, 4))
tk.Button(frame, text="Browse…", command=browse_csv).grid(row=row_offset + 1, column=2)

tk.Label(frame, text="Excel file:").grid(row=row_offset + 2, column=0, sticky="w", pady=(6, 0))
tk.Entry(frame, textvariable=excel_path_var, width=50).grid(row=row_offset + 2, column=1, padx=(4, 4), pady=(6, 0))
tk.Button(frame, text="Browse…", command=browse_excel).grid(row=row_offset + 2, column=2, pady=(6, 0))

tk.Button(frame, text="Import to Excel", command=run_import, width=14).grid(
    row=row_offset + 3, column=0, columnspan=3, pady=(12, 4)
)

tk.Label(frame, textvariable=status_var, fg="green").grid(
    row=row_offset + 4, column=0, columnspan=3
)

root.mainloop()
