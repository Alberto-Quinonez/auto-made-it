import csv
import tkinter as tk
from tkinter import filedialog, messagebox


def browse_file():
    path = filedialog.askopenfilename(
        title="Select CSV file",
        filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
    )
    if path:
        path_var.set(path)


def load_csv():
    path = path_var.get().strip()
    if not path:
        messagebox.showwarning("No file", "Please select a CSV file first.")
        return
    try:
        with open(path, newline="", encoding="utf-8") as f:
            reader = csv.reader(f)
            for row in reader:
                print(row)
        print(f"Done reading: {path}")
    except FileNotFoundError:
        messagebox.showerror("Error", f"File not found:\n{path}")
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

path_var = tk.StringVar()

frame = tk.Frame(root, padx=12, pady=12)
frame.pack()

try:
    logo = tk.PhotoImage(file="logo.png")
    tk.Label(frame, image=logo).grid(row=0, column=0, columnspan=3, pady=(0, 10))
    frame._logo = logo  # prevent garbage collection
    row_offset = 1
except Exception:
    row_offset = 0

tk.Label(frame, text="Select CSV file to process", font=("Helvetica", 10)).grid(
    row=row_offset, column=0, columnspan=3, pady=(0, 8)
)
tk.Label(frame, text="CSV file:").grid(row=row_offset + 1, column=0, sticky="w")
tk.Entry(frame, textvariable=path_var, width=50).grid(row=row_offset + 1, column=1, padx=(4, 4))
tk.Button(frame, text="Browse…", command=browse_file).grid(row=row_offset + 1, column=2)
tk.Button(frame, text="Load & Print", command=load_csv, width=14).grid(
    row=row_offset + 2, column=0, columnspan=3, pady=(10, 0)
)

root.mainloop()
