import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path

from core import scan_folder, export_docx


def generate() -> None:
    folder = folder_var.get().strip()
    if not folder:
        messagebox.showwarning("No folder", "Please choose a folder first.")
        return

    p = Path(folder)
    if not p.is_dir():
        messagebox.showerror("Invalid folder", "That folder does not exist.")
        return

    try:
        rows = scan_folder(str(p))
        if not rows:
            messagebox.showinfo("No files", "No files were found in that folder.")
            return
        out_path = p / "files-report.docx"
        export_docx(rows, str(out_path), "Files report")
    except Exception as e:  # noqa: BLE001
        messagebox.showerror("Error", str(e))
        return

    messagebox.showinfo("Done", f"Created Word file:\n{out_path}")


def browse() -> None:
    d = filedialog.askdirectory(title="Choose folder to scan")
    if d:
        folder_var.set(d)


root = tk.Tk()
root.title("Folder → Word table (simple)")
root.geometry("600x140")

folder_var = tk.StringVar()

frame = tk.Frame(root, padx=12, pady=12)
frame.pack(fill="both", expand=True)

tk.Label(frame, text="Folder:").grid(row=0, column=0, sticky="w")
entry = tk.Entry(frame, textvariable=folder_var, width=60)
entry.grid(row=0, column=1, padx=(4, 4), sticky="we")
tk.Button(frame, text="Browse…", command=browse).grid(row=0, column=2, sticky="e")

tk.Button(frame, text="Generate Word table", command=generate).grid(
    row=1, column=1, columnspan=2, pady=(12, 0), sticky="e"
)

frame.columnconfigure(1, weight=1)

root.mainloop()

