"""
This GUI dashboard provides quick access to essential tools for database and file operations. 
Users can manage tables, validate folders, extract Excel formulas, and more with a single click.
"""
import tkinter as tk
from tkinter import ttk, filedialog
import subprocess
import os

# --- Config ---
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SCRIPTS = {
    "CURD (Table Editor)": "CURD_FILE.py",
    "Folder to Database (Auto Sync)": "Folder_to_Database.py",
    "Folder Validation (PDF Checker)": "folder_validation.py",
    "Extract Table Formulas": "Table_area.py",
    "Highlight Formulas in Excel": "validation_check_excel.py",
}

selected_file_name = ""  # Track selected file name

# --- Script Runner ---
def run_script(script_name, with_file_dialog=False):
    global selected_file_name
    script_path = os.path.join(BASE_DIR, script_name)
    if with_file_dialog:
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if not file_path:
            return
        selected_file_name = os.path.basename(file_path)
        selected_file_label.config(text=f"Selected File: {selected_file_name}")
        subprocess.Popen(["python", script_path, file_path])
    else:
        selected_file_name = ""
        selected_file_label.config(text="")
        subprocess.Popen(["python", script_path])

# --- Tooltip Helper ---
def create_tooltip(widget, text):
    tooltip = tk.Toplevel(widget)
    tooltip.withdraw()
    tooltip.wm_overrideredirect(True)
    tooltip_label = tk.Label(tooltip, text=text, background="#333", foreground="white",
                             font=("Segoe UI", 8), padx=6, pady=2)
    tooltip_label.pack()

    def on_enter(event):
        x = event.x_root + 10
        y = event.y_root + 5
        tooltip.geometry(f"+{x}+{y}")
        tooltip.deiconify()

    def on_leave(event):
        tooltip.withdraw()

    widget.bind("<Enter>", on_enter)
    widget.bind("<Leave>", on_leave)

# --- GUI Main ---
def main():
    global selected_file_label

    root = tk.Tk()
    root.title("Sperene Compounding Dashboard")
    root.geometry("1000x600")
    root.configure(bg="#e7eff6")

    style = ttk.Style()
    style.theme_use("clam")
    style.configure("Rounded.TButton",
                    font=("Segoe UI", 11, "bold"),
                    padding=10,
                    relief="flat",
                    background="#4A90E2",
                    foreground="white")
    style.map("Rounded.TButton",
              foreground=[("active", "white")],
              background=[("active", "#357ABD")])

    # --- Layout Frames ---
    left_frame = tk.Frame(root, bg="#4A90E2", width=360)
    left_frame.pack(side="left", fill="y")

    right_frame = tk.Frame(root, bg="#ffffff")
    right_frame.pack(side="right", expand=True, fill="both", padx=30, pady=30)

    # --- Left Pane ---
    tk.Label(
        left_frame,
        text="Welcome to\nSperene Compounding System",
        font=("Helvetica", 20, "bold"),
        bg="#4A90E2",
        fg="white",
        justify="left"
    ).pack(padx=25, pady=(50, 20), anchor="w")

    description = (
        "Empower your workflow with tools to:\n\n"
        "• 📊 Manage database tables efficiently\n"
        "• 🔄 Auto-sync folders to your database\n"
        "• 🧾 Validate and clean PDF files\n"
        "• 📐 Extract and audit Excel formulas\n"
        "• 🖍️ Highlight formulas for accuracy"
    )
    tk.Label(
        left_frame,
        text=description,
        font=("Segoe UI", 10, "bold"),
        bg="#4A90E2",
        fg="#eaf4fc",
        justify="left",
        wraplength=300
    ).pack(padx=25, anchor="w")

    tk.Label(
        left_frame,
        text="\nNeed Help? Contact Support",
        font=("Segoe UI", 9, "italic"),
        bg="#4A90E2",
        fg="#cce0f5"
    ).pack(side="bottom", pady=20)

    # --- Right Pane ---
    tk.Label(
        right_frame,
        text="Available Tools",
        font=("Helvetica", 16, "bold"),
        bg="#ffffff",
        fg="#222"
    ).pack(pady=(0, 20), anchor="w")

    card = tk.Frame(right_frame, bg="#f8fbfd", bd=2, relief="groove")
    card.pack(fill="both", expand=True, padx=10, pady=10)

    button_defs = [
        ("🧾 CURD (Table Editor)", SCRIPTS["CURD (Table Editor)"], False, "Edit and manage your SQL table records"),
        ("📁 Folder to Database (Auto Sync)", SCRIPTS["Folder to Database (Auto Sync)"], False, "Automatically sync folder contents to database"),
        ("🗃️ Folder Validation (PDF Checker)", SCRIPTS["Folder Validation (PDF Checker)"], False, "Scan and validate PDF files"),
        ("📐 Extract Table Formulas", SCRIPTS["Extract Table Formulas"], False, "Extract Excel formulas from selected sheet"),
        ("🖍️ Highlight Formulas in Excel", SCRIPTS["Highlight Formulas in Excel"], True, "Mark formulas directly in Excel file"),
    ]

    for text, script, needs_file, tooltip in button_defs:
        btn = ttk.Button(
            card,
            text=text,
            width=52,
            style="Rounded.TButton",
            command=lambda s=script, f=needs_file: run_script(s, f)
        )
        btn.pack(pady=12)
        create_tooltip(btn, tooltip)

    selected_file_label = tk.Label(
        right_frame,
        text="",
        font=("Segoe UI", 10, "italic"),
        bg="#ffffff",
        fg="#333"
    )
    selected_file_label.pack(pady=(0, 10), anchor="w")

    tk.Label(
        right_frame,
        text="~ Powered by Sperene Technologies",
        font=("Segoe UI", 9, "italic"),
        bg="#ffffff",
        fg="#999"
    ).pack(side="bottom", pady=10)

    root.mainloop()

# --- Launch ---
if __name__ == "__main__":
    main()