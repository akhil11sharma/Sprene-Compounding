import tkinter as tk
from tkinter import messagebox, StringVar
from ttkbootstrap import Style
from ttkbootstrap.widgets import Frame, Entry, Button, Label, Treeview, Combobox, Scrollbar
import pyodbc

# --- Database Configuration ---
DB_CONFIG = {
    'server': 'DESKTOP-PLHJ2M7',
    'database': 'test',
    'trusted_connection': 'yes',
}

changes_buffer = []

# --- Execute SQL Queries ---
def execute_query(query, params=None, fetch=False):
    try:
        conn = pyodbc.connect(
            f'DRIVER={{ODBC Driver 17 for SQL Server}};'
            f'SERVER={DB_CONFIG["server"]};'
            f'DATABASE={DB_CONFIG["database"]};'
            'Trusted_Connection=yes;'
        )
        cursor = conn.cursor()
        cursor.execute(query, params or [])
        if fetch:
            result = cursor.fetchall()
            conn.close()
            return result
        conn.commit()
        conn.close()
    except pyodbc.Error as err:
        messagebox.showerror("SQL Error", str(err))

# --- Load Table Columns ---
def load_table_columns(table_name):
    cols = execute_query("SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = ?", (table_name,), fetch=True)
    return [col[0] for col in cols]

# --- Entry Fields Section ---
def add_entry_fields():
    global entry_vars
    for widget in entry_frame.winfo_children():
        widget.destroy()

    table = selected_table.get()
    if not table:
        return

    columns = load_table_columns(table)
    entry_vars = {}

    for i, col in enumerate(columns):
        Label(entry_frame, text=col, font=("Times New Roman", 11, "bold")).grid(row=0, column=i, padx=10, pady=5)
        var = StringVar()
        entry_vars[col] = var
        Entry(entry_frame, textvariable=var, font=("Times New Roman", 10), width=18).grid(row=1, column=i, padx=10, pady=5)

    # Buttons
    btn_kwargs = {"bootstyle": "success-outline", "width": 18, "padding": 10}
    Button(entry_frame, text="➕ Add", command=add_record, **btn_kwargs).grid(row=2, column=0, pady=15)
    Button(entry_frame, text="✏️ Update", command=update_record, **btn_kwargs).grid(row=2, column=1, pady=15)
    Button(entry_frame, text="🗑️ Delete", command=delete_record, **btn_kwargs).grid(row=2, column=2, pady=15)
    Button(entry_frame, text="💾 Save to Database", command=save_changes, bootstyle="primary-outline", width=20, padding=10).grid(row=2, column=3, pady=15)

# --- CRUD Buffer Operations ---
def add_record():
    table = selected_table.get()
    columns = load_table_columns(table)
    values = [entry_vars[col].get().strip() for col in columns]
    if all(v == '' for v in values):
        messagebox.showwarning("Input Error", "All fields are empty!")
        return
    tree.insert("", "end", values=values)
    changes_buffer.append(("add", values))
    clear_inputs()

def update_record():
    selected = tree.focus()
    if not selected:
        messagebox.showwarning("No Selection", "Please select a row to update.")
        return
    columns = load_table_columns(selected_table.get())
    new_values = [entry_vars[col].get().strip() for col in columns]
    old_values = tree.item(selected, "values")
    tree.item(selected, values=new_values)
    changes_buffer.append(("update", old_values, new_values))
    clear_inputs()

def delete_record():
    selected = tree.focus()
    if not selected:
        messagebox.showwarning("No Selection", "Please select a row to delete.")
        return
    old_values = tree.item(selected, "values")
    tree.delete(selected)
    changes_buffer.append(("delete", old_values))
    clear_inputs()

def clear_inputs():
    for var in entry_vars.values():
        var.set("")

def save_changes():
    table = selected_table.get()
    columns = load_table_columns(table)

    for change in changes_buffer:
        if change[0] == "add":
            values = change[1]
            valid_cols = [col for col, val in zip(columns, values) if val != '']
            valid_vals = [val for val in values if val != '']
            query = f"INSERT INTO [{table}] ({', '.join(f'[{c}]' for c in valid_cols)}) VALUES ({', '.join(['?'] * len(valid_vals))})"
            execute_query(query, valid_vals)

        elif change[0] == "update":
            old_values, new_values = change[1], change[2]
            pk = old_values[0]
            query = f"UPDATE [{table}] SET {', '.join([f'[{col}] = ?' for col in columns[1:]])} WHERE [{columns[0]}] = ?"
            execute_query(query, new_values[1:] + [pk])

        elif change[0] == "delete":
            pk = change[1][0]
            query = f"DELETE FROM [{table}] WHERE [{columns[0]}] = ?"
            execute_query(query, (pk,))

    messagebox.showinfo("Saved", "All changes saved successfully.")
    changes_buffer.clear()
    refresh_table()

def refresh_table():
    for i in tree.get_children():
        tree.delete(i)
    table = selected_table.get()
    if not table:
        return
    columns = load_table_columns(table)
    rows = execute_query(f"SELECT * FROM [{table}]", fetch=True)
    tree["columns"] = columns
    tree["show"] = "headings"
    for col in columns:
        tree.heading(col, text=col)
        tree.column(col, width=150)
    for row in rows:
        tree.insert("", "end", values=row)

# --- Main GUI ---
def main():
    global root, selected_table, entry_frame, tree

    style = Style("litera")
    root = style.master
    root.title("🗂️ Sperene CURD Manager")
    root.geometry("1100x700")
    root.configure(bg="#f0f8ff")

    # Title Header with Uniform Background
    title_frame = Frame(root, padding=15, style="primary.TFrame")
    title_frame.pack(fill=tk.X)
    Label(title_frame,
          text="📄 CURD System for SQL Tables",
          font=("Times New Roman", 26, "bold"),
          anchor="center",
          bootstyle="inverse-primary").pack()

    # Top Control Frame
    top_frame = Frame(root, padding=10)
    top_frame.pack(pady=10)

    selected_table = StringVar()
    tables = execute_query("SELECT name FROM sys.tables", fetch=True)
    table_list = [tbl[0] for tbl in tables]
    if table_list:
        selected_table.set(table_list[0])

    Label(top_frame, text="Select Table:", font=("Times New Roman", 11)).grid(row=0, column=0, padx=10)
    Combobox(top_frame, textvariable=selected_table, values=table_list, font=("Times New Roman", 10), width=20).grid(row=0, column=1)
    Button(top_frame, text="🔄 Load Table", command=refresh_table, bootstyle="info-outline", width=16, padding=7).grid(row=0, column=2, padx=10)

    # Entry Fields
    entry_frame = Frame(root, padding=10)
    entry_frame.pack(pady=10)

    # TreeView Display
    tree_frame = Frame(root, bootstyle="secondary", padding=5)
    tree_frame.pack(padx=20, pady=10, fill=tk.BOTH, expand=True)

    tree_scroll = Scrollbar(tree_frame)
    tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)

    tree = Treeview(tree_frame, yscrollcommand=tree_scroll.set)
    tree.pack(fill=tk.BOTH, expand=True)
    tree_scroll.config(command=tree.yview)

    # Populate Fields and Table
    add_entry_fields()
    refresh_table()
    root.mainloop()

if __name__ == "__main__":
    main()
