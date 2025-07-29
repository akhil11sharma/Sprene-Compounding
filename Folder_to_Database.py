import os
import pandas as pd
import pyodbc
import threading
import time
import tkinter as tk
from tkinter import scrolledtext

# --- SQL Server Configuration ---
DB_CONFIG = {
    'server': 'DESKTOP-PLHJ2M7',
    'database': 'test',
    'trusted_connection': 'yes'
}

FOLDER_PATH = r'C:\Users\Nikhil Sharma\Desktop\SSMS'

# --- DB Connection ---
def get_connection():
    return pyodbc.connect(
        f"DRIVER={{ODBC Driver 17 for SQL Server}};"
        f"SERVER={DB_CONFIG['server']};"
        f"DATABASE={DB_CONFIG['database']};"
        f"Trusted_Connection={DB_CONFIG['trusted_connection']};"
    )

# --- Clean Column Names ---
def validate_and_clean_data(df):
    seen = set()
    new_columns = []
    for i, col in enumerate(df.columns):
        new_col = col.strip() if str(col).strip() else f"Unnamed_{i+1}"
        while new_col in seen:
            new_col += "_dup"
        seen.add(new_col)
        new_columns.append(new_col)
    df.columns = new_columns

    for col in df.columns:
        if df[col].dtype == 'object':
            df[col] = df[col].fillna("").astype(str)
        elif pd.api.types.is_numeric_dtype(df[col]):
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
        elif pd.api.types.is_datetime64_any_dtype(df[col]):
            df[col] = pd.to_datetime(df[col], errors='coerce')
        else:
            df[col] = df[col].fillna("")
    return df

def clean_large_ints(df):
    for col in df.select_dtypes(include=['int', 'float']).columns:
        if df[col].max() > 2147483647:
            df[col] = df[col].astype('int64')
    return df

# --- SQL Helpers ---
def table_exists(cursor, table_name):
    cursor.execute("SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = ?", (table_name,))
    return cursor.fetchone() is not None

def drop_table(cursor, table_name):
    cursor.execute(f"DROP TABLE IF EXISTS [{table_name}]")

def create_table_information_table(cursor):
    cursor.execute("""
    IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='table_information' and xtype='U')
    CREATE TABLE table_information (
        table_name NVARCHAR(255) NOT NULL PRIMARY KEY,
        column_count INT NOT NULL,
        row_count INT NOT NULL DEFAULT 0
    )
    """)

def create_table(cursor, table_name, columns, dtypes):
    sql_types = {
        'object': 'NVARCHAR(MAX)',
        'int64': 'BIGINT',
        'float64': 'FLOAT',
        'datetime64[ns]': 'DATETIME',
        'bool': 'BIT'
    }
    defs = [f"[{col}] {sql_types.get(str(dtype), 'NVARCHAR(MAX)')}" for col, dtype in zip(columns, dtypes)]
    cursor.execute(f"CREATE TABLE [{table_name}] ({', '.join(defs)})")

def insert_data(cursor, table_name, columns, data):
    placeholders = ', '.join(['?'] * len(columns))
    cursor.executemany(
        f"INSERT INTO [{table_name}] ({', '.join(f'[{col}]' for col in columns)}) VALUES ({placeholders})",
        data
    )

def insert_metadata(cursor, table_name, col_count, row_count):
    cursor.execute("""
    IF EXISTS (SELECT * FROM table_information WHERE table_name = ?)
        UPDATE table_information SET column_count = ?, row_count = ? WHERE table_name = ?
    ELSE
        INSERT INTO table_information (table_name, column_count, row_count) VALUES (?, ?, ?)
    """, (table_name, col_count, row_count, table_name, table_name, col_count, row_count))

def read_file(file_path):
    if file_path.endswith(".csv"):
        return pd.read_csv(file_path)
    elif file_path.endswith(('.xlsx', '.xls')):
        return pd.read_excel(file_path, engine='openpyxl')
    else:
        raise ValueError("Unsupported file format.")

# --- File Processing Logic ---
def process_files(log_box):
    try:
        conn = get_connection()
        cursor = conn.cursor()
        create_table_information_table(cursor)

        for file in os.listdir(FOLDER_PATH):
            if not file.lower().endswith((".csv", ".xlsx", ".xls")):
                continue

            path = os.path.join(FOLDER_PATH, file)
            table_name = os.path.splitext(file)[0]

            try:
                df = read_file(path)
                df = validate_and_clean_data(df)
                df = clean_large_ints(df)

                columns, dtypes, data = df.columns.tolist(), df.dtypes.tolist(), df.values.tolist()

                if table_exists(cursor, table_name):
                    drop_table(cursor, table_name)

                create_table(cursor, table_name, columns, dtypes)
                insert_data(cursor, table_name, columns, data)
                insert_metadata(cursor, table_name, len(columns), len(data))

                log_box.insert(tk.END, f"[✓] {file} ➜ {table_name} updated.\n")
            except Exception as e:
                log_box.insert(tk.END, f"[✗] Error with {file}: {e}\n")

        conn.commit()
        conn.close()
    except Exception as e:
        log_box.insert(tk.END, f"!! Error: {e}\n")

# --- GUI Setup ---
class FolderToDBApp:
    def __init__(self, root):
        self.running = False
        self.thread = None

        root.title("Folder to Database Uploader")
        root.geometry("780x580")
        root.config(bg="#f0f8ff")

        # Title
        tk.Label(root, text="Sperene Folder to DB Sync", bg="#4682B4", fg="white",
                 font=("Times New Roman", 18, "bold"), pady=10).pack(fill=tk.X)

        # Description
        tk.Label(root,
                 text="FROM GIVEN FOLDER PATH IT WILL EXTRACT FILE EXTENSIONS AS XLSX & CSV AND SAVE THEM DYNAMICALLY TO THE DATABASE\n"
                      "ALSO IT WILL USE TIME DURATION OF 30 SEC TO ITERATE THE SAME PROCEDURE.\n"
                      "IF IN ANY CASE THEY ARE SOME CHANGES THEY WILL GET UPDATED AUTOMATICALLY.",
                 bg="#f0f8ff", fg="#222", font=("Times New Roman", 10), justify="center", wraplength=750).pack(pady=5)

        # Timer Label
        self.timer_var = tk.StringVar(value="")
        self.timer_label = tk.Label(root, textvariable=self.timer_var, bg="#f0f8ff",
                                    font=("Times New Roman", 11, "bold"), fg="#333")
        self.timer_label.pack()

        # Log box
        self.log_box = scrolledtext.ScrolledText(root, height=20, font=("Times New Roman", 10),
                                                 bg="white", fg="black")
        self.log_box.pack(padx=20, pady=10, fill=tk.BOTH, expand=True)

        # Button Frame
        btn_frame = tk.Frame(root, bg="#f0f8ff")
        btn_frame.pack(pady=10)

        self.start_btn = tk.Button(btn_frame, text="Start", command=self.start_loop, bg="#5F9EA0",
                                   fg="white", font=("Times New Roman", 12, "bold"), width=15)
        self.start_btn.grid(row=0, column=0, padx=10)

        self.stop_btn = tk.Button(btn_frame, text="Stop", command=self.stop_loop, bg="#B22222",
                                  fg="white", font=("Times New Roman", 12, "bold"), width=15)
        self.stop_btn.grid(row=0, column=1, padx=10)

    def start_loop(self):
        if not self.running:
            self.running = True
            self.log_box.insert(tk.END, ">> Initial sync started...\n")
            self.thread = threading.Thread(target=self.initial_and_loop)
            self.thread.daemon = True
            self.thread.start()

    def initial_and_loop(self):
        # First process immediately
        self.timer_var.set("🔄 Initial sync in progress...")
        process_files(self.log_box)
        self.log_box.insert(tk.END, "-" * 60 + "\n")
        # Then start timer loop
        while self.running:
            for i in range(30, 0, -1):
                self.timer_var.set(f"⏳ Next sync in: {i}s")
                time.sleep(1)
            self.timer_var.set("🔄 Syncing now...")
            self.log_box.insert(tk.END, f"\n[⏱] Processing folder...\n")
            process_files(self.log_box)
            self.log_box.insert(tk.END, "-" * 60 + "\n")

    def stop_loop(self):
        self.running = False
        self.timer_var.set("⏹️ Sync paused.")
        self.log_box.insert(tk.END, "\n>> Stopped syncing.\n")

# --- Launch ---
if __name__ == "__main__":
    root = tk.Tk()
    app = FolderToDBApp(root)
    root.mainloop()
