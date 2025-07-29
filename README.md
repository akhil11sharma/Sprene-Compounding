# Sprene-Compounding: Automated Data Management & Excel Toolkit

---

## 🌟 Project Overview

**Sprene-Compounding** is a Python-based suite of GUI tools developed to streamline data management, Excel formula extraction, and folder validation tasks. The core of this system is a user-friendly dashboard that integrates various utilities for efficient interaction with SQL Server databases and local files—all wrapped in an intuitive interface.

---

## 🚀 Features

- **Main Dashboard (GUI):** Central hub for launching all utilities.
- **Table Editor (CRUD System):** Perform Create, Read, Update, and Delete operations on SQL Server tables with an intuitive interface.
- **Folder Auto-Sync to Database:** Automatically import and synchronize `.csv`, `.xlsx`, and `.xls` files from a specified folder into SQL Server, ensuring the database is always up-to-date.
- **PDF Folder Checker & Cleaner:** Scan folders for unwanted PDF files and easily remove them.
- **Excel Formula Extractor:** Extract all formulas from an Excel worksheet, map them to headers and cell locations, and store them in the database for auditing and documentation.

---

## 🛠️ How to Run the Application

### Prerequisites

- **Python 3.x** installed  
- **SQL Server** configured (ensure connection details are updated in the respective scripts, e.g., `CURD_FILE.py`, `Folder_to_Database.py`, `Table_area.py`)
- **Required Python libraries:**  
  Install dependencies with:

  ```bash
  pip install tkinter ttkbootstrap pyodbc pandas openpyxl
  ```

  > _Note_:  
  > - `tkinter` is usually bundled with Python, but `ttkbootstrap` provides modern styling.  
  > - `openpyxl` is crucial for Excel operations.  
  > - You may need to install the [ODBC Driver for SQL Server](https://docs.microsoft.com/en-us/sql/connect/odbc/download-odbc-driver-for-sql-server) if it's not already present.

---

### 🚦 Running the Dashboard

1. **Clone the repository:**

   ```bash
   git clone https://github.com/akhil11sharma/Sprene-Compounding.git
   cd Sprene-Compounding
   ```

2. **Launch the main dashboard:**

   ```bash
   python Sprerene.py
   ```

3. **Interact with the tools:**  
   From the dashboard, click buttons to open each specific utility in a separate window.

---

## 📦 Module Descriptions

This project is modular, with each core functionality encapsulated in its own Python script.

---

### 🖥️ Sprerene.py — Main Dashboard Launcher

**Purpose:**  
Primary GUI dashboard providing quick access to all core tools for database and file operations.

**Highlights:**
- Built with `tkinter` for a modern split-pane dashboard.
- Launches other utility scripts (`CURD_FILE.py`, `Folder_to_Database.py`, `folder_validation.py`, `Table_area.py`) as independent processes using `subprocess.Popen`.
- Includes tooltips, file dialog support, and tracks the currently selected file.
- Provides status and help messages for an improved user experience.

**Usage:**  
```bash
python Sprerene.py
```
_Run this to open the main application dashboard._

---

### 📊 CURD_FILE.py — Table Editor (CRUD System)

**Purpose:**  
Graphical user interface to perform Create, Read, Update, and Delete (CRUD) operations on SQL Server tables.

**Highlights:**
- Connects to SQL Server via `pyodbc`.
- Loads available tables into a dropdown for selection.
- Displays table contents in a `Treeview` grid.
- Add, edit, and delete rows directly from the GUI.
- Changes are committed to the SQL Server database.

**Usage:**  
- Run directly, or launch from `Sprerene.py`.
- Select a table and manage its data visually.

---

### 📁 Folder_to_Database.py — Folder Auto-Sync to Database

**Purpose:**  
Automatically imports all `.csv`, `.xlsx`, and `.xls` files from a specified folder into SQL Server, keeping tables updated.

**Highlights:**
- Scans folder for supported file types.
- Reads files with `pandas`, validates, and cleans data.
- Dynamically creates or updates matching tables in SQL Server.
- Inserts data and metadata (table name, column/row count).
- Periodically rescans for new or changed files (every 30s).
- Logs actions and errors in the GUI.
- Start/stop synchronization with a button.

**Usage:**  
- Run directly, or launch from `Sprerene.py`.
- Configure your folder, then start the sync.

---

### 📂 folder_validation.py — PDF Folder Checker & Cleaner

**Purpose:**  
Scans a folder for unwanted `.pdf` files and provides a GUI to selectively or entirely delete them.

**Highlights:**
- Displays all found PDF files in a dropdown.
- Allows deletion of individual or all PDFs with a single click.
- Periodically rescans folder (every 10s).
- Informative message boxes for actions, errors, and statuses.

**Usage:**  
- Run directly, or launch from `Sprerene.py`.
- Keep Excel-focused folders free of unwanted PDFs.

---

### 📐 Table_area.py — Excel Formula Extractor

**Purpose:**  
Extracts all formulas from a selected Excel worksheet, maps them to headers and cell locations, and stores them in the database for auditing.

**Highlights:**
- Prompts user to select an Excel file via file dialog.
- Loads workbook, scans all cells for formulas.
- Extracts formula string, associated header, cell name, and involved columns.
- Saves discovered formulas and metadata to a database table.
- Provides feedback on extraction process.

**Usage:**  
- Run directly, or launch from `Sprerene.py`.
- Select an Excel file to extract and store its formulas for auditing.

---

## 📷 Screenshots

> _Add screenshots of the dashboard and each utility here for better visualization!_  
> Example:

<p align="center">
  <img src="assets/dashboard.png" width="600" alt="Main Dashboard Preview">
</p>

---

## 🤝 Contributing

Contributions, suggestions, and issue reports are welcome!  
Feel free to fork, submit pull requests, or open issues.

---

## 📞 Support

For any questions or support, please contact the repository owner via GitHub or refer to the "Need Help? Contact Support" message within the dashboard.

---

## ⚖️ License

Copyright (c) 2025 Akhil Sharma

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

- The above copyright notice and this permission notice shall be included in all
  copies or substantial portions of the Software.

**THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.**

---
