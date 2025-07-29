"""There shouldnt be PDF in the Excel Format"""
import os
import tkinter as tk
from tkinter import ttk, messagebox

# Folder path to check PDFs in
FOLDER_PATH = r'C:\Users\Nikhil Sharma\Desktop\Internship'  # Update if needed
pdf_files = []
file_combo = None
root = None

# ------------------- Utility Functions -------------------
def check_pdf_files():
    global pdf_files, file_combo
    if not os.path.exists(FOLDER_PATH):
        messagebox.showerror("Folder Not Found", f"The folder path '{FOLDER_PATH}' does not exist.")
        return

    pdf_files.clear()
    for filename in os.listdir(FOLDER_PATH):
        if filename.lower().endswith(".pdf"):
            pdf_files.append(filename)

    if not pdf_files:
        file_combo['values'] = []
        messagebox.showinfo("No PDF Files", "No PDF files found in the folder.")
    else:
        file_combo['values'] = pdf_files
        file_combo.current(0)

def delete_selected_file():
    selected_file = file_combo.get()
    if not selected_file:
        messagebox.showwarning("No File Selected", "Please select a PDF file to delete.")
        return

    confirmation = messagebox.askyesno("Delete File", f"Delete the file '{selected_file}'?")
    if confirmation:
        try:
            os.remove(os.path.join(FOLDER_PATH, selected_file))
            messagebox.showinfo("Deleted", f"'{selected_file}' deleted.")
            check_pdf_files()
        except Exception as e:
            messagebox.showerror("Error", f"Error deleting file: {e}")

def delete_all_files():
    if not pdf_files:
        messagebox.showinfo("No PDFs", "No PDF files to delete.")
        return

    confirmation = messagebox.askyesno("Delete All", "Delete ALL PDF files in the folder?")
    if confirmation:
        try:
            for file in pdf_files:
                os.remove(os.path.join(FOLDER_PATH, file))
            messagebox.showinfo("Deleted", "All PDF files deleted.")
            check_pdf_files()
        except Exception as e:
            messagebox.showerror("Error", f"Error deleting files: {e}")

def periodic_check():
    check_pdf_files()
    root.after(10000, periodic_check)  # Refresh every 10 seconds

# ------------------- UI Setup -------------------
def create_gui():
    global file_combo, root
    root = tk.Tk()
    root.title("PDF Validation & Cleanup")
    root.geometry("600x350")
    root.config(bg="#f0f8ff")

    # Header
    tk.Label(root, text="Sperene PDF Folder Validation", font=("Helvetica", 18, "bold"),
             bg="#4682B4", fg="white", pady=10).pack(fill=tk.X)

    # File Selection
    form_frame = tk.Frame(root, bg="#f0f8ff")
    form_frame.pack(pady=30)

    tk.Label(form_frame, text="Select a PDF to delete:", font=("Arial", 12), bg="#f0f8ff").grid(row=0, column=0, padx=10, pady=10)

    file_combo = ttk.Combobox(form_frame, state='readonly', font=("Arial", 11), width=40)
    file_combo.grid(row=0, column=1, padx=10, pady=10)

    # Buttons
    btn_frame = tk.Frame(root, bg="#f0f8ff")
    btn_frame.pack(pady=10)

    tk.Button(btn_frame, text="Delete Selected PDF", command=delete_selected_file, width=20,
              bg="#5F9EA0", fg="white", font=("Arial", 11, "bold")).grid(row=0, column=0, padx=10, pady=5)

    tk.Button(btn_frame, text="Delete All PDFs", command=delete_all_files, width=20,
              bg="#B22222", fg="white", font=("Arial", 11, "bold")).grid(row=0, column=1, padx=10, pady=5)

    # Initial Load
    check_pdf_files()
    periodic_check()
    root.mainloop()

# Entry
if __name__ == "__main__":
    create_gui()
