import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import pandas as pd

def select_file1():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    if file_path:
        entry_file1.delete(0, tk.END)
        entry_file1.insert(tk.END, file_path)

def select_file2():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    if file_path:
        entry_file2.delete(0, tk.END)
        entry_file2.insert(tk.END, file_path)

def compare_files():
    file1_path = entry_file1.get()
    file2_path = entry_file2.get()
    field_name = entry_field.get().strip()  

    if file1_path and file2_path:
        try:
            df1 = pd.read_excel(file1_path)
            df2 = pd.read_excel(file2_path)

            # Convert column names to lowercase and trim extra spaces
            df1.columns = df1.columns.str.strip()
            df2.columns = df2.columns.str.strip()

            # Check if the specified field exists in both DataFrames
            if field_name not in df1.columns or field_name not in df2.columns:
                messagebox.showerror("Error", f"Field '{field_name}' does not exist in one or both Excel files.")
                return
                        
            duplicates = pd.merge(df1, df2, on=field_name, how='inner')[field_name]
            duplicates = duplicates.dropna()  
            if not duplicates.empty:
                duplicate_names_text.delete('1.0', tk.END)  
                duplicate_names_text.insert(tk.END, duplicates.to_string(index=False))                
            else:
                duplicate_names_text.delete('1.0', tk.END)  
                duplicate_names_text.insert(tk.END, "No duplicate values found in field '{}'.".format(field_name))
        except Exception as e:
            messagebox.showerror("Error", "An error occurred: " + str(e))
    else:
        messagebox.showerror("Error", "Please select both Excel files.")

# Creating main application window
root = tk.Tk()
root.title("Excel File Duplicate Value Comparison")
root.geometry("700x400")

label_file1 = tk.Label(root, text="Select the first Excel file:")
label_file1.grid(row=0, column=0, padx=5, pady=5)

entry_file1 = tk.Entry(root, width=50)
entry_file1.grid(row=0, column=1, padx=5, pady=5)

button_browse1 = tk.Button(root, text="Browse", command=select_file1)
button_browse1.grid(row=0, column=2, padx=5, pady=5)

label_file2 = tk.Label(root, text="Select the second Excel file:")
label_file2.grid(row=1, column=0, padx=5, pady=5)

entry_file2 = tk.Entry(root, width=50)
entry_file2.grid(row=1, column=1, padx=5, pady=5)

button_browse2 = tk.Button(root, text="Browse", command=select_file2)
button_browse2.grid(row=1, column=2, padx=5, pady=5)

label_field = tk.Label(root, text="Enter the field name:")
label_field.grid(row=2, column=0, padx=5, pady=5)

entry_field = tk.Entry(root, width=30)
entry_field.grid(row=2, column=1, padx=5, pady=5)

button_compare = tk.Button(root, text="Compare", command=compare_files)
button_compare.grid(row=3, column=1, padx=5, pady=5)

label_result = tk.Label(root, text="Duplicate Names:")
label_result.grid(row=4, column=0, padx=5, pady=5, sticky='w')

duplicate_names_text = tk.Text(root, height=10, width=50)
duplicate_names_text.grid(row=4, column=1, padx=5, pady=5)

root.mainloop()