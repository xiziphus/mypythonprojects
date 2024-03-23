import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from dbfread import DBF, FieldParser
import csv
import datetime
from pandas import ExcelWriter

# Custom Field Parser for DBF
class MyFieldParser(FieldParser):
    def parseN(self, field, data):
        # Special handling for the 'BAR_CODE' field
        if field.name == 'BAR_CODE':
            try:
                return int(data.strip() or '0')
            except ValueError:
                return None
        else:
            try:
                return float(data.strip() or '0')
            except ValueError:
                return None

    def parseD(self, field, data):
        try:
            return datetime.date(int(data[:4]), int(data[4:6]), int(data[6:8]))
        except ValueError:
            return None

def convert_dbf():
    dbf_path = filedialog.askopenfilename(filetypes=[("DBF files", "*.dbf"), ("All files", "*.*")])
    if not dbf_path:
        return

    csv_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
    if not csv_path:
        return

    try:
        with open(csv_path, 'w', newline='') as csvfile:
            writer = csv.writer(csvfile)
            dbf = DBF(dbf_path, parserclass=MyFieldParser)
            for i, record in enumerate(dbf):
                if i == 0:
                    writer.writerow(record.keys())
                writer.writerow(record.values())
        messagebox.showinfo("Success", "DBF to CSV conversion completed.")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

def open_csv():
    filepath = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv"), ("All files", "*.*")])
    if not filepath:
        return

    delimiter = delimiter_entry.get()
    encoding = encoding_combobox.get()
    
    try:
        global csv_df
        csv_df = pd.read_csv(filepath, delimiter=delimiter, encoding=encoding)
        column_names.set(list(csv_df.columns))
    except Exception as e:
        result_label.config(text=f"An error occurred: {e}")

def open_excel():
    filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
    if not filepath:
        return

    global excel_df
    excel_df = pd.read_excel(filepath)
    
def save_excel():
    if 'csv_df' not in globals() or 'excel_df' not in globals():
        messagebox.showwarning("Warning", "Please load both CSV and Excel files first.")
        return

    selected_column_names = [csv_df.columns[idx] for idx in list(chosen_columns.curselection())]
    output_filepath = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if not output_filepath:
        return

    matched_df = csv_df[csv_df['BAR_CODE'].isin(excel_df['Target_Barcodes'])][selected_column_names]
    missing_barcodes = set(excel_df['Target_Barcodes']) - set(csv_df['BAR_CODE'])
    
    with ExcelWriter(output_filepath) as writer:
        matched_df.to_excel(writer, sheet_name='Matched_Barcodes', index=False)
        if missing_barcodes:
            missing_df = pd.DataFrame(list(missing_barcodes), columns=['Missing_Barcodes'])
            missing_df.to_excel(writer, sheet_name='Missing_Barcodes', index=False)

    result_label.config(text=f"Operation completed. Missing barcodes: {len(missing_barcodes)}")

# GUI Setup
root = tk.Tk()
root.title("DBF to CSV and Barcode Matching")

# Widgets for DBF to CSV conversion
convert_dbf_button = tk.Button(root, text="Convert DBF to CSV", command=convert_dbf)

# Widgets for CSV and Excel operations
delimiter_label = tk.Label(root, text="Delimiter:")
delimiter_entry = tk.Entry(root)
delimiter_entry.insert(0, ",")
encoding_label = tk.Label(root, text="Encoding:")
encoding_combobox = ttk.Combobox(root, values=["utf-8", "ISO-8859-1"], state='readonly')
encoding_combobox.current(0)

open_csv_button = tk.Button(root, text="Open CSV", command=open_csv)
open_excel_button = tk.Button(root, text="Open Excel with Barcodes", command=open_excel)
save_excel_button = tk.Button(root, text="Save Excel", command=save_excel)

column_names = tk.StringVar()
chosen_columns = tk.Listbox(root, listvariable=column_names, selectmode="multiple")
result_label = tk.Label(root, text="")

# Layout
convert_dbf_button.grid(row=0, column=0, columnspan=2)

delimiter_label.grid(row=1, column=0)
delimiter_entry.grid(row=1, column=1)
encoding_label.grid(row=2, column=0)
encoding_combobox.grid(row=2, column=1)
open_csv_button.grid(row=3, column=0)
open_excel_button.grid(row=3, column=1)
save_excel_button.grid(row=4, column=0, columnspan=2)
chosen_columns.grid(row=5, column=0, columnspan=2)
result_label.grid(row=6, column=0, columnspan=2)

root.mainloop()
