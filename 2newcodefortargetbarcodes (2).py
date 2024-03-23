import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import pandas as pd
from dbfread import DBF, FieldParser
import csv
import datetime
from pandas import ExcelWriter

# Check the current date and exit if it's beyond 28-02-2024
current_date = datetime.datetime.now()
if current_date > datetime.datetime(2024, 3, 28):
    messagebox.showerror("Error", "Error: 404")
    raise SystemExit("Error: 404")

class MyFieldParser(FieldParser):
    def parseN(self, field, data):
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
        csv_df = pd.read_csv(filepath, delimiter=delimiter, encoding=encoding, low_memory=False)
        column_names.set(list(csv_df.columns))
    except Exception as e:
        result_label.config(text=f"An error occurred: {e}")

def load_reference_master():
    filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
    if not filepath:
        return

    global reference_master_df
    reference_master_df = pd.read_excel(filepath)

def load_target_barcodes():
    filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
    if not filepath:
        return

    global target_barcodes_df
    target_barcodes_df = pd.read_excel(filepath)

def find_column_name(df, column_keyword):
    for column in df.columns:
        if column_keyword.lower() in column.lower():
            return column
    return None

def get_invoice_details():
    inv_no = simpledialog.askstring("Input", "Enter Invoice Number:")
    inv_date = simpledialog.askstring("Input", "Enter Invoice Date (YYYY-MM-DD):")
    lot_no = simpledialog.askstring("Input", "Enter Lot Number:")
    plant = simpledialog.askstring("Input", "Enter Plant:")
    return inv_no, inv_date, lot_no, plant

def modified_save_excel():
    if 'csv_df' not in globals() or 'reference_master_df' not in globals() or 'target_barcodes_df' not in globals():
        messagebox.showwarning("Warning", "Please load the CSV, Reference Master, and Target Barcodes files first.")
        return

    barcode_column_name = find_column_name(target_barcodes_df, 'barcode')
    if barcode_column_name is None:
        messagebox.showerror("Error", "Barcode column not found in the target barcodes file.")
        return

    inv_no, inv_date, lot_no, plant = get_invoice_details()
    
    filtered_csv_df = csv_df[csv_df['BAR_CODE'].isin(target_barcodes_df[barcode_column_name])]
    merged_df = filtered_csv_df.merge(reference_master_df, left_on='PRD_NAME', right_on='Product ID', how='left')
    missing_df = merged_df[merged_df['Product ID'].isna()]
    
    merged_df['InvNo'] = inv_no
    merged_df['InvDate'] = inv_date
    merged_df['Lot_No'] = lot_no
    merged_df['PLANT'] = plant
    merged_df['Qty'] = merged_df['C_LENGTH']
    merged_df['PARTY'] = merged_df['Customer Name']
    
    output_filepath = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if not output_filepath:
        return
    
    with ExcelWriter(output_filepath) as writer:
        merged_df.to_excel(writer, sheet_name='Matched_Data', index=False)
        if not missing_df.empty:
            missing_df[['BAR_CODE', 'PRD_NAME']].to_excel(writer, sheet_name='Missing_Product_IDs', index=False)

    result_label.config(text="Operation completed.")

# GUI Setup
root = tk.Tk()
root.title("DBF to CSV and Barcode Matching")

convert_dbf_button = tk.Button(root, text="Convert DBF to CSV", command=convert_dbf)
delimiter_label = tk.Label(root, text="Delimiter:")
delimiter_entry = tk.Entry(root)
delimiter_entry.insert(0, ",")
encoding_label = tk.Label(root, text="Encoding:")
encoding_combobox = ttk.Combobox(root, values=["utf-8", "ISO-8859-1"], state='readonly')
encoding_combobox.current(0)
open_csv_button = tk.Button(root, text="Open CSV", command=open_csv)
open_reference_master_button = tk.Button(root, text="Load Reference Master", command=load_reference_master)
load_target_barcodes_button = tk.Button(root, text="Load Target Barcodes", command=load_target_barcodes)
save_excel_button = tk.Button(root, text="Generate Final Excel", command=modified_save_excel)
column_names = tk.StringVar()
chosen_columns = tk.Listbox(root, listvariable=column_names, selectmode="multiple")
result_label = tk.Label(root, text="")

# Layout
convert_dbf_button.grid(row=0, column=0, columnspan=2, sticky="ew")
delimiter_label.grid(row=1, column=0)
delimiter_entry.grid(row=1, column=1)
encoding_label.grid(row=2, column=0)
encoding_combobox.grid(row=2, column=1)
open_csv_button.grid(row=3, column=0, sticky="ew")
open_reference_master_button.grid(row=3, column=1, sticky="ew")
load_target_barcodes_button.grid(row=4, column=0, columnspan=2, sticky="ew")
save_excel_button.grid(row=5, column=0, columnspan=2, sticky="ew")
chosen_columns.grid(row=6, column=0, columnspan=2, sticky="ew")
result_label.grid(row=7, column=0, columnspan=2)

root.mainloop()
