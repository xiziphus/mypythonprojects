import tkinter as tk
from tkinter import scrolledtext, messagebox, simpledialog, ttk
import pandas as pd
import pyodbc
import io
import datetime
from tkinter import filedialog

def select_database():
    global mdb_path
    mdb_path = filedialog.askopenfilename(title="Select MDB Database",
                                          filetypes=(("MDB files", "*.mdb"), ("ACCDB files", "*.accdb"), ("All files", "*.*")))
    if mdb_path:
        print(f"Selected database: {mdb_path}")
    else:
        print("No file selected.")

# MDB Database Path
mdb_path = ""

# List of parties for dropdown
party_list = [
    "DHOOT TRANSMISSION PVT LTD (1112)",
    "DTPL HOSUR(1116)",
    "MCPL MANESAR",
    "HTL",
    "APTIV COMPONENTS INDIA PVT LTD (II)",
    "MCPL BHIWADI",
    "KOPERTEK",
    "MCPL DELHI",
    "SEIKODENKI",
    "APTIV HARYANA",
    "INFAC",
    "APTIV KOCHI",
    "MCPL GUJ",
    "Dhoot Transmission Pvt Ltd (1124)",
    "BSA CORPORATION LIMITED",
    "Other"
]

def on_party_select(event):
    # If 'Other' is selected, show input box for custom party name
    if party_var.get() == "Other":
        custom_party = simpledialog.askstring("Input", "Enter custom party name:", parent=root)
        if custom_party:
            party_var.set(custom_party)

def append_to_database(data, lot_no, lot_date, party):
    # Database connection string
    conn_str = (r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
                r'DBQ=' + mdb_path + ';')
    conn = pyodbc.connect(conn_str)
    cursor = conn.cursor()

    # Generate serial number (sno) - this logic might need to be adjusted
    sno = 1

    # Transform and append data to the BoxMaster table
    for index, row in data.iterrows():
        transformed_row = {
            "BOX": 1,
            "LOT_NO": lot_no,
            "Lot_Date": lot_date,
            "PartNo": row['FG PART NO'],
            "Guage": str(row['GAUGE']) if pd.notna(row['GAUGE']) else '1111',
            "TYPE": row['TYPE'],
            "COLOUR": row['COLOR'],
            "NOJ": 1,
            "SpoolNumber": 1,
            "Spool": 1,
            "Qty": row['LENGTH'],
            "TotalQty": row['LENGTH'],
            "BarCode": row['SERIAL NO'],
            "Party": party,
            "sno": sno
        }
        sno += 1  # Increment sno for each row

        try:
            sql = '''INSERT INTO BoxMaster (BOX, LOT_NO, Lot_Date, PartNo, Guage, TYPE, COLOUR, NOJ, SpoolNumber, Spool, Qty, TotalQty, BarCode, Party, sno) 
                     VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)'''
            values = (transformed_row['BOX'], transformed_row['LOT_NO'], transformed_row['Lot_Date'], 
                      transformed_row['PartNo'], transformed_row['Guage'], transformed_row['TYPE'], 
                      transformed_row['COLOUR'], transformed_row['NOJ'], transformed_row['SpoolNumber'], 
                      transformed_row['Spool'], transformed_row['Qty'], transformed_row['TotalQty'], 
                      transformed_row['BarCode'], transformed_row['Party'], transformed_row['sno'])
            cursor.execute(sql, values)
            print(f"Attempting to insert row {sno}: {transformed_row}")
            print(f"Inserted row {sno} successfully.")
        except Exception as e:
            print(f"Error inserting row {sno}: {e}")

    conn.commit()
    cursor.close()
    conn.close()
    messagebox.showinfo("Success", "Data appended successfully to the database.")

def process_data():
    lot_no = simpledialog.askstring("Input", "Enter LOT_NO:", parent=root)
    lot_date_str = simpledialog.askstring("Input", "Enter Lot_Date (dd-mm-yyyy):", parent=root)

    try:
        lot_date = datetime.datetime.strptime(lot_date_str, '%d-%m-%Y').date()
        party = party_var.get()
        data = io.StringIO(txt_input.get("1.0", tk.END))
        df = pd.read_csv(data, sep="\t")  # Adjust the separator if necessary
        append_to_database(df, lot_no, lot_date, party)
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")
        print(f"Data processing error: {e}")

# GUI setup
root = tk.Tk()
root.title("Data Input for MDB Appending")
root.geometry("800x600")

txt_input = scrolledtext.ScrolledText(root, wrap=tk.WORD)
txt_input.grid(row=0, column=0, columnspan=3, padx=10, pady=10, sticky="nsew")

# Dropdown setup in GUI
party_var = tk.StringVar(value=party_list[0])
party_dropdown = ttk.Combobox(root, textvariable=party_var, values=party_list)
party_dropdown.grid(row=1, column=0, padx=10, pady=10, sticky="ew")
party_dropdown.bind("<<ComboboxSelected>>", on_party_select)

btn_process = tk.Button(root, text="Process and Append Data", command=process_data)
btn_process.grid(row=1, column=2, padx=10, pady=10, sticky="ew")
btn_select_db = tk.Button(root, text="Select Database", command=select_database)
btn_select_db.grid(row=2, column=0, padx=10, pady=10, sticky="ew")

root.grid_columnconfigure(0, weight=1)
root.grid_columnconfigure(1, weight=1)
root.grid_columnconfigure(2, weight=1)
root.grid_rowconfigure(0, weight=1)
root.grid_rowconfigure(1, weight=0)

root.mainloop()
