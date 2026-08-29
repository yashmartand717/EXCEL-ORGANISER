import pandas as pd
import re
import tkinter as tk
from tkinter import filedialog, messagebox
import os

def normalize_column_name(col):
    return re.sub(r'[^a-z0-9]', '', str(col).lower())

def process_excel(input_file):
    try:
        # Read all sheets from the excel file
        all_sheets = pd.read_excel(input_file, sheet_name=None, header=None)
    except Exception as e:
        return False, f"Error reading input file:\n{e}", None

    target_columns = [
        "Sr.No.",
        "Reg No",
        "Class Name",
        "Student Name",
        "Father's Name",
        "Mother Name",
        "Gender",
        "D.O.B",
        "Mother Mobile Num",
        "Email"
    ]
    
    column_mappings = {
        "Reg No": ["regno", "admno", "admissionno", "admissionnumber", "registrationnumber", "registrationno", "admnno", "adm", "enrollmentno", "enrollno", "rollno"],
        "Class Name": ["classname", "class", "grade", "standard", "classsec", "section"],
        "Student Name": ["studentname", "name", "nameofstudent", "student"],
        "Father's Name": ["fathersname", "fathername", "fname"],
        "Mother Name": ["mothername", "mothersname", "mname"],
        "Gender": ["gender", "sex"],
        "D.O.B": ["dob", "dateofbirth", "birthdate"],
        "Mother Mobile Num": ["mothermobilenum", "mothermobileno", "mothermobile", "mobileno", "phone", "phonenumber", "contactno", "fathermobileno", "contact", "mobile"],
        "Email": ["email", "emailid", "mailid", "mail"]
    }

    mandatory_aliases = set(column_mappings["Reg No"] + column_mappings["Class Name"] + column_mappings["Student Name"])
    
    all_processed_dfs = {}
    missing_mandatory_global = set()

    for sheet_name, df_raw in all_sheets.items():
        if df_raw.empty:
            continue
            
        best_row_idx = 0
        max_matches = 0
        
        # Check the first 20 rows to find headers in this sheet
        for i in range(min(20, len(df_raw))):
            row_values = df_raw.iloc[i].astype(str).tolist()
            normalized_row = [normalize_column_name(val) for val in row_values]
            
            matches = sum(1 for val in normalized_row if val in mandatory_aliases)
                    
            if matches > max_matches:
                max_matches = matches
                best_row_idx = i
                
        # Set the actual headers and data
        df = df_raw.copy()
        if max_matches > 0:
            new_cols = df_raw.iloc[best_row_idx].values
            df = df_raw.iloc[best_row_idx+1:].reset_index(drop=True)
            df.columns = new_cols
        else:
            # Fallback to row 0 if we somehow couldn't find any matches
            new_cols = df_raw.iloc[0].values
            df = df_raw.iloc[1:].reset_index(drop=True)
            df.columns = new_cols

        input_cols = {normalize_column_name(col): col for col in df.columns}
        mapped_data = {}
        for target_key, aliases in column_mappings.items():
            found_col = None
            for alias in aliases:
                if alias in input_cols:
                    found_col = input_cols[alias]
                    break
            
            if found_col:
                mapped_data[target_key] = df[found_col]

        output_df = pd.DataFrame(mapped_data)
        
        # Ensure all target columns exist
        for col in target_columns:
            if col not in output_df.columns:
                output_df[col] = "" 
                
        output_df = output_df[target_columns]
        
        # Drop completely empty rows where Name and Reg No are missing
        output_df.dropna(subset=["Reg No", "Student Name"], how='all', inplace=True)
        
        # Track missing mandatory columns for this sheet
        for col in ["Reg No", "Class Name", "Student Name"]:
            if col not in mapped_data:
                missing_mandatory_global.add(f"'{col}' in sheet '{sheet_name}'")
                
        # Auto-generate Serial Number for the sheet
        output_df['Sr.No.'] = range(1, len(output_df) + 1)
                
        all_processed_dfs[sheet_name] = output_df

    if not all_processed_dfs:
        return False, "No data found in any sheets.", None
        
    if missing_mandatory_global:
        warning_msg = f"Warning: Mandatory column(s) not found:\n{', '.join(missing_mandatory_global)}\nThey were left empty."
    else:
        warning_msg = None

    # Auto generate output file name in the same folder
    output_file = os.path.splitext(input_file)[0] + "_formatted.xlsx"
    
    try:
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            for sheet_name, final_df in all_processed_dfs.items():
                # Excel sheet names can be at most 31 characters long
                final_df.to_excel(writer, sheet_name=str(sheet_name)[:31], index=False)
        return True, warning_msg, output_file
    except Exception as e:
        return False, f"Error saving output file:\n{e}", None

def select_input():
    filename = filedialog.askopenfilename(
        title="Select Input Excel",
        filetypes=[("Excel Files", "*.xlsx *.xls")]
    )
    if filename:
        input_var.set(filename)

def run_conversion():
    input_path = input_var.get()
    
    if not input_path:
        messagebox.showwarning("Missing Info", "Please select an input file.")
        return
        
    if not os.path.exists(input_path):
        messagebox.showerror("Error", "Selected input file does not exist.")
        return
        
    success, msg, out_file = process_excel(input_path)
    
    if success:
        if msg: 
            messagebox.showwarning("Success with Warnings", f"Saved as:\n{out_file}\n\n{msg}")
        else:
            messagebox.showinfo("Success", f"Conversion completed successfully!\n\nSaved as:\n{out_file}")
    else:
        messagebox.showerror("Error", msg)

def create_gui():
    app = tk.Tk()
    app.title("Excel Organiser Tool")
    app.geometry("450x120")
    app.resizable(False, False)

    global input_var
    input_var = tk.StringVar()

    tk.Label(app, text="Input Excel File:").place(x=20, y=20)
    tk.Entry(app, textvariable=input_var, width=40).place(x=120, y=20)
    tk.Button(app, text="Browse", command=select_input).place(x=370, y=15)

    tk.Button(app, text="Convert & Organize", command=run_conversion, bg="green", fg="white", font=("Arial", 10, "bold")).place(x=150, y=70)

    app.mainloop()

if __name__ == "__main__":
    create_gui()
