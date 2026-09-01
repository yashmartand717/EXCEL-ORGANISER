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
        
        # Clean and format D.O.B column
        def clean_dob(val):
            if pd.isna(val):
                return ""
            if hasattr(val, 'strftime'):
                return val.strftime("%d/%m/%Y")
            val_str = str(val).strip()
            if val_str.startswith("'"):
                val_str = val_str[1:]
            if val_str.lower() in ('nan', 'nat', 'none', 'null', ''):
                return ""
            try:
                dt = pd.to_datetime(val_str, dayfirst=True)
                return dt.strftime("%d/%m/%Y")
            except:
                return val_str

        output_df["D.O.B"] = output_df["D.O.B"].apply(clean_dob)
                
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

from PIL import Image
import io

def extract_teacher_dfs(all_sheets, school_name):
    def normalize_column_name(col):
        return re.sub(r'[^a-z0-9]', '', str(col).lower())
    
    target_aliases = {
        "name": ["name", "teachername"],
        "subject": ["subject", "subjects"],
        "classes": ["takingclasses", "classes", "takingclass"],
        "mobile": ["contactno", "mobile", "mobileno", "phone"]
    }
    
    mandatory_aliases = set(target_aliases["name"] + target_aliases["classes"])
    
    subject_map = {
        "sst": "Social Studies",
        "science": "Science",
        "maths": "Maths",
        "english": "English",
        "hindi": "Hindi",
        "punjabi": "Punjabi",
        "physics": "Physics",
        "chemistry": "Chemistry",
        "chem": "Chemistry",
        "phy": "Physics",
        "sci": "Science",
        "accounts": "Accountancy",
        "accountancy": "Accountancy",
        "eco": "Economics",
        "business": "Business Studies",
        "ai": "AI",
        "it": "IT",
        "music": "Music",
        "painting": "Painting",
        "sports": "Sports",
        "physical education": "Physical Education"
    }
    
    def map_subject(s):
        s_lower = str(s).strip().lower()
        if s_lower in subject_map:
            return subject_map[s_lower]
        for k, v in subject_map.items():
            if k == s_lower:
                return v
        return str(s).strip().title()
    
    def format_class_name(cls_str):
        cls_str = cls_str.strip().upper()
        m = re.match(r'^(\d+)\s*([A-Z])?$', cls_str)
        if m:
            num = m.group(1)
            sec = m.group(2)
            if sec:
                return f"{num}-{sec}"
            else:
                return f"{num}-A"
        return cls_str
    
    all_processed_dfs = {}
    
    for sheet_name, df_raw in all_sheets.items():
        if df_raw.empty:
            continue
            
        best_row_idx = 0
        max_matches = 0
        
        for i in range(min(20, len(df_raw))):
            row_values = df_raw.iloc[i].astype(str).tolist()
            normalized_row = [normalize_column_name(val) for val in row_values]
            matches = sum(1 for val in normalized_row if val in mandatory_aliases)
            if matches > max_matches:
                max_matches = matches
                best_row_idx = i
                
        if max_matches == 0:
            continue
            
        new_cols = df_raw.iloc[best_row_idx].values
        df = df_raw.iloc[best_row_idx+1:].reset_index(drop=True)
        df.columns = new_cols
        
        input_cols = {normalize_column_name(col): col for col in df.columns}
        
        name_col = next((input_cols[a] for a in target_aliases["name"] if a in input_cols), None)
        subj_col = next((input_cols[a] for a in target_aliases["subject"] if a in input_cols), None)
        classes_col = next((input_cols[a] for a in target_aliases["classes"] if a in input_cols), None)
        mobile_col = next((input_cols[a] for a in target_aliases["mobile"] if a in input_cols), None)
        
        if not name_col or not classes_col:
            continue
            
        output_rows = []
        for idx, row in df.iterrows():
            name_val = row[name_col] if pd.notna(row[name_col]) else ""
            if str(name_val).strip() == "" or str(name_val).strip().lower() == "name":
                continue
                
            name_clean = re.sub(r'^(Mr\.|Ms\.|Mrs\.|Dr\.)\s*', '', str(name_val).strip(), flags=re.IGNORECASE)
            parts = name_clean.split(maxsplit=1)
            first_name = parts[0] if len(parts) > 0 else ""
            last_name = parts[1] if len(parts) > 1 else ""
            
            mobile = row[mobile_col] if mobile_col and pd.notna(row[mobile_col]) else ""
            if pd.isna(mobile): mobile = ""
            
            main_subject = row[subj_col] if subj_col and pd.notna(row[subj_col]) else ""
            main_subject = map_subject(main_subject)
            
            classes_str = row[classes_col] if pd.notna(row[classes_col]) else ""
            tokens = [t.strip() for t in re.split(r'[,]', str(classes_str))]
            
            for token in tokens:
                if not token: continue
                
                match = re.match(r'^([^(]+)(?:\(([^)]+)\))?$', token)
                if match:
                    c_name = match.group(1).strip()
                    sub_override = match.group(2)
                    
                    range_match = re.search(r'(\d+)\s*TO\s*(\d+)', c_name, flags=re.IGNORECASE)
                    classes_to_emit = []
                    if range_match:
                        start = int(range_match.group(1))
                        end = int(range_match.group(2))
                        for c_num in range(start, end + 1):
                            classes_to_emit.append(f"{c_num}-A")
                    else:
                        classes_to_emit.append(format_class_name(c_name))
                    
                    for emit_class in classes_to_emit:
                        final_subj = main_subject
                        if sub_override:
                            final_subj = map_subject(sub_override)
                            
                        output_rows.append({
                            "first_name": first_name.title(),
                            "last_name": last_name.title(),
                            "school": school_name,
                            "class_name": emit_class,
                            "subject": final_subj,
                            "Mobile no": mobile
                        })
        
        if output_rows:
            df_out = pd.DataFrame(output_rows, columns=["first_name", "last_name", "school", "class_name", "subject", "Mobile no"])
            df_out = df_out.groupby(
                ["first_name", "last_name", "school", "subject", "Mobile no"], 
                as_index=False, dropna=False
            ).agg({
                "class_name": lambda x: ", ".join(x.unique())
            })
            df_out = df_out[["first_name", "last_name", "school", "class_name", "subject", "Mobile no"]]
            all_processed_dfs[sheet_name] = df_out
            
    return all_processed_dfs, None

def process_teacher_excel(input_file, school_name):
    try:
        all_sheets = pd.read_excel(input_file, sheet_name=None, header=None)
    except Exception as e:
        return False, f"Error reading input file:\n{e}", None
        
    all_processed_dfs, err = extract_teacher_dfs(all_sheets, school_name)
    
    if not all_processed_dfs:
        return False, "No valid teacher data found in any sheets.", None
        
    output_file = os.path.splitext(input_file)[0] + "_formatted_teachers.xlsx"
    try:
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            for sheet_name, final_df in all_processed_dfs.items():
                final_df.to_excel(writer, sheet_name=str(sheet_name)[:31], index=False)
        return True, None, output_file
    except Exception as e:
        return False, f"Error saving output file:\n{e}", None

import os
from dotenv import load_dotenv

load_dotenv()

def process_teacher_image(input_image, school_name):
    try:
        from openai import OpenAI
        import base64
        
        api_key = os.environ.get("OPENAI_API_KEY")
        if not api_key:
            return False, "OpenAI API Key not found. Please add it to your .env file.", None
            
        client = OpenAI(
            api_key=api_key,
            default_headers={"Accept-Encoding": "identity"}
        )
        
        with open(input_image, "rb") as f:
            base64_image = base64.b64encode(f.read()).decode('utf-8')
        
        prompt = "Extract the table from this image and output it as a valid CSV format exactly matching the data in the image. Do not include any other text or markdown formatting, just the CSV. Use the pipe character '|' as the delimiter instead of commas, and newlines for rows. Make sure every row has the exact same number of columns."
        
        response = client.chat.completions.create(
            model="gpt-4o",
            messages=[
                {
                    "role": "user",
                    "content": [
                        {"type": "text", "text": prompt},
                        {
                            "type": "image_url",
                            "image_url": {
                                "url": f"data:image/jpeg;base64,{base64_image}"
                            }
                        }
                    ]
                }
            ],
            max_tokens=2000,
        )
        csv_text = response.choices[0].message.content.strip()
        
        # Remove markdown code block if present
        if csv_text.startswith("```csv"):
            csv_text = csv_text[6:]
        if csv_text.startswith("```"):
            csv_text = csv_text[3:]
        if csv_text.endswith("```"):
            csv_text = csv_text[:-3]
            
        csv_text = csv_text.strip()
        df_raw = pd.read_csv(io.StringIO(csv_text), sep='|', header=None, skipinitialspace=True)
        
        all_sheets = {"Sheet1": df_raw}
        all_processed_dfs, err = extract_teacher_dfs(all_sheets, school_name)
        
        if not all_processed_dfs:
            return False, "Could not extract valid teacher data from the image.", None
            
        output_file = os.path.splitext(input_image)[0] + "_formatted_teachers.xlsx"
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            for sheet_name, final_df in all_processed_dfs.items():
                final_df.to_excel(writer, sheet_name=str(sheet_name)[:31], index=False)
                
        return True, None, output_file
        
    except Exception as e:
        return False, f"Error processing image with OpenAI API:\n{e}", None


def select_input():
    filename = filedialog.askopenfilename(
        title="Select Input File",
        filetypes=[("Excel or Image Files", "*.xlsx *.xls *.png *.jpg *.jpeg"), ("All Files", "*.*")]
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
        
    mode = mode_var.get()
    is_image = input_path.lower().endswith(('.png', '.jpg', '.jpeg'))
    
    if mode == "student":
        if is_image:
            messagebox.showerror("Error", "Image input is only supported for Teacher Data mode.")
            return
        success, msg, out_file = process_excel(input_path)
    else:
        school_name = school_var.get().strip()
        if is_image:
            success, msg, out_file = process_teacher_image(input_path, school_name)
        else:
            success, msg, out_file = process_teacher_excel(input_path, school_name)
    
    if success:
        if msg: 
            messagebox.showwarning("Success with Warnings", f"Saved as:\n{out_file}\n\n{msg}")
        else:
            messagebox.showinfo("Success", f"Conversion completed successfully!\n\nSaved as:\n{out_file}")
    else:
        messagebox.showerror("Error", msg)

def on_mode_change():
    if mode_var.get() == "teacher":
        school_label.place(x=20, y=90)
        school_entry.place(x=120, y=90)
        convert_btn.place(x=150, y=130)
        app.geometry("450x180")
    else:
        school_label.place_forget()
        school_entry.place_forget()
        convert_btn.place(x=150, y=130)
        app.geometry("450x180")

def create_gui():
    global app
    app = tk.Tk()
    app.title("Excel Organiser Tool")
    app.geometry("450x180")
    app.resizable(False, False)

    global input_var, mode_var, school_var
    global school_label, school_entry, convert_btn
    
    input_var = tk.StringVar()
    mode_var = tk.StringVar(value="student")
    school_var = tk.StringVar(value="Prime Steps International School")

    tk.Label(app, text="Input File:").place(x=20, y=20)
    tk.Entry(app, textvariable=input_var, width=40).place(x=120, y=20)
    tk.Button(app, text="Browse", command=select_input).place(x=370, y=15)

    tk.Label(app, text="Mode:").place(x=20, y=55)
    tk.Radiobutton(app, text="Student Data", variable=mode_var, value="student", command=on_mode_change).place(x=120, y=55)
    tk.Radiobutton(app, text="Teacher Data", variable=mode_var, value="teacher", command=on_mode_change).place(x=220, y=55)

    school_label = tk.Label(app, text="School Name:")
    school_entry = tk.Entry(app, textvariable=school_var, width=40)
    
    convert_btn = tk.Button(app, text="Convert & Organize", command=run_conversion, bg="green", fg="white", font=("Arial", 10, "bold"))
    
    # Hidden initially
    on_mode_change()

    app.mainloop()

if __name__ == "__main__":
    create_gui()
