import streamlit as st
import pandas as pd
import re
import io

def normalize_column_name(col):
    return re.sub(r'[^a-z0-9]', '', str(col).lower())

def process_excel_to_bytes(uploaded_file):
    try:
        all_sheets = pd.read_excel(uploaded_file, sheet_name=None, header=None)
    except Exception as e:
        st.error(f"Error reading input file: {e}")
        return None, None

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
        
        for i in range(min(20, len(df_raw))):
            row_values = df_raw.iloc[i].astype(str).tolist()
            normalized_row = [normalize_column_name(val) for val in row_values]
            matches = sum(1 for val in normalized_row if val in mandatory_aliases)
            if matches > max_matches:
                max_matches = matches
                best_row_idx = i
                
        df = df_raw.copy()
        if max_matches > 0:
            new_cols = df_raw.iloc[best_row_idx].values
            df = df_raw.iloc[best_row_idx+1:].reset_index(drop=True)
            df.columns = new_cols
        else:
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
        
        for col in target_columns:
            if col not in output_df.columns:
                output_df[col] = "" 
                
        output_df = output_df[target_columns]
        output_df.dropna(subset=["Reg No", "Student Name"], how='all', inplace=True)
        
        for col in ["Reg No", "Class Name", "Student Name"]:
            if col not in mapped_data:
                missing_mandatory_global.add(f"'{col}' in sheet '{sheet_name}'")
                
        # Auto-generate Serial Number for the sheet
        output_df['Sr.No.'] = range(1, len(output_df) + 1)
                
        all_processed_dfs[sheet_name] = output_df

    if not all_processed_dfs:
        return None, "No data found in any sheets."
        
    warning_msg = None
    if missing_mandatory_global:
        warning_msg = f"Warning: Mandatory column(s) not found:\n{', '.join(missing_mandatory_global)}"

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for sheet_name, final_df in all_processed_dfs.items():
            final_df.to_excel(writer, sheet_name=str(sheet_name)[:31], index=False)
    processed_data = output.getvalue()
    
    return processed_data, warning_msg

st.set_page_config(page_title="Excel Organiser", layout="centered")
st.title("Excel Organiser Tool 📊")
st.markdown("Upload your raw school data Excel file to convert it into the standardized format.")

uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx", "xls"])

if uploaded_file is not None:
    st.info(f"File uploaded: {uploaded_file.name}")
    
    if st.button("Process File"):
        with st.spinner("Processing..."):
            processed_bytes, warning = process_excel_to_bytes(uploaded_file)
            
            if processed_bytes:
                st.success("File processed successfully!")
                if warning:
                    st.warning(warning)
                
                # Auto-generate download filename
                out_name = uploaded_file.name.rsplit('.', 1)[0] + "_formatted.xlsx"
                
                st.download_button(
                    label="Download Formatted Excel",
                    data=processed_bytes,
                    file_name=out_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
