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

from PIL import Image

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

def process_teacher_excel_to_bytes(uploaded_file, school_name):
    try:
        all_sheets = pd.read_excel(uploaded_file, sheet_name=None, header=None)
    except Exception as e:
        st.error(f"Error reading input file: {e}")
        return None, None
        
    all_processed_dfs, err = extract_teacher_dfs(all_sheets, school_name)
    if not all_processed_dfs:
        return None, "No valid teacher data found in any sheets."
        
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for sheet_name, final_df in all_processed_dfs.items():
            final_df.to_excel(writer, sheet_name=str(sheet_name)[:31], index=False)
    processed_data = output.getvalue()
    
    return processed_data, None

import os
from dotenv import load_dotenv

load_dotenv()

def process_teacher_image_to_bytes(uploaded_image, school_name):
    try:
        from openai import OpenAI
        import base64
        
        api_key = os.environ.get("OPENAI_API_KEY")
        if not api_key:
            st.error("OpenAI API Key not found. Please add it to your .env file.")
            return None, None
            
        client = OpenAI(
            api_key=api_key,
            default_headers={"Accept-Encoding": "identity"}
        )
        
        # Streamlit UploadedFile has a getvalue() method
        base64_image = base64.b64encode(uploaded_image.getvalue()).decode('utf-8')
        
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
            return None, "Could not extract valid teacher data from the image."
            
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            for sheet_name, final_df in all_processed_dfs.items():
                final_df.to_excel(writer, sheet_name=str(sheet_name)[:31], index=False)
        processed_data = output.getvalue()
        
        return processed_data, None
        
    except Exception as e:
        st.error(f"Error processing image with OpenAI API: {e}")
        return None, None

st.set_page_config(page_title="Excel Organiser", layout="centered")
st.title("Excel Organiser Tool 📊")
st.markdown("Upload your raw school data to convert it into the standardized format.")

mode = st.radio("Select Processing Mode", ["Student Data", "Teacher Data"])

school_name = ""
if mode == "Teacher Data":
    school_name = st.text_input("School Name", value="Prime Steps International School")

if mode == "Student Data":
    uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx", "xls"])
else:
    uploaded_file = st.file_uploader("Upload Excel File or Image", type=["xlsx", "xls", "png", "jpg", "jpeg"])

if uploaded_file is not None:
    st.info(f"File uploaded: {uploaded_file.name}")
    is_image = uploaded_file.name.lower().endswith(('png', 'jpg', 'jpeg'))
    
    if st.button("Process File"):
        with st.spinner("Processing..."):
            if mode == "Student Data":
                processed_bytes, warning = process_excel_to_bytes(uploaded_file)
            else:
                if is_image:
                    processed_bytes, warning = process_teacher_image_to_bytes(uploaded_file, school_name)
                else:
                    processed_bytes, warning = process_teacher_excel_to_bytes(uploaded_file, school_name)
            
            if processed_bytes:
                st.success("File processed successfully!")
                if warning:
                    st.warning(warning)
                
                # Auto-generate download filename
                suffix = "_formatted_students.xlsx" if mode == "Student Data" else "_formatted_teachers.xlsx"
                out_name = uploaded_file.name.rsplit('.', 1)[0] + suffix
                
                st.download_button(
                    label="Download Formatted Excel",
                    data=processed_bytes,
                    file_name=out_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
                st.markdown("### Preview")
                try:
                    preview_dfs = pd.read_excel(io.BytesIO(processed_bytes), sheet_name=None)
                    for sheet_name, df in preview_dfs.items():
                        with st.expander(f"Sheet: {sheet_name}", expanded=True):
                            output = io.BytesIO()
                            with pd.ExcelWriter(output) as writer:
                                df.to_excel(writer, sheet_name=sheet_name, index=False)
                            sheet_bytes = output.getvalue()
                            
                            st.download_button(
                                label=f"Download {sheet_name}.xlsx",
                                data=sheet_bytes,
                                file_name=f"{sheet_name}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key=f"download_{sheet_name}"
                            )
                            st.dataframe(df, use_container_width=True)
                except Exception as e:
                    st.error(f"Could not load preview: {e}")
