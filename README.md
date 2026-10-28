# Excel Organiser

Converts messy school data spreadsheets into a single standard format. The tool reads every sheet and auto-matches dozens of column-name variants (for example *Adm No*, *Enrollment No* and *Roll No* all become **Reg No**).

## Output columns
Sr.No. · Reg No · Class Name · Student Name · Father's Name · Mother Name · Gender · D.O.B · Mother Mobile Num · Email

## Two ways to use it
- **Web app:** `streamlit run app.py`
- **Desktop app:** `python excel_organiser.py` (Tkinter file picker)

## Tech stack
Python · pandas · Streamlit · Tkinter

## Setup
```bash
pip install -r requirements.txt
```
