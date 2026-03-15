import streamlit as st
import pandas as pd
import docx
import re
from openpyxl import Workbook
from openpyxl.styles import Font
from io import BytesIO

st.title("Document Clause Change Log Generator")

# -------- Extract Clauses -------- #

def extract_clauses(file):

    doc = docx.Document(file)
    clauses = {}

    pattern = r'^(\d+(\.\d+)+)'

    for para in doc.paragraphs:
        text = para.text.strip()

        match = re.match(pattern, text)

        if match:
            clause_no = match.group(1)
            clauses[clause_no] = text

    return clauses


# -------- Upload Files -------- #

before_file = st.file_uploader("Upload BEFORE Document", type=["docx"])
after_file = st.file_uploader("Upload AFTER Document", type=["docx"])

if before_file and after_file:

    before = extract_clauses(before_file)
    after = extract_clauses(after_file)

    document_name = after_file.name

    all_clauses = set(before.keys()).union(set(after.keys()))

    data = []

    for clause in sorted(all_clauses):

        before_text = before.get(clause, "")
        after_text = after.get(clause, "")

        if clause in before and clause not in after:
            status = "Removed"

        elif clause not in before and clause in after:
            status = "New Clause Added"

        elif before_text != after_text:
            status = "Statement Modified / Revised"

        else:
            continue

        data.append([
            document_name,
            before_text,
            after_text,
            status
        ])

    df = pd.DataFrame(data, columns=[
        "Document Name",
        "Before Clause",
        "After Clause",
        "Status"
    ])

    st.dataframe(df, use_container_width=True)

# -------- Excel Export -------- #

    wb = Workbook()
    ws = wb.active

    ws.append(df.columns.tolist())

    red = Font(color="FF0000")
    green = Font(color="008000")
    blue = Font(color="0000FF")

    for row in data:
        ws.append(row)

        status_cell = ws.cell(row=ws.max_row, column=4)

        if row[3] == "Removed":
            status_cell.font = red

        elif row[3] == "New Clause Added":
            status_cell.font = green

        elif row[3] == "Statement Modified / Revised":
            status_cell.font = blue


    buffer = BytesIO()
    wb.save(buffer)

    st.download_button(
        "Download Excel File",
        buffer.getvalue(),
        "Clause_Change_Log.xlsx"
    )
