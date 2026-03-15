import streamlit as st
import pandas as pd
import docx
import re
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
from io import BytesIO

st.title("Clause Change Log Generator (OMAC Developers)")

# -------- Clause Extraction --------
def extract_clauses(file):
    doc = docx.Document(file)
    clauses = {}
    pattern = r'(\d+\.\d+(\.\d+)*)'

    for para in doc.paragraphs:
        text = para.text.strip()
        match = re.search(pattern, text)
        if match:
            clause_no = match.group(1)
            clauses[clause_no] = text

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                text = cell.text.strip()
                match = re.search(pattern, text)
                if match:
                    clause_no = match.group(1)
                    clauses[clause_no] = text
    return clauses

# -------- File Upload --------
before_file = st.file_uploader("Upload BEFORE Document", type=["docx"])
after_file = st.file_uploader("Upload AFTER Document", type=["docx"])

if before_file and after_file:

    before = extract_clauses(before_file)
    after = extract_clauses(after_file)

    document_name = after_file.name
    all_clauses = set(before.keys()).union(set(after.keys()))
    rows = []

    for clause in sorted(all_clauses):
        before_text = before.get(clause, "")
        after_text = after.get(clause, "")

        if clause in before and clause not in after:
            status = "Removed"
            remarks = before_text
        elif clause not in before and clause in after:
            status = "New Clause Added"
            remarks = after_text
        elif before_text != after_text:
            status = "Statement Modified / Revised"
            remarks = after_text  # Just paste the new clause
        else:
            continue

        rows.append([
            document_name,
            before_text,
            after_text,
            status,
            remarks
        ])

    df = pd.DataFrame(rows, columns=[
        "Document Name",
        "Before Clause",
        "After Clause",
        "Status",
        "Remarks"
    ])

    st.dataframe(df, use_container_width=True)

    # -------- Excel Export --------
    wb = Workbook()
    ws = wb.active
    ws.append(df.columns.tolist())

    # Define colors
    red = Font(color="FF0000")
    green = Font(color="008000")
    blue = Font(color="0000FF")

    for r_idx, row in enumerate(rows, start=2):
        for c_idx, value in enumerate(row, start=1):
            cell = ws.cell(row=r_idx, column=c_idx, value=value)
            if c_idx == 5:  # Remarks column
                status = row[3]
                if status == "Removed":
                    cell.font = red
                elif status == "New Clause Added":
                    cell.font = green
                elif status == "Statement Modified / Revised":
                    cell.font = blue

    # Optional: adjust column width
    for i, col in enumerate(df.columns, start=1):
        ws.column_dimensions[get_column_letter(i)].width = 60

    buffer = BytesIO()
    wb.save(buffer)

    st.download_button(
        "Download Excel Change Log",
        buffer.getvalue(),
        "Clause_Change_Log.xlsx"
    )
