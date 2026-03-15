import streamlit as st
import pandas as pd
import docx
import re
import difflib
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter
from io import BytesIO

st.title("Clause Change Log Generator")

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

# -------- Word-level diff for modified clauses --------
def generate_diff_html(before, after):
    """
    Returns a string where removed words are marked with ~~ and red,
    added words are green.
    """
    diff = difflib.ndiff(before.split(), after.split())
    result = []

    for token in diff:
        if token.startswith('- '):
            # Removed word: red and strike-through
            result.append(f"{token[2:]}")
        elif token.startswith('+ '):
            # Added word: green
            result.append(f"{token[2:]}")
        else:
            # Unchanged
            result.append(token[2:] if token.startswith('  ') else token)
    return " ".join(result)

# -------- Streamlit File Upload --------
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
            remarks = generate_diff_html(before_text, after_text)
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

    red = Font(color="FF0000")
    green = Font(color="008000")
    blue = Font(color="0000FF")

    for r_idx, row in enumerate(rows, start=2):
        for c_idx, value in enumerate(row, start=1):
            cell = ws.cell(row=r_idx, column=c_idx, value=value)

            # Apply coloring in Remarks column
            if c_idx == 5:
                status = row[3]
                if status == "Removed":
                    cell.font = red
                elif status == "New Clause Added":
                    cell.font = green
                elif status == "Statement Modified / Revised":
                    # For simplicity: make whole cell blue
                    cell.font = blue

    # Optional: adjust column width
    for i, col in enumerate(df.columns, start=1):
        ws.column_dimensions[get_column_letter(i)].width = 50

    buffer = BytesIO()
    wb.save(buffer)

    st.download_button(
        "Download Excel Change Log",
        buffer.getvalue(),
        "Clause_Change_Log.xlsx"
    )
