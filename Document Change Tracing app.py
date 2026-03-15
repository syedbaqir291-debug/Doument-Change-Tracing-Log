import streamlit as st
import pandas as pd
import docx
import re
from openpyxl import Workbook
from openpyxl.styles import Font
from io import BytesIO

st.set_page_config(page_title="Clause Change Log Generator", layout="wide")

st.title("📄 Document Clause Change Log Generator")

# -------- Clause Extraction -------- #

def extract_clauses(file):

    doc = docx.Document(file)
    clauses = {}

    pattern = r'^\d+(\.\d+)*'

    for para in doc.paragraphs:
        text = para.text.strip()

        match = re.match(pattern, text)

        if match:
            clause_no = match.group()
            clauses[clause_no] = text

    return clauses


# -------- Upload Files -------- #

before_file = st.file_uploader("Upload BEFORE Document", type=["docx"])
after_file = st.file_uploader("Upload AFTER Document", type=["docx"])


if before_file and after_file:

    before_clauses = extract_clauses(before_file)
    after_clauses = extract_clauses(after_file)

    all_clauses = set(before_clauses.keys()).union(set(after_clauses.keys()))

    rows = []

    doc_name = after_file.name

    for clause in sorted(all_clauses):

        before_text = before_clauses.get(clause, "")
        after_text = after_clauses.get(clause, "")

        if clause in before_clauses and clause not in after_clauses:
            status = "Removed"

        elif clause not in before_clauses and clause in after_clauses:
            status = "New Clause Added"

        elif before_text != after_text:
            status = "Statement Modified / Revised"

        else:
            continue

        rows.append([
            doc_name,
            before_text,
            after_text,
            status
        ])

    df = pd.DataFrame(rows, columns=[
        "Document Name",
        "Before Clause",
        "After Clause",
        "Status"
    ])

    st.subheader("📊 Change Log Preview")
    st.dataframe(df, use_container_width=True)


# -------- Excel Creation -------- #

    wb = Workbook()
    ws = wb.active
    ws.title = "Change Log"

    headers = ["Document Name", "Before Clause", "After Clause", "Status"]
    ws.append(headers)

    red_font = Font(color="FF0000")
    green_font = Font(color="008000")
    blue_font = Font(color="0000FF")

    for row in rows:
        ws.append(row)

        status_cell = ws.cell(row=ws.max_row, column=4)

        if row[3] == "Removed":
            status_cell.font = red_font

        elif row[3] == "New Clause Added":
            status_cell.font = green_font

        elif row[3] == "Statement Modified / Revised":
            status_cell.font = blue_font


# -------- Save Excel -------- #

    buffer = BytesIO()
    wb.save(buffer)

    st.download_button(
        label="📥 Download Excel Change Log",
        data=buffer.getvalue(),
        file_name="Clause_Change_Log.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
