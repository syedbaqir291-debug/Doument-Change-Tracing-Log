import streamlit as st
import pandas as pd
import docx
import re
import difflib
from openpyxl import Workbook
from openpyxl.styles import Font
from io import BytesIO

st.title("Clause Change Log Generator")


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


def find_added_words(before, after):

    before_words = before.split()
    after_words = after.split()

    diff = difflib.ndiff(before_words, after_words)

    added = [word[2:] for word in diff if word.startswith('+ ')]

    return " ".join(added)


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

        remarks = ""

        if clause in before and clause not in after:

            status = "Removed"
            remarks = before_text

        elif clause not in before and clause in after:

            status = "New Clause Added"
            remarks = after_text

        elif before_text != after_text:

            status = "Statement Modified / Revised"
            remarks = find_added_words(before_text, after_text)

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


# Excel export

    wb = Workbook()
    ws = wb.active

    ws.append(df.columns.tolist())

    red = Font(color="FF0000")
    green = Font(color="008000")
    blue = Font(color="0000FF")

    for row in rows:

        ws.append(row)

        status = row[3]
        remarks_cell = ws.cell(row=ws.max_row, column=5)

        if status == "Removed":
            remarks_cell.font = red

        elif status == "New Clause Added":
            remarks_cell.font = green

        elif status == "Statement Modified / Revised":
            remarks_cell.font = blue


    buffer = BytesIO()
    wb.save(buffer)

    st.download_button(
        "Download Excel Change Log",
        buffer.getvalue(),
        "Clause_Change_Log.xlsx"
    )
