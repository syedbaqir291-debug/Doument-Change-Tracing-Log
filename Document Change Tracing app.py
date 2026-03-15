import streamlit as st
import pandas as pd
import docx
import re
import difflib
from openpyxl import Workbook
from openpyxl.styles import Font
from io import BytesIO

st.set_page_config(page_title="Document Clause Comparison Tool", layout="wide")

st.title("📄 Clause Change Log Generator")


# -----------------------------
# Function: Extract Clauses
# -----------------------------
def extract_clauses(file):

    document = docx.Document(file)
    clauses = {}

    pattern = r'(\d+\.\d+(\.\d+)*)'

    # Read normal paragraphs
    for para in document.paragraphs:

        text = para.text.strip()

        if text == "":
            continue

        match = re.search(pattern, text)

        if match:
            clause_no = match.group(1)
            clauses[clause_no] = text

    # Read tables
    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:

                text = cell.text.strip()

                if text == "":
                    continue

                match = re.search(pattern, text)

                if match:
                    clause_no = match.group(1)
                    clauses[clause_no] = text

    return clauses


# -----------------------------
# Function: Detect Differences
# -----------------------------
def find_changes(before, after):

    before_words = before.split()
    after_words = after.split()

    diff = difflib.ndiff(before_words, after_words)

    added = []
    removed = []

    for word in diff:

        if word.startswith("+ "):
            added.append(word[2:])

        elif word.startswith("- "):
            removed.append(word[2:])

    result = ""

    if added:
        result += "Added: " + " ".join(added)

    if removed:
        result += " | Removed: " + " ".join(removed)

    return result


# -----------------------------
# File Upload
# -----------------------------
before_file = st.file_uploader("Upload BEFORE Document", type=["docx"])
after_file = st.file_uploader("Upload AFTER Document", type=["docx"])


# -----------------------------
# Comparison Logic
# -----------------------------
if before_file and after_file:

    before_clauses = extract_clauses(before_file)
    after_clauses = extract_clauses(after_file)

    document_name = after_file.name

    all_clauses = set(before_clauses.keys()).union(set(after_clauses.keys()))

    results = []

    for clause in sorted(all_clauses):

        before_text = before_clauses.get(clause, "")
        after_text = after_clauses.get(clause, "")

        if clause in before_clauses and clause not in after_clauses:

            status = "Removed"
            remarks = before_text

        elif clause not in before_clauses and clause in after_clauses:

            status = "New Clause Added"
            remarks = after_text

        elif before_text != after_text:

            status = "Statement Modified / Revised"
            remarks = find_changes(before_text, after_text)

        else:
            continue

        results.append([
            document_name,
            before_text,
            after_text,
            status,
            remarks
        ])


# -----------------------------
# Create DataFrame
# -----------------------------
    df = pd.DataFrame(
        results,
        columns=[
            "Document Name",
            "Before Clause",
            "After Clause",
            "Status",
            "Remarks"
        ]
    )

    st.subheader("📊 Change Log Preview")
    st.dataframe(df, use_container_width=True)


# -----------------------------
# Create Excel File
# -----------------------------
    wb = Workbook()
    ws = wb.active
    ws.title = "Change Log"

    ws.append(df.columns.tolist())

    red = Font(color="FF0000")
    green = Font(color="008000")
    blue = Font(color="0000FF")

    for row in results:

        ws.append(row)

        status = row[3]

        remarks_cell = ws.cell(row=ws.max_row, column=5)

        if status == "Removed":
            remarks_cell.font = red

        elif status == "New Clause Added":
            remarks_cell.font = green

        elif status == "Statement Modified / Revised":
            remarks_cell.font = blue


# -----------------------------
# Download Excel
# -----------------------------
    buffer = BytesIO()
    wb.save(buffer)

    st.download_button(
        label="📥 Download Excel Change Log",
        data=buffer.getvalue(),
        file_name="Clause_Change_Log.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
