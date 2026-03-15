import streamlit as st
import pandas as pd
import docx
import re
import difflib
from openpyxl import Workbook
from openpyxl.styles import Font
from io import BytesIO

st.set_page_config(page_title="Clause Change Log Generator", layout="wide")

st.title("📄 Document Clause Change Log Generator")


# -----------------------------
# Extract Clauses
# -----------------------------
def extract_clauses(file):

    doc = docx.Document(file)
    clauses = {}

    pattern = r'(\d+\.\d+(\.\d+)*)'

    for para in doc.paragraphs:

        text = para.text.strip()

        if not text:
            continue

        match = re.search(pattern, text)

        if match:
            clause_no = match.group(1)
            clauses[clause_no] = text

    # detect clauses in tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:

                text = cell.text.strip()

                if not text:
                    continue

                match = re.search(pattern, text)

                if match:
                    clause_no = match.group(1)
                    clauses[clause_no] = text

    return clauses


# -----------------------------
# Word Level Difference
# -----------------------------
def build_diff(before, after):

    before_clean = re.sub(r'^\d+(\.\d+)*\s*', '', before)
    after_clean = re.sub(r'^\d+(\.\d+)*\s*', '', after)

    before_words = before_clean.split()
    after_words = after_clean.split()

    diff = list(difflib.ndiff(before_words, after_words))

    result = []

    for d in diff:

        if d.startswith("+ "):
            result.append(("added", d[2:]))

        elif d.startswith("- "):
            result.append(("removed", d[2:]))

        elif d.startswith("  "):
            result.append(("same", d[2:]))

    return result


# -----------------------------
# Upload Files
# -----------------------------
before_file = st.file_uploader("Upload BEFORE Document", type=["docx"])
after_file = st.file_uploader("Upload AFTER Document", type=["docx"])


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
            remarks_type = "removed"
            remarks_content = before_text

        elif clause not in before_clauses and clause in after_clauses:

            status = "New Clause Added"
            remarks_type = "added"
            remarks_content = after_text

        elif before_text != after_text:

            status = "Statement Modified / Revised"
            remarks_type = "modified"
            remarks_content = build_diff(before_text, after_text)

        else:
            continue

        results.append([
            document_name,
            before_text,
            after_text,
            status,
            remarks_type,
            remarks_content
        ])

    df = pd.DataFrame(results, columns=[
        "Document Name",
        "Before Clause",
        "After Clause",
        "Status",
        "Remark Type",
        "Remarks Content"
    ])

    st.subheader("Preview Change Log")

    preview = df.drop(columns=["Remark Type", "Remarks Content"])
    st.dataframe(preview, use_container_width=True)


# -----------------------------
# Create Excel
# -----------------------------
    wb = Workbook()
    ws = wb.active
    ws.title = "Change Log"

    headers = [
        "Document Name",
        "Before Clause",
        "After Clause",
        "Status",
        "Remarks"
    ]

    ws.append(headers)

    red = Font(color="FF0000", strike=True)
    green = Font(color="008000")
    blue = Font(color="0000FF")

    for row in results:

        docname, before_c, after_c, status, rtype, rcontent = row

        ws.append([docname, before_c, after_c, status, ""])

        remarks_cell = ws.cell(row=ws.max_row, column=5)

        if rtype == "removed":

            remarks_cell.value = re.sub(r'^\d+(\.\d+)*\s*', '', rcontent)
            remarks_cell.font = red

        elif rtype == "added":

            remarks_cell.value = re.sub(r'^\d+(\.\d+)*\s*', '', rcontent)
            remarks_cell.font = green

        elif rtype == "modified":

            remarks_text = ""

            for t, word in rcontent:

                if t == "same":
                    remarks_text += word + " "

                elif t == "added":
                    remarks_text += word + " "

                elif t == "removed":
                    remarks_text += word + " "

            remarks_cell.value = remarks_text.strip()

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
