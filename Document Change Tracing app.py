import streamlit as st
import pandas as pd
import docx
import re
import difflib
from openpyxl import Workbook
from openpyxl.styles import Font
from io import BytesIO

st.title("Clause Change Log Generator – Reliable Version")

# -----------------------------
# Extract clauses from Word (working method)
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
    return clauses

# Word-level diff for modified clauses
def build_diff(before, after):
    before_words = before.split()
    after_words = after.split()
    diff = list(difflib.ndiff(before_words, after_words))
    return [(d[0], d[2:]) for d in diff if d[0] in ("+","-"," ")]

# -----------------------------
# File Upload
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
            rtype = "removed"
            rcontent = before_text
        elif clause not in before_clauses and clause in after_clauses:
            status = "New Clause Added"
            rtype = "added"
            rcontent = after_text
        elif before_text != after_text:
            status = "Statement Modified / Revised"
            rtype = "modified"
            rcontent = build_diff(before_text, after_text)
        else:
            continue

        results.append([document_name, before_text, after_text, status, rtype, rcontent])

    df = pd.DataFrame(results, columns=["Document Name","Before Clause","After Clause","Status","Remark Type","Remarks Content"])
    df.reset_index(inplace=True)
    df.rename(columns={'index':'SR #'}, inplace=True)
    df['SR #'] = df['SR #'] + 1

    # Remarks column for preview
    def format_remarks(row):
        if row['Remark Type'] in ["added","removed"]:
            return row['Remarks Content']
        elif row['Remark Type']=="modified":
            return " ".join([w for t,w in row['Remarks Content']])
        return ""

    df['Remarks'] = df.apply(format_remarks, axis=1)
    st.dataframe(df[['SR #','Document Name','Before Clause','After Clause','Status','Remarks']], use_container_width=True)

    # Excel export
    wb = Workbook()
    ws = wb.active
    ws.title = "Change Log"
    ws.append(['SR #','Document Name','Before Clause','After Clause','Status','Remarks'])
    red = Font(color="FF0000", strike=True)
    green = Font(color="008000")
    blue = Font(color="0000FF")

    for idx, row in df.iterrows():
        sr, docname, before_c, after_c, status, rtype, rcontent, remarks = row
        ws.append([sr, docname, before_c, after_c, status, ""])
        cell = ws.cell(row=ws.max_row, column=6)
        if rtype=="removed":
            cell.value = rcontent
            cell.font = red
        elif rtype=="added":
            cell.value = rcontent
            cell.font = green
        elif rtype=="modified":
            cell.value = " ".join([w for t,w in rcontent])
            cell.font = blue

    buffer = BytesIO()
    wb.save(buffer)
    st.download_button("Download Excel", data=buffer.getvalue(), file_name="Clause_Log.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
