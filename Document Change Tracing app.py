import streamlit as st
import pandas as pd
import docx
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from io import BytesIO

st.title("Amendment History Extractor (Multi-Document)")

# -------- File Upload --------
uploaded_files = st.file_uploader(
    "Upload Word Documents", type=["docx"], accept_multiple_files=True
)

if uploaded_files:
    all_rows = []
    excel_headers = None  # Will set header from first document

    for file in uploaded_files:
        doc = docx.Document(file)
        document_name = file.name

        # Find the table with 'Amendment History' in first row or any cell
        amendment_table = None
        for table in doc.tables:
            first_row_text = " ".join([cell.text.strip() for cell in table.rows[0].cells])
            if "Amendment History" in first_row_text or "Amendment" in first_row_text:
                amendment_table = table
                break

        if not amendment_table:
            st.warning(f"No Amendment History table found in {document_name}")
            continue

        # Extract header once from first document only
        if excel_headers is None:
            excel_headers = ["Document Name"] + [cell.text.strip() for cell in amendment_table.rows[0].cells]

        # Extract data rows only (skip header row)
        for row in amendment_table.rows[1:]:
            row_data = [cell.text.strip() for cell in row.cells]
            all_rows.append([document_name] + row_data)

    if all_rows:
        # Create DataFrame
        df = pd.DataFrame(all_rows, columns=excel_headers)
        st.dataframe(df, use_container_width=True)

        # -------- Excel Export --------
        wb = Workbook()
        ws = wb.active
        ws.append(df.columns.tolist())

        # Append all rows
        for r_idx, row in enumerate(all_rows, start=2):
            for c_idx, value in enumerate(row, start=1):
                ws.cell(row=r_idx, column=c_idx, value=value)

        # Merge document name column per document
        current_doc = None
        start_row = 2
        for r_idx, row in enumerate(all_rows, start=2):
            doc_name = row[0]
            if doc_name != current_doc:
                # Merge previous doc rows
                if current_doc is not None and r_idx - start_row > 1:
                    ws.merge_cells(start_row=start_row, start_column=1, end_row=r_idx-1, end_column=1)
                current_doc = doc_name
                start_row = r_idx
        # Merge last document
        if current_doc is not None and r_idx - start_row >= 1:
            ws.merge_cells(start_row=start_row, start_column=1, end_row=r_idx, end_column=1)

        # Adjust column widths
        for i, col in enumerate(df.columns, start=1):
            ws.column_dimensions[get_column_letter(i)].width = 30

        # Save Excel to buffer
        buffer = BytesIO()
        wb.save(buffer)

        st.download_button(
            "Download Combined Amendment History Excel",
            buffer.getvalue(),
            "Amendment_History.xlsx"
        )
    else:
        st.info("No Amendment History data found in uploaded documents.")
