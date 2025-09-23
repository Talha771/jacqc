import pandas as pd
import streamlit as st
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment, PatternFill

st.set_page_config(page_title="Expense Formatter", layout="wide")

st.title("💳 Expense Formatter")
st.write("Upload your Excel/CSV, and download the formatted grouped report.")

# File upload
uploaded_file = st.file_uploader("Upload file", type=["xlsx", "csv"])

if uploaded_file:
    # Read input
    if uploaded_file.name.endswith(".csv"):
        data = pd.read_csv(uploaded_file)
    else:
        data = pd.read_excel(uploaded_file)

    st.success("File uploaded successfully!")

    # Process
    data.columns = data.iloc[5]
    data = data.iloc[6:]
    data = data.sort_values(by=["Account #", "Card Member", "Category", "Amount"]).reset_index(drop=True)

    cols = ["Date", "Appears On Your Statement As", "Amount", "Category", "Description", "Card Member", "Account #"]
    df = data[cols].copy()
    df = df.rename(columns={
        "Appears On Your Statement As": "Location",
        "Description": "Descript"
    })

    # Empty JOBS
    df["JOBS"] = " - "

    # Split into groups
    dfs, order = [], []
    for (member, acct), group in df.groupby(["Card Member", "Account #"]):
        group = group.reset_index(drop=True)
        group = group[["JOBS", "Date", "Location", "Amount", "Category", "Descript"]]
        group.columns = ["JOBS", "Date", "Location", "Amount", "Category", "Description"]
        dfs.append(group)
        order.append((member, acct))

    final = pd.concat(dfs, axis=1)


    # Save to BytesIO
    output = BytesIO()
    final.to_excel(output, index=False)
    output.seek(0)

   # Post-process with openpyxl
    wb = load_workbook(output)
    ws = wb.active

    colors = ["FFCCCC", "CCFFCC", "CCCCFF", "FFFFCC", "FFCCFF", "CCFFFF", "E0E0E0"]

    col_offset = 1
    ws.insert_rows(1)
    for idx, (member, acct) in enumerate(order):
        width = 6
        start_col, end_col = col_offset, col_offset + width - 1

        # Merge top row and set group header
        ws.merge_cells(start_row=1, start_column=start_col, end_row=1, end_column=end_col)
        ws.cell(row=1, column=start_col).value = f"{member} {acct}"
        ws.cell(row=1, column=start_col).alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        # Fill header rows with color
        fill = PatternFill(start_color=colors[idx % len(colors)], end_color=colors[idx % len(colors)], fill_type="solid")
        for col in range(start_col, end_col + 1):
            for row in [1]:  # top 2 rows for headers
                cell = ws.cell(row=row, column=col)
                cell.fill = fill
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        # Format Amount column and add conditional highlighting
        amount_col = start_col + 3
        for row in range(3, ws.max_row + 1):
            cell = ws.cell(row=row, column=amount_col)
            if isinstance(cell.value, (int, float)):
                cell.number_format = u'"$"#,##0.00'
                # Highlight based on value
                if cell.value > 300:
                    cell.fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
                elif cell.value > 75:
                    cell.fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
            cell.alignment = Alignment(wrap_text=True)

        # Wrap text for other columns
        for col in range(start_col, end_col + 1):
            if col != amount_col:
                for row in range(3, ws.max_row + 1):
                    ws.cell(row=row, column=col).alignment = Alignment(wrap_text=True, vertical="top")

        col_offset += width


    # Save to buffer again
    final_buf = BytesIO()
    wb.save(final_buf)
    final_buf.seek(0)

    # Download button
    st.download_button(
        label="⬇️ Download Formatted Excel",
        data=final_buf,
        file_name="expenses_by_cardmember.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
