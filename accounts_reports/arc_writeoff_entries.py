def write_arc_entry_format(df, file_path, sheet_title):
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_title[:31]

    blue_fill = PatternFill("solid", fgColor="BFEAF5")
    bold_font = Font(bold=True)
    border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin")
    )

    center = Alignment(horizontal="center", vertical="center")

    ws.merge_cells("B1:D1")
    ws.merge_cells("E1:G1")

    headers_row1 = {
        "A1": "B.Code",
        "B1": "Principal Amount",
        "E1": "Interest Amount",
        "H1": "OTS Amount"
    }

    for cell, value in headers_row1.items():
        ws[cell] = value
        ws[cell].font = bold_font
        ws[cell].fill = blue_fill
        ws[cell].alignment = center
        ws[cell].border = border

    second_headers = ["JLG", "CPP", "IL", "JLG", "CPP", "IL"]

    for col_idx, header in enumerate(second_headers, start=2):
        cell = ws.cell(row=2, column=col_idx)
        cell.value = header
        cell.font = bold_font
        cell.fill = blue_fill
        cell.alignment = center
        cell.border = border

    ws["A2"].fill = blue_fill
    ws["H2"].fill = blue_fill

    start_row = 3

    for idx, row in df.iterrows():
        r = start_row + idx

        values = [
            row["B.Code"],
            row["Principal_JLG"],
            row["Principal_CPP"],
            row["Principal_IL"],
            row["Interest_JLG"],
            row["Interest_CPP"],
            row["Interest_IL"],
            row["OTS Amount"]
        ]

        for c, val in enumerate(values, start=1):
            cell = ws.cell(row=r, column=c)
            cell.value = val
            cell.alignment = center
            cell.border = border

    total_row = start_row + len(df)

    ws.cell(row=total_row, column=1).value = "Grand Total"

    for col in range(2, 9):
        ws.cell(row=total_row, column=col).value = f"=SUM({get_column_letter(col)}{start_row}:{get_column_letter(col)}{total_row-1})"

    for col in range(1, 9):
        cell = ws.cell(row=total_row, column=col)
        cell.font = bold_font
        cell.fill = blue_fill
        cell.alignment = center
        cell.border = border

    for row in ws.iter_rows(min_row=1, max_row=total_row, min_col=1, max_col=8):
        for cell in row:
            cell.border = border
            cell.alignment = center

    ws.column_dimensions["A"].width = 14

    for col in range(2, 9):
        ws.column_dimensions[get_column_letter(col)].width = 18

    ws.freeze_panes = "A3"

    wb.save(file_path)
