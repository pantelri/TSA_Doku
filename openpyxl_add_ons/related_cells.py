def cell_below(sheet, cell):
    row_below = cell.row + 1
    cell_below = sheet.cell(row=row_below, column=cell.column)
    return cell_below.value

