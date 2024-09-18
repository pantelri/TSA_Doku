from openpyxl.utils import get_column_letter

def write_basic_data(sheet, data):
    for i, row in enumerate(data.itertuples(), start=28):
        sheet[f'B{i}'] = row.Date
        sheet[f'C{i}'] = row.Month
        sheet[f'D{i}'] = row.Fiscal_Year
        sheet[f'E{i}'] = row.Period

def write_total_data(sheet, data, account):
    sheet['G27'] = f"{account} total"
    total_column = next((col for col in data.columns if '_total' in col), None)
    if total_column:
        for i, value in enumerate(data[total_column], start=28):
            sheet[f'G{i}'] = value

def write_subtotal_data(sheet, data, account, account_name):
    subtotal_columns = [col for col in data.columns if f"{account_name}_subtotal" in col]
    if len(subtotal_columns) > 3:
        raise ValueError("Doku-Template wird bisher nur für max. 3 Subtotal-Spalten des Accounts_to_audit unterstützt")
    for idx, col in enumerate(subtotal_columns):
        cell = chr(ord('H') + idx)  # H, I, J
        suffix = col.split(f"{account_name}_subtotal_")[-1]
        sheet[f'{cell}27'] = f"{account} {suffix}"
        for i, value in enumerate(data[col], start=28):
            sheet[f'{cell}{i}'] = value

def write_volume_data(sheet, data):
    volume_columns = [col for col in data.columns if 'Volume' in col]
    if len(volume_columns) > 3:
        sheet.insert_cols(15, len(volume_columns) - 3)
    for idx, col in enumerate(volume_columns):
        cell = get_column_letter(13 + idx)  # M, N, O, ...
        header = col.replace('_', ' ')
        sheet[f'{cell}27'] = header
        for i, value in enumerate(data[col], start=28):
            sheet[f'{cell}{i}'] = value

def write_index_data(sheet, data):
    index_columns = [col for col in data.columns if 'index_' in col]
    if len(index_columns) > 2:
        sheet.insert_cols(18, len(index_columns) - 2)
        for col_idx in range(18, 18 + len(index_columns) - 2):
            for row in sheet.iter_rows(min_row=1, max_row=sheet.max_row, min_col=col_idx, max_col=col_idx):
                for cell in row:
                    cell.font = sheet['Q' + str(cell.row)].font.copy()
                    cell.border = sheet['Q' + str(cell.row)].border.copy()
                    cell.fill = sheet['Q' + str(cell.row)].fill.copy()
                    cell.number_format = sheet['Q' + str(cell.row)].number_format
                    cell.alignment = sheet['Q' + str(cell.row)].alignment.copy()
        sheet.column_dimensions['S'].width = sheet.column_dimensions['Q'].width
    
    last_index_column = None
    for idx, col in enumerate(index_columns):
        cell = get_column_letter(17 + idx)  # Q, R, ...
        last_index_column = cell
        header = f"price {col.replace('_', ' ')}"
        sheet[f'{cell}27'] = header
        for i, value in enumerate(data[col], start=28):
            sheet[f'{cell}{i}'] = value
    
    # Lösche alle Spalten rechts von der letzten index_data Spalte
    if last_index_column:
        last_column_index = sheet.max_column
        columns_to_delete = last_column_index - ord(last_index_column) + ord('A')
        if columns_to_delete > 0:
            sheet.delete_cols(ord(last_index_column) - ord('A') + 2, columns_to_delete)
