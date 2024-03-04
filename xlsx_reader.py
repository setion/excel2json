'''Module that reads an xlsx spreadsheet and can produce json data from it'''
import pylightxl as xl

def get_column_names(sheet):
    '''Takes a single worksheet, returns the strings in the top row of each column'''
    column_lists = sheet.cols
    column_names = []

    for column_list in column_lists:
        column_names.append(column_list[00])

    return column_names

def get_row_data(row, column_names, column_type):
    '''takes a single row of a worksheet and an array of rows,
        returns an object with column_name:rowvalue
    '''
    row_data = {}
    counter = 0
    d_column_type = {"int": int, "float": float, "list": eval, "number": int,
                     "array": eval, "string": str, "boolean": eval, "bool": eval, 'object': eval}
    for cell in row:
        column_name = column_names[counter]
        if column_type[counter] == "array" or column_type[counter] == "list" and cell == "":
            cell = '[]'
        if column_type[counter] == "boolean" or column_type[counter] == "bool":
            if cell is False or cell is True:
                cell = str(cell)
            else:
                if cell.lower() == 'true':
                    cell = 'true'
                elif cell.lower() == 'false':
                    cell = 'False'
                else:
                    raise 'The cell value must be True or False, but actually is ' + str(cell)

        row_data[column_name] = d_column_type.get(column_type[counter])(cell)
        counter = counter + 1
    return row_data

def get_sheet_data(sheet, column_names):
    '''Takes a single worksheet, returns an object with row data'''
    max_rows = sheet.size[0]
    sheet_data = {}
    column_type = sheet.row(2)
    for idx in range(4, max_rows):
        row = sheet.row(idx)
        row_data = get_row_data(row, column_names, column_type)
        sheet_data[row_data.get(column_names[0])] = row_data
    return sheet_data

def get_workbook_data(workbook):
    '''Takes a workbook and returns all worksheet data'''
    workbook_sheet_names = workbook.ws_names
    print(workbook_sheet_names)
    sn_no = 1
    output_file_suffix = ""
    if len(workbook_sheet_names) > 1:
        print("当前Excel有多个表格，请选择需要导出的表格")
        for sn in workbook_sheet_names:
            print(f'{sn_no}.{sn}\t', end='')
            sn_no += 1
        print(f'{sn_no}.全部')
    sheet_num = input("请输入要导出的表格序号：")
    print()
    workbook_data = {}
    if int(sheet_num) == sn_no:
        export_sheet = workbook_sheet_names
    else:
        export_sheet = [workbook_sheet_names[int(sheet_num) - 1]]
        output_file_suffix = '_'+export_sheet[0]
    for sheet_name in export_sheet:
        worksheet = workbook.ws(ws=sheet_name)
        column_names = get_column_names(worksheet)
        sheet_data = get_sheet_data(worksheet, column_names)
        if int(sheet_num) == sn_no:
            workbook_data[sheet_name.lower().replace(' ', '_')] = sheet_data
        else:
            workbook_data = sheet_data
    return [output_file_suffix, workbook_data]

def get_workbook(filename):
    '''opens a workbook for reading'''
    return xl.readxl(filename)
