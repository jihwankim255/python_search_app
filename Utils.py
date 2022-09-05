import openpyxl
from openpyxl.utils import get_column_letter

def excel_search(file):
    wb = openpyxl.load_workbook(file)
    worksheets_list = []
    sheet_row_list = []
    sheet_column_list = []
    for sheet in wb.worksheets:
        # print(sheet.title)
        worksheets_list.append(sheet.title)

        for i in range(sheet.max_row+1):
            sheet_row_list.append(i+1)
        for i in range(sheet.max_column+1):
            sheet_column_list.append(get_column_letter(i+1))
        print(sheet_row_list)
        print(sheet_column_list)
        print(sheet.title," 행의 수:",sheet.max_row)
        print(sheet.title," 열의 수:",sheet.max_column)

    return worksheets_list


def buttonStart_pressed2(file1, file2, sheetA, sheetB, row1A, row2A, col1A, col2A, row1B, row2B, col1B, col2B):
    print("file name: ",file1, file2)
    print("sheet name: ",sheetA, sheetB)
    print("rowA: ",row1A,row2A)
    print("colA: ",col1A,col2A)
    print("rowB: ",row1B,row2B)
    print("colB: ",col1B,col2B)


    wb1 = openpyxl.load_workbook(file1)
    wb2 = openpyxl.load_workbook(file2)

    sheet1 = wb1[sheetA]
    sheet2 = wb2[sheetB]


    for row in sheet1.rows:
        for row2 in sheet2.rows:
            # print(row2[4].value)
            # if row2 !="":
            if row[1].value is not None:
                # print(row[1].value)
                # print(row2[4].value)
                if row[1].value == row2[4].value and row[1].value != '수령인':

                    print(row[1], " = ", row2[4], row2[4].value)
                    print(row[1].row, row[1].column)

                    # row_col = get_column_letter(row[1].column) + str(14)
                    # # sheet2[row_col] = row[8]                    print(row_col)
                    sheet2.cell(row = row2[4].row, column=15).value = row[8].value


    result_file = file2.split('.xlsx')[0] + "(result).xlsx"
    wb2.save(result_file)


    # print(sheet2.max_row)
# excel_search()