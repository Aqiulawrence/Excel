import openpyxl

def main(start, end, file):
    wb = openpyxl.load_workbook(file)
    ws = wb.active

    data = []
    for row in ws[start:end]:
        row_data = []
        for cell in row:
            row_data.append(str(cell.value).split('\n')[0])
        if len(row_data) <= 1: # 单列
            data.append(row_data[0])
        else: # 多列
            data.append(row_data)

    return data
