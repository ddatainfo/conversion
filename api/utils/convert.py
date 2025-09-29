import xlrd
from openpyxl import Workbook

def convert_xls_to_xlsx(xls_file, xlsx_file):
    """
    Converts an .xls file to .xlsx format.

    :param xls_file: Path to the .xls file.
    :param xlsx_file: Path to save the converted .xlsx file.
    """
    workbook = xlrd.open_workbook(xls_file)
    sheet = workbook.sheet_by_index(0)

    new_workbook = Workbook()
    new_sheet = new_workbook.active

    for row_idx in range(sheet.nrows):
        for col_idx in range(sheet.ncols):
            new_sheet.cell(row=row_idx + 1, column=col_idx + 1, value=sheet.cell_value(row_idx, col_idx))

    new_workbook.save(xlsx_file)
    print(f"Converted {xls_file} to {xlsx_file}")

if __name__ == "__main__":
    xls_file = "/mnt/c/Users/admin/Desktop/conversion/conversion/REPORT/302/302.xls"
    xlsx_file = "converted_302.xlsx"
    convert_xls_to_xlsx(xls_file, xlsx_file)