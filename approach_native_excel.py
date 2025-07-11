import win32com.client
from pywintypes import com_error


WB_PATH = r"test.xlsx"
PATH_TO_PDF = r"test.pdf"

excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False

try:
    print("Start conversion to PDF")
    # Open
    wb = excel.Workbooks.Open(WB_PATH)
    # Specify the sheet you want to save by index. 1 is the first (leftmost) sheet.
    ws_index_list = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12]
    wb.WorkSheets(ws_index_list).Select()
    # Save
    wb.ActiveSheet.ExportAsFixedFormat(0, PATH_TO_PDF)
except com_error as e:
    print("failed.")
else:
    print("Succeeded.")
finally:
    wb.Close()
    excel.Quit()
