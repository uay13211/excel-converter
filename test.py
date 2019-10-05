import win32com.client
import os

# convert all the sheets in a excel file to individual pdf file
num = 1
o = win32com.client.Dispatch("Excel.Application")
o.Visible = False
o.DisplayAlerts = False
wb = o.Workbooks.Open(r"C:\Users\User\PycharmProjects\test\test.xls")

for sh in wb.Sheets:
    path = os.path.join(r"C:\Users\User\PycharmProjects\test", sh.Name + ".pdf")
    print(path)
    wb.Worksheets(sh.Name).Select()
    wb.ActiveSheet.ExportAsFixedFormat(0, path)


o.Quit()