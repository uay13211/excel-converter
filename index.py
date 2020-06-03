import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import win32com.client
import os


class ExcelConverter:
    def __init__(self):
        # excel
        self.excelApp = win32com.client.Dispatch("Excel.Application")
        self.excelApp.Visible = False
        self.excelApp.DisplayAlerts = False

        # GUI
        self.window = tk.Tk();
        self.window.title('Excel converter')
        self.window.geometry('300x300')

        self.openFileBtn = tk.Button(self.window, text='Open File', command=self.openfile)
        self.saveAsBtn = tk.Button(self.window, text='Save in', command=self.savefile)
        self.execBtn = tk.Button(self.window, text='Convert', command=self.convert)

        self.openFileBtn.place(height=50, width=200, relx=0.17, rely=0.15)
        self.saveAsBtn.place(height=50, width=200, relx=0.17, rely=0.4)
        self.execBtn.place(height=50, width=200, relx=0.17, rely=0.65)

    def run(self):
        self.window.mainloop()

    # convert all the sheets in a excel file to individual pdf file
    def convert(self):
        if len(self.window.fileNames) == 0:
            messagebox.showerror(title='Error', message='Please select a excel file')
            return

        for excelFile in self.window.fileNames:
            excelFilePath = os.path.join(r'', excelFile)
            wb = self.excelApp.Workbooks.Open(excelFilePath)
            folderName = os.path.splitext(os.path.basename(excelFile))[0].replace(' ', '_')
            folderPath = os.path.join(self.window.saveDir, folderName)

            if not os.path.exists(folderPath):
                os.mkdir(folderPath)

            for sh in wb.Sheets:
                # only convert the visible sheets
                if sh.Visible != 0:
                    sheetNameRepalceWithScore = sh.Name.replace(' ', '_')
                    try:
                        pdfPath = os.path.join(r"", folderPath, sheetNameRepalceWithScore + ".pdf")
                    except:
                        messagebox.showerror(title='Error', message='Please provide an valid path')

                    if not (os.path.exists(pdfPath) or os.path.exists(os.path.join(folderPath, sh.Name + ".pdf"))):
                        wb.Worksheets(sh.Name).Activate()
                        wb.ActiveSheet.ExportAsFixedFormat(0, pdfPath)
                        os.rename(pdfPath, os.path.join(folderPath, sh.Name + ".pdf"))

        self.excelApp.Quit()

    def openfile(self):
        self.window.fileNames = filedialog.askopenfilenames(initialdir="C://", title="select excel files",
                                                       filetypes=(('xls files', '*.xls'), ('xlsx files', '*.xlsx')))

    def savefile(self):
        self.window.saveDir = filedialog.askdirectory()


app = ExcelConverter()
app.run()







