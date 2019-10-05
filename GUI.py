import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import win32com.client
import os


# convert all the sheets in a excel file to individual pdf file
def convert():
    o = win32com.client.Dispatch("Excel.Application")
    o.Visible = False
    o.DisplayAlerts = False
    for xls_file in window.filenames:

        try:
            excelfile_path = os.path.join(r'', xls_file)
            wb = o.Workbooks.Open(excelfile_path)

        except:
            messagebox.showerror(title='Error', message='Please select a excel file')

        filename_noextension = os.path.splitext(os.path.basename(xls_file))[0].replace(' ', '_')
        save_path = os.path.join(window.savedir, filename_noextension)
        if not os.path.exists(save_path):
            os.mkdir(save_path)

        for sh in wb.Sheets:
            sheet_name_repalce = sh.Name.replace(' ', '_')
            try:
                pdf_path = os.path.join(r"", save_path, sheet_name_repalce + ".pdf")
            except:
                messagebox.showerror(title='Error', message='Please provide an valid path')

            if not (os.path.exists(pdf_path) or os.path.exists(os.path.join(save_path, sh.Name + ".pdf"))):
                wb.Worksheets(sh.Name).Select()
                wb.ActiveSheet.ExportAsFixedFormat(0, pdf_path)

                os.rename(pdf_path, os.path.join(save_path, sh.Name + ".pdf"))

    o.Quit()


def openfile():
    window.filenames = filedialog.askopenfilenames(initialdir="C://", title="select excel files",
                                                   filetypes=(('xls files', '*.xls'), ('xlsx files', '*.xlsx')))
    print(window.filenames)


def savefile():
    window.savedir = filedialog.askdirectory()


window = tk.Tk()

window.title('Excel converter')
window.geometry('300x300')

open_button = tk.Button(window, text='Open File', command=openfile)
save_as_button = tk.Button(window, text='Save in', command=savefile)
convert_button = tk.Button(window, text='Convert', command=convert)

open_button.place(height=50, width=200, relx=0.17, rely=0.15)
save_as_button.place(height=50, width=200, relx=0.17, rely=0.4)
convert_button.place(height=50, width=200, relx=0.17, rely=0.65)

window.mainloop()