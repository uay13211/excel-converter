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

        except:
            messagebox.showerror(title='Error', message='Please select a excel file')

        for sh in wb.Sheets:

            wb = o.Workbooks.Open(excelfile_path)
            filename_noextension = os.path.splitext(os.path.basename(xls_file))[0]
            save_path = os.path.join(window.savedir, filename_noextension)

            if not os.path.exists(save_path):
                os.mkdir(save_path)
            try:
                pdf_path = os.path.join(r"", save_path, sh.Name + ".pdf")
            except:
                messagebox.showerror(title='Error', message='Please provide an valid path')

            wb.Worksheets(sh.Name).Select()
            wb.ActiveSheet.ExportAsFixedFormat(0, pdf_path)

    o.Quit()


def openfile():
    window.filenames = filedialog.askopenfilenames(initialdir="C://", title="select excel files",
                                                   filetypes=(('xls files', '*.xls'), ('xlsx files', '*.xlsx')))
    print(window.filenames)

def savefile():
    window.savedir = filedialog.askdirectory()

window = tk.Tk()

window.title('Excel converter')
window.geometry('500x300')
menubar = tk.Menu(window)
filemenu = tk.Menu(menubar, tearoff=0)

menubar.add_cascade(label='File', menu=filemenu)

filemenu.add_command(label='Open', command=openfile)
filemenu.add_command(label='Save as', command=savefile)

window.config(menu=menubar)


button_convert = tk.Button(window, text='convert', command=convert)
button_convert.pack()

window.mainloop()