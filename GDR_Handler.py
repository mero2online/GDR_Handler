from tkinter import filedialog
from tkinter import *
import os
import shutil

from HandleExcel import HandleExcel
from HelperFunc import resource_path, writeLocalFile

filetypes = (
    ('EXCEL files', '*.xls'),
    ('EXCEL files', '*.xlsx'),
    ('All files', '*.*'),
)


def createDocxReport():
    input_file = resource_path(f'input\\input{input_file_extension.get()}')
    HandleExcel(input_file, input_file_extension.get())


def browseFile():
    # open-file dialog
    filename = filedialog.askopenfilename(
        title='Select a file...',
        filetypes=filetypes,)

    clearFiles()
    if (filename):
        selectedFilePath.set(filename)
        filenameOnly, file_extension = os.path.splitext(filename)
        input_file_extension.set('')
        if os.path.isfile(filename):
            shutil.copy(filename, resource_path(
                f'input\\input{file_extension}'))
            input_file_extension.set(file_extension)
    else:
        selectedFilePath.set('')
        input_file_extension.set('')


def saveFile():
    # save-as dialog
    filename = filedialog.askdirectory()

    if (filename):
        src_files = os.listdir(resource_path('output\\'))
        dest_dir = f'{filename}/Output'
        os.mkdir(dest_dir)
        for file_name in src_files:
            if os.path.isfile(resource_path(f'output\\{file_name}')):
                shutil.copy(resource_path(f'output\\{file_name}'), dest_dir)


def clearFiles():
    paths = ['input', 'output']
    for path in paths:
        if (os.path.exists(resource_path(path))):
            shutil.rmtree(resource_path(f'{path}\\'))
        os.mkdir(resource_path(path))
    writeLocalFile(resource_path('output\\output.txt'), '')
    shutil.copy(resource_path('GEOLOGICAL_DESCRIPTION_REPORT.docx'),
                resource_path('input'))


clearFiles()

root = Tk()

input_file_extension = StringVar()

browseBtn = Button(root, text="Browse File", background='#006778', foreground='#FFD124', borderwidth=2, relief="raised", padx=5, pady=5,
                   command=browseFile)
browseBtn.place(x=5, y=5, width=100, height=37)

createReportBtn = Button(root, text="Create Report", background='#0093AB', foreground='#FFD124', borderwidth=2, relief="groove", padx=5, pady=5,
                         disabledforeground='#00AFC1', command=createDocxReport)
createReportBtn.place(x=110, y=5, width=100, height=35)
# createReportBtn.config(state="disabled")


saveBtn = Button(root, text="Save File", background='#006778', foreground='#FFD124', borderwidth=2, relief="raised", padx=5, pady=5,
                 disabledforeground='#00AFC1', command=saveFile)
saveBtn.place(x=395, y=5, width=100, height=37)
# saveBtn.config(state='disabled')

selectedFilePath = StringVar()
currentFilePath = Label(
    root, textvariable=selectedFilePath, background='#006778', foreground='#FFD124', anchor=W)
currentFilePath.place(x=5, y=50, width=490, height=20)

root.title('GDR_Handler')
root.geometry('500x200')
root.configure(bg='#000')

root.resizable(False, False)
# Setting icon of master window
root.iconbitmap(resource_path('gdr.ico'))
# Start program
root.mainloop()
