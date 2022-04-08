from tkinter import filedialog
from tkinter import *
from tkinter import messagebox
import asyncio
import os
import shutil

from HandleExcel import HandleExcel
from HelperFunc import checkInputFile, getFinalWellDate, getTimeNowText, resource_path, writeLocalFile

try:
    import pyi_splash  # type: ignore
    pyi_splash.close()
except:
    pass

filetypes = (
    ('EXCEL files', '*.xls'),
    ('EXCEL files', '*.xlsx'),
    ('All files', '*.*'),
)


def createDocxReport():
    asyncio.run(createReport())


async def createReport():
    input_file = resource_path(f'input\\input{input_file_extension.get()}')
    await HandleExcel(input_file, input_file_extension.get())
    if os.path.isfile(resource_path('output\\GEOLOGICAL_DESCRIPTION_REPORT.docx')):
        saveBtn.config(state='normal')


def browseFile():
    # open-file dialog
    filename = filedialog.askopenfilename(
        title='Select a file...',
        filetypes=filetypes,)

    clearFiles()
    if (filename):
        selectedFilePath.set(filename)
        filenameOnly, file_extension = os.path.splitext(filename)
        res = checkInputFile(filename, file_extension)
        if (res == 'LITHOLOGY DESCRIPTION SHEET'):
            createReportBtn.config(state="normal")
            input_file_extension.set('')
            if os.path.isfile(filename):
                shutil.copy(filename, resource_path(
                    f'input\\input{file_extension}'))
                input_file_extension.set(file_extension)
        else:
            messagebox.showerror(
                'File error', 'Please load valid LITHOLOGY DESCRIPTION SHEET')
            selectedFilePath.set('')
            setButtonsDisabled()

    else:
        selectedFilePath.set('')
        input_file_extension.set('')
        setButtonsDisabled()


def saveFile():
    # save-as dialog
    filename = filedialog.askdirectory()
    asyncio.run(saveAllFiles(filename))


async def saveAllFiles(filename):
    if (filename):
        src_files = os.listdir(resource_path('output\\'))
        date = getFinalWellDate()
        time = getTimeNowText()
        dest_dir = f'{filename}/GDR-Output-{date}-{time}'
        await copyFiles(src_files, dest_dir)
        messagebox.showinfo(
            'Success', f'Files saved successfully to\n{dest_dir}')


async def copyFiles(src_files, dest_dir):
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
    writeLocalFile(resource_path(
        'output\\GEOLOGICAL_DESCRIPTION_REPORT.txt'), '')
    shutil.copy(resource_path('GEOLOGICAL_DESCRIPTION_REPORT.docx'),
                resource_path('input'))


def setButtonsDisabled():
    createReportBtn.config(state="disabled")
    saveBtn.config(state="disabled")


clearFiles()

root = Tk()

input_file_extension = StringVar()

browseBtn = Button(root, text="Browse File", background='#006778', foreground='#FFD124', borderwidth=2, relief="raised", padx=5, pady=5,
                   command=browseFile)
browseBtn.place(x=5, y=5, width=100, height=37)

createReportBtn = Button(root, text="Create Report", background='#0093AB', foreground='#FFD124', borderwidth=2, relief="groove", padx=5, pady=5,
                         disabledforeground='#00AFC1', command=createDocxReport)
createReportBtn.place(x=110, y=5, width=100, height=35)


saveBtn = Button(root, text="Save File", background='#006778', foreground='#FFD124', borderwidth=2, relief="raised", padx=5, pady=5,
                 disabledforeground='#00AFC1', command=saveFile)
saveBtn.place(x=395, y=5, width=100, height=37)

selectedFilePath = StringVar()
currentFilePath = Label(
    root, textvariable=selectedFilePath, background='#006778', foreground='#FFD124', anchor=W)
currentFilePath.place(x=5, y=50, width=490, height=20)

madeWithLoveBy = Label(
    root, text='Made with ❤ by Mohamed Omar', background='#006778', foreground='#FFD124', font=('monospace', 9, 'bold'))
madeWithLoveBy.place(x=310, y=180, width=190, height=20)

setButtonsDisabled()

root.title('GDR_Handler')
root.geometry('500x200')
root.configure(bg='#000')

root.resizable(False, False)
# Setting icon of master window
root.iconbitmap(resource_path('gdr.ico'))
# Start program
root.mainloop()
