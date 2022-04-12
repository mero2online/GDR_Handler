from tkinter import filedialog
from tkinter import *
from tkinter import messagebox
from tkinter import scrolledtext
import pyperclip
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


def selectAll(*event):
    inputPathBox.tag_add(SEL, "1.0", END)
    inputPathBox.mark_set(INSERT, "1.0")
    inputPathBox.see(INSERT)

    outputPathBox.tag_add(SEL, "1.0", END)
    outputPathBox.mark_set(INSERT, "1.0")
    outputPathBox.see(INSERT)

    return "break"


def copy(*event):
    try:
        text = inputPathBox.selection_get()
        text = outputPathBox.selection_get()
        pyperclip.copy(text)
    except TclError:
        pass
    return "break"


def addToInputPathBox(txt):
    inputPathBox.config(state="normal")
    inputPathBox.delete('1.0', END)
    inputPathBox.insert(INSERT, txt)
    inputPathBox.config(state="disabled")


def addToOutputPathBox(txt):
    outputPathBox.config(state="normal")
    outputPathBox.delete('1.0', END)
    outputPathBox.insert(INSERT, txt)
    outputPathBox.config(state="disabled")


def createDocxReport():
    asyncio.run(createReport())


async def createReport():
    input_file = resource_path(f'input\\input{input_file_extension.get()}')
    await HandleExcel(input_file, input_file_extension.get())
    if os.path.isfile(resource_path('output\\GEOLOGICAL_DESCRIPTION_REPORT.docx')):
        saveBtn.config(state='normal')
        messagebox.showinfo(
            'Success', f'GEOLOGICAL DESCRIPTION REPORT\nCreated successfully')


def browseFile():
    # open-file dialog
    filename = filedialog.askopenfilename(
        title='Select a file...',
        filetypes=filetypes,)

    clearFiles()
    if (filename):
        pathOnly, file_extension = os.path.splitext(filename)
        selectedFilePath.set(pathOnly.split('/')[-1])
        addToInputPathBox(filename)
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
            addToInputPathBox('')
            addToOutputPathBox('')
            setButtonsDisabled()

    else:
        selectedFilePath.set('')
        addToInputPathBox('')
        addToOutputPathBox('')
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
        addToOutputPathBox(dest_dir)
        await copyFiles(src_files, dest_dir)
        result = messagebox.askquestion(
            'Success', f'Files saved successfully\n\nOpen output folder?')
        if result == 'yes':
            opd = os.getcwd()  # Get original directory
            os.chdir(dest_dir)  # Change directory to run command
            os.system('start.')  # Run command
            os.chdir(opd)  # Return to original directory


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

inputPathBoxLabel = LabelFrame(root, text="Input file path",
                               background='#006778', foreground='#FFD124', padx=5, pady=5)
inputPathBoxLabel.place(x=5, y=75, width=490, height=20)

inputPathBox = scrolledtext.ScrolledText(
    root, background='#006778', foreground='#FFD124')
inputPathBox.place(x=5, y=95, width=490, height=40)

inputPathBox.config(state="disabled")

outputPathBoxLabel = LabelFrame(root, text="Output file path",
                                background='#006778', foreground='#FFD124', padx=5, pady=5)
outputPathBoxLabel.place(x=5, y=140, width=490, height=20)

outputPathBox = scrolledtext.ScrolledText(
    root, background='#006778', foreground='#FFD124')
outputPathBox.place(x=5, y=160, width=490, height=40)

outputPathBox.config(state="disabled")

madeWithLoveBy = Label(
    root, text='Made with ❤ by Mohamed Omar', background='#006778', foreground='#FFD124', font=('monospace', 9, 'bold'))
madeWithLoveBy.place(x=310, y=230, width=190, height=20)

m = Menu(root, tearoff=0)
m.add_command(label="Select All", command=selectAll)
m.add_command(label="Copy", command=copy)


def do_popup(event):
    try:
        m.tk_popup(event.x_root, event.y_root)
    finally:
        m.grab_release()


inputPathBox.bind("<Button-3>", do_popup)
outputPathBox.bind("<Button-3>", do_popup)

setButtonsDisabled()

root.title('GDR_Handler')
root.geometry('500x250')
root.configure(bg='#000')

root.resizable(False, False)
# Setting icon of master window
root.iconbitmap(resource_path('gdr.ico'))
# Start program
root.mainloop()
