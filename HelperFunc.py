import os
import sys
import datetime
import openpyxl
import xlrd


def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, 'src\\', relative_path)


def readLocalFile(filename):
    f = open(filename, 'r')
    txt = f.read()
    f.close()

    return txt


def writeLocalFile(filename, txt):
    f = open(filename, 'w')
    f.write(txt)
    f.close()


def getFinalWellDate():
    day = datetime.datetime.now().strftime("%d")
    month = datetime.datetime.now().strftime("%b").upper()
    year = datetime.datetime.now().strftime("%Y")
    return f'{day}_{month}_{year}'


def getTimeNowText():
    time = datetime.datetime.now()
    return f'{time.hour}_{time.minute}_{time.second}'


def checkInputFile(excelFilename, file_extension):
    if (file_extension == '.xls'):
        workbook = xlrd.open_workbook(excelFilename)
        sh = workbook.sheet_by_index(workbook.nsheets-1)
        if (sh.nrows >= 1 and sh.ncols >= 7):
            return sh.cell_value(0, 6)
        else:
            return ''
    elif (file_extension == '.xlsx'):
        workbook = openpyxl.load_workbook(excelFilename, data_only=True)

        worksheet = workbook.active

        return worksheet._get_cell(1, 7).value
    else:
        return ''
