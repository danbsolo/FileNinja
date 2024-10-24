import xlsxwriter
# from os import path, walk, getcwd, open
import os
from tkinter import filedialog
import time

def populateWorkbook(dirAbsolute, workbook, worksheet):
    # set color and width of col 0
    dirColFormat = workbook.add_format({'bold': True})
    dirColFormat.set_bg_color("#99CCFF")
    worksheet.set_column(0, 0, 120)

    # set width of col 1
    worksheet.set_column(1, 1, 70)

    row = 0

    # every item in the tuple is a tuple of: each directory, what directories reside in it, what files reside in it 
    for (dirPath, dirs, files) in os.walk(dirAbsolute):
        worksheet.write(row, 0, dirPath, dirColFormat)

        for file in files:
            worksheet.write(row, 1, file)
            # worksheet.write(row, 2, time.ctime(os.path.getmtime(dirPath + "\\" + file))) # crashes for long file names, or something
            row += 1

def main():
    # currentDir = os.path.dirname(os.path.abspath(__file__))
    dirAbsolute = filedialog.askdirectory()
    dirName = dirAbsolute.split("/")[-1]
    fileName = dirName + "Directory" + ".xlsx"

    # the excel workbook and worksheet after the directory chosen
    workbook = xlsxwriter.Workbook(fileName)
    worksheet = workbook.add_worksheet(dirName + " Sheet")

    populateWorkbook(dirAbsolute, workbook, worksheet)

    print(fileName, "has been created.")

    workbook.close()


main()