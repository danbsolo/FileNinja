import tkinter as tk
from tkinter import filedialog, messagebox
import os
from workbookManager import WorkbookManager
from datetime import datetime
from defs import *


def control(dirAbsolute:str, includeSubfolders:bool, modify:bool, selectedFindProcedures:list[str], selectedFixProcedure:str, unprocessedArg:str):
    if (not dirAbsolute): return -2
    
    # Create fileCrawlerResults directory name if does not exist
    try: os.mkdir(RESULTS_DIRECTORY)
    except: pass

    # >>> "folderName-20##-##-##.xlsx"
    workbookPathName = RESULTS_DIRECTORY + "\\" + dirAbsolute.split("/")[-1] + "-" \
    + datetime.today().strftime("%Y-%m-%d") + ".xlsx"

    # Initialize objects
    wbm = WorkbookManager(workbookPathName)
    setWorkbookManager(wbm)

    # Errors if this file already exists and is currently opened
    try:
        fileHandler = open(workbookPathName, 'w')
        fileHandler.close()
    except PermissionError:
        return -1

    # Set findProcedures and fixProcedure
    for fm in selectedFindProcedures:
        wbm.addFindProcedure(FIND_PROCEDURES[fm])

    if selectedFixProcedure != NULL_OPTION:
        fixProcedureObject = FIX_PROCEDURES[selectedFixProcedure]

        if not wbm.setFixArg(fixProcedureObject, unprocessedArg):
            return -3
        
        wbm.setFixProcedure(fixProcedureObject, modify)
                    
    wbm.styleSummarySheet(dirAbsolute, includeSubfolders, modify)    

    print("\nCreating " + workbookPathName + "...")
    if (includeSubfolders):
        wbm.folderCrawl(os.walk(dirAbsolute))
    else:
        # mimic os.walk()'s output but only for the current directory
        dirFolders = []
        dirFiles = []
        
        for item in os.listdir(dirAbsolute):
            if os.path.isfile(os.path.join(dirAbsolute, item)):
                dirFiles.append(item)
            else:
                dirFolders.append(item)

        wbm.folderCrawl([(dirAbsolute, dirFolders, dirFiles)])

    wbm.close()
    print("Opening " + workbookPathName + ".")
    os.startfile(workbookPathName)
    return 0



def view():
    def launchController():
        if (modifyState.get() and not messagebox.askyesnocancel("Are you sure?", "You have chosen to modify items. This is an IRREVERSIBLE action. Are you sure?")):
            print("Aborted. Continuing selection.")
            return

        exitCode = control(dirAbsoluteVar.get(), bool(includeSubFoldersState.get()), bool(modifyState.get()), 
                           [findProcedureListbox.get(fm) for fm in findProcedureListbox.curselection()], fixOption.get(), parameterVar.get())

        if (exitCode == 0):
            pass
        elif (exitCode == -1):
            print("Could not open file. Close file and try again.")
        elif (exitCode == -2):
            print("Invalid directory.")
        elif (exitCode == -3):
            print("Invalid argument.")
        else:
            print("An error has occurred.")

    def selectDirectory():
        potentialDirectory = filedialog.askdirectory()

        # If user actually selected something, set dirAbsolute accordingly
        if (potentialDirectory):
            dirAbsoluteVar.set(potentialDirectory)

    def openResultsFolder():
        if os.path.exists(RESULTS_DIRECTORY):
            os.startfile(RESULTS_DIRECTORY)
        else:
            print("Folder does not exist. Try executing a file crawl first.")


    root = tk.Tk()
    root.title("File Crawler")
    root.resizable(0, 0)
    root.attributes('-topmost', True)  # keeps root window at top layer

    frame1 = tk.Frame(root)
    frame1.pack()
    frame2 = tk.Frame(root)
    frame2.pack()
    frame3 = tk.Frame(root)
    frame3.pack()

    fontGeneral = ("Calibri", 15)
    dirAbsoluteVar = tk.StringVar()
    parameterVar = tk.StringVar()
    includeSubFoldersState = tk.IntVar()
    modifyState = tk.IntVar()
    fixOptionsList = list(FIX_PROCEDURES)
    fixOption = tk.StringVar()

    includeSubfoldersCheckbutton = tk.Checkbutton(frame1, text="Subfolders?", variable=includeSubFoldersState, font=fontGeneral)
    modifyCheckbutton = tk.Checkbutton(frame1, text="Modify?", variable=modifyState, font=fontGeneral)

    fixOption.set(fixOptionsList[0])
    fixDropdownMenu = tk.OptionMenu(frame1, fixOption, *fixOptionsList)
    fixDropdownMenu.config(font=fontGeneral) # set the font of fixDropdownMenu

    findProcedureListbox = tk.Listbox(frame1, selectmode="multiple", exportselection=0, height=8)
    
    # dictionaries are ordered as of Python version 3.7
    for fm in FIND_PROCEDURES.keys(): findProcedureListbox.insert(tk.END, fm)
    findProcedureListbox.select_set(0)  # Hard code to select the first option.

    parameterEntry = tk.Entry(frame1, textvariable=parameterVar, width=7, font=fontGeneral)

    dirHeaderLabel = tk.Label(frame2, text = "Directory: ", font=fontGeneral)
    dirLabel = tk.Label(frame2, textvariable=dirAbsoluteVar, font=fontGeneral)

    browseButton = tk.Button(frame3, text="Browse", command=selectDirectory, font=fontGeneral)
    executeButton = tk.Button(frame3, text="Execute", command=launchController, font=fontGeneral)
    resultsButton = tk.Button(frame3, text="Results", command=openResultsFolder, font=fontGeneral)


    includeSubfoldersCheckbutton.pack(side=tk.LEFT)
    modifyCheckbutton.pack(side=tk.LEFT)
    findProcedureListbox.pack(side=tk.LEFT)
    fixDropdownMenu.pack(side=tk.LEFT)
    parameterEntry.pack(side=tk.LEFT)

    dirHeaderLabel.pack(side=tk.LEFT)
    dirLabel.pack(side=tk.LEFT)

    browseButton.pack(side=tk.LEFT)
    executeButton.pack(side=tk.LEFT)
    resultsButton.pack(side=tk.LEFT)

    root.mainloop()
        

def main():
    view()


if __name__ == "__main__":
    main()
