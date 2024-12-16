import tkinter as tk
from tkinter import filedialog, messagebox
import os
from workbookManager import WorkbookManager
from datetime import datetime
import findFixMethodsPCPAEC

RESULTS_DIRECTORY = "fileCrawlerResults"
LIST_ALL = "ListAll"
CHARACTER_LIMIT_FIND = "CharLimit-Find"
BAD_CHARACTER_FIND = "BadChar-Find"
SPACE_FIND = "SPC-Find"
FILE_EXTENSION_SUMMARY = "FileExt-Summary"
DUPLICATE_FILE_FIND = "DupFile-Find"

NULL_OPTION = "..."
SPACE_FIX = "SPC-Fix"
DELETE_OLD_FILES = "DelOldFiles-Fix"
DELETE_EMPTY_DIRECTORIES_FIX = "DelEmptyDirs-Fix"


STATELESS_FIND_METHODS = {
    LIST_ALL: findFixMethodsPCPAEC.listAll,
    CHARACTER_LIMIT_FIND: findFixMethodsPCPAEC.overCharLimitFind,
    BAD_CHARACTER_FIND: findFixMethodsPCPAEC.badCharErrorFind,
    SPACE_FIND: findFixMethodsPCPAEC.spaceErrorFind,
}
STATEFUL_FIND_METHODS = {
    FILE_EXTENSION_SUMMARY: findFixMethodsPCPAEC.fileExtensionMisc,
    DUPLICATE_FILE_FIND: findFixMethodsPCPAEC.duplicateFileMisc
}
POST_METHODS = {
    FILE_EXTENSION_SUMMARY: findFixMethodsPCPAEC.fileExtensionPost,
    DUPLICATE_FILE_FIND: findFixMethodsPCPAEC.duplicateFilePost
}

# First argument is Log. Second is Execute. Third is True if fileFixMethod, False if folderFixMethod. Fourth is fixArg's minimum (optional).
FIX_METHODS = {
    NULL_OPTION: None,
    SPACE_FIX: (findFixMethodsPCPAEC.spaceErrorFixLog, findFixMethodsPCPAEC.spaceErrorFixExecute, True),
    DELETE_OLD_FILES: (findFixMethodsPCPAEC.deleteOldFilesLog, findFixMethodsPCPAEC.deleteOldFilesExecute, True, 1),
    DELETE_EMPTY_DIRECTORIES_FIX: (findFixMethodsPCPAEC.deleteEmptyDirectoriesLog, findFixMethodsPCPAEC.deleteEmptyDirectoriesExecute, False, 0)
}


def validateArgument(arg:int, minimum:int):
    try:
        arg.strip()
        arg = int(arg)
        if (arg >= minimum): return arg
    except:
        pass


def control(dirAbsolute:str, includeSubfolders:bool, modify:bool, selectedFindMethods:list[str], selectedFixMethod:str, arg:str):
    if (not dirAbsolute): return -2
    
    # Create fileCrawlerResults directory name if does not exist
    try: os.mkdir(RESULTS_DIRECTORY)
    except: pass

    # >>> "folderName-20##-##-##.xlsx"
    workbookPathName = RESULTS_DIRECTORY + "\\" + dirAbsolute.split("/")[-1] + "-" \
    + datetime.today().strftime("%Y-%m-%d") + ".xlsx"

    # Initialize objects
    wbm = WorkbookManager(workbookPathName)
    findFixMethodsPCPAEC.setWorkBookManager(wbm)

    # Errors if this file already exists and is currently opened
    try:
        fileHandler = open(workbookPathName, 'w')
        fileHandler.close()
    except PermissionError:
        return -1

    # Set findMethods and fixMethod
    for fm in selectedFindMethods:
        if fm in STATELESS_FIND_METHODS:
            wbm.addStatelessFindMethod(fm, STATELESS_FIND_METHODS[fm])
        elif fm in STATEFUL_FIND_METHODS:
            wbm.addStatefulFindMethod(fm, STATEFUL_FIND_METHODS[fm])

        if fm in POST_METHODS:
           wbm.addPostStatefulFindMethod(fm, POST_METHODS[fm])

    if selectedFixMethod != NULL_OPTION:
        if len(FIX_METHODS[selectedFixMethod]) >= 4:
            arg = validateArgument(arg, FIX_METHODS[selectedFixMethod][2])
            if (arg is None): return -3
            else: wbm.setFixArg(arg)
        
        # If it's a fileFixMethod. Otherwise, it's a folderFixMethod
        if FIX_METHODS[selectedFixMethod][2]:
            wbm.setFileFixMethod(selectedFixMethod, FIX_METHODS[selectedFixMethod][int(modify)])
        else:
            wbm.setFolderFixMethod(selectedFixMethod, FIX_METHODS[selectedFixMethod][int(modify)])

    # After sheets have been created, set formatting
    wbm.setDefaultFormatting(dirAbsolute, includeSubfolders, modify)    

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
        if (renameState.get() and not messagebox.askyesnocancel("Are you sure?", "You have chosen to modify items. This is an IRREVERSIBLE action. Are you sure?")):
            print("Aborted. Continuing selection.")
            return

        exitCode = control(dirAbsoluteVar.get(), bool(includeSubFoldersState.get()), bool(renameState.get()), 
                           [findMethodListbox.get(fm) for fm in findMethodListbox.curselection()], fixOption.get(), parameterVar.get())

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
    renameState = tk.IntVar()
    fixOptionsList = list(FIX_METHODS)
    fixOption = tk.StringVar()

    includeSubfoldersCheckbutton = tk.Checkbutton(frame1, text="Subfolders?", variable=includeSubFoldersState, font=fontGeneral)
    modifyCheckbutton = tk.Checkbutton(frame1, text="Modify?", variable=renameState, font=fontGeneral)

    fixOption.set(fixOptionsList[0])
    fixDropdownMenu = tk.OptionMenu(frame1, fixOption, *fixOptionsList)
    fixDropdownMenu.config(font=fontGeneral) # set the font of fixDropdownMenu

    findMethodListbox = tk.Listbox(frame1, selectmode="multiple", exportselection=0, height=7)
    
    # dictionaries are ordered as of Python version 3.7
    for fm in STATELESS_FIND_METHODS.keys(): findMethodListbox.insert(tk.END, fm)
    for fm in STATEFUL_FIND_METHODS.keys(): findMethodListbox.insert(tk.END, fm)
    findMethodListbox.select_set(0)  # Hard code to select the first option.

    parameterEntry = tk.Entry(frame1, textvariable=parameterVar, width=7, font=fontGeneral)

    dirHeaderLabel = tk.Label(frame2, text = "Directory: ", font=fontGeneral)
    dirLabel = tk.Label(frame2, textvariable=dirAbsoluteVar, font=fontGeneral)

    browseButton = tk.Button(frame3, text="Browse", command=selectDirectory, font=fontGeneral)
    executeButton = tk.Button(frame3, text="Execute", command=launchController, font=fontGeneral)
    resultsButton = tk.Button(frame3, text="Results", command=openResultsFolder, font=fontGeneral)


    includeSubfoldersCheckbutton.pack(side=tk.LEFT)
    modifyCheckbutton.pack(side=tk.LEFT)
    findMethodListbox.pack(side=tk.LEFT)
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
