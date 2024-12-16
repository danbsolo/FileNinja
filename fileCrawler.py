import tkinter as tk
from tkinter import filedialog, messagebox
import os
from workbookManager import WorkbookManager
from datetime import datetime
import checkMethodsPCPAEC

# constants
LIST_ALL = "ListAll"
SPACE_FIX = "SPC-Fix"
DELETE_OLD_FILES = "DelOldFiles-Fix"
DELETE_EMPTY_DIRECTORIES_FIX = "DelEmptyDirs-Fix"
CHARACTER_LIMIT_CHECK = "CharLimit-Check"
BAD_CHARACTER_CHECK = "BadChar-Check"
SPACE_CHECK = "SPC-Check"
FILE_EXTENSION_SUMMARY = "FileExt-Summary"
DUPLICATE_FILE_CHECK = "DupFile-Check"

# a constant dictionary of available checks and their corresponding function
CONCURRENT_CHECK_METHODS = {
    CHARACTER_LIMIT_CHECK: checkMethodsPCPAEC.overCharLimitCheck,
    BAD_CHARACTER_CHECK: checkMethodsPCPAEC.badCharErrorCheck,
    SPACE_CHECK: checkMethodsPCPAEC.spaceErrorCheck,
}

MISC_CHECK_METHODS = {
    FILE_EXTENSION_SUMMARY: checkMethodsPCPAEC.fileExtensionMisc,
    DUPLICATE_FILE_CHECK: checkMethodsPCPAEC.duplicateFileMisc
}

POST_CHECK_METHODS = {
    FILE_EXTENSION_SUMMARY: checkMethodsPCPAEC.fileExtensionPost,
    DUPLICATE_FILE_CHECK: checkMethodsPCPAEC.duplicateFilePost
}


def validateArgument(arg, minimum):
    # Before execution, ensure the arg given is valid
    try:
        arg.strip()  # remove leading and trailing whitespace
        arg = int(arg)  # attempt to convert to int
        if (arg >= minimum): return arg
    except:
        pass

    return  # None


def control(dirAbsolute: str, includeSubFolders: bool, renameFiles: bool, selectedCheckMethods: list[str], fixOption: str, arg: str):
    if (not dirAbsolute):
        return -2
    
    # Create fileCrawlerResults if does not exist. If it does, just pass
    try:
        os.mkdir("fileCrawlerResults")
    except:
        pass

    workbookPathName = "fileCrawlerResults\\" + dirAbsolute.split("/")[-1] + "-" \
    + datetime.today().strftime("%Y-%m-%d") + ".xlsx"

    # Initialize objects
    wbm = WorkbookManager(workbookPathName)
    checkMethodsPCPAEC.setWorkBookManager(wbm)

    # The selectedCheckMethods are guaranteed to be in CONCURRENT_CHECK_METHODS, but not necessarily POST_CHECK_METHODS
    for cm in selectedCheckMethods:
        if cm in CONCURRENT_CHECK_METHODS:
            wbm.addConcurrentCheckMethod(cm, CONCURRENT_CHECK_METHODS[cm])

        if cm in MISC_CHECK_METHODS:
            wbm.addMiscCheckMethod(cm, MISC_CHECK_METHODS[cm])

        if cm in POST_CHECK_METHODS:
           wbm.addPostCheckMethod(cm, POST_CHECK_METHODS[cm])

        


    if (renameFiles):
        if fixOption == SPACE_FIX:
            wbm.setFixMethod(SPACE_FIX, checkMethodsPCPAEC.spaceErrorFixExecute)
        elif fixOption == DELETE_OLD_FILES:
            arg = validateArgument(arg, 1)
            if (arg is None): return -3
            wbm.setFixMethod(DELETE_OLD_FILES, checkMethodsPCPAEC.deleteOldFilesExecute) 
            wbm.setFixArg(arg)
        elif fixOption == LIST_ALL:
            wbm.setFixMethod(LIST_ALL, checkMethodsPCPAEC.listAll)
        elif fixOption == DELETE_EMPTY_DIRECTORIES_FIX:
            arg = validateArgument(arg, 0)
            if (arg is None): return -4
            wbm.setFolderFixMethod(DELETE_EMPTY_DIRECTORIES_FIX, checkMethodsPCPAEC.deleteEmptyDirectoriesExecute)
            wbm.setFixArg(arg)
    else:
        if fixOption == SPACE_FIX:
            wbm.setFixMethod(SPACE_FIX, checkMethodsPCPAEC.spaceErrorFixLog)
        elif fixOption == DELETE_OLD_FILES:
            arg = validateArgument(arg, 1)
            if (not arg): return -3
            wbm.setFixMethod(DELETE_OLD_FILES, checkMethodsPCPAEC.deleteOldFilesLog)
            wbm.setFixArg(arg)
        elif fixOption == DELETE_EMPTY_DIRECTORIES_FIX:
            arg = validateArgument(arg, 0)
            if (arg is None): return -4
            wbm.setFolderFixMethod(DELETE_EMPTY_DIRECTORIES_FIX, checkMethodsPCPAEC.deleteEmptyDirectoriesLog)
            wbm.setFixArg(arg)
        elif fixOption == LIST_ALL:
            wbm.setFixMethod(LIST_ALL, checkMethodsPCPAEC.listAll)

    # Errors if this file already exists and is currently opened
    # Also solved by just having more unique file names, such as including the current time of execution    
    try:
        fileHandler = open(workbookPathName, 'w')
        fileHandler.close()
    except PermissionError:
        return -1
    
    print("\nCreating " + workbookPathName + "...")

    wbm.setDefaultFormatting(dirAbsolute, includeSubFolders, renameFiles)

    # Distinguish between the inclusion of exclusion of subfolders
    if (includeSubFolders):
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


    # Close gracefully and open newly created file for the user
    wbm.close()
    print("Opening " + workbookPathName + ".")
    os.startfile(workbookPathName)

    return 0


def view():
    def launchController():
        if (renameState.get() and not messagebox.askyesnocancel("Are you sure?", "You have chosen to modify items. This is an IRREVERSIBLE action. Are you sure?")):
            print("Continuing selection.")
            return

        exitCode = control(dirAbsoluteVar.get(), bool(includeSubFoldersState.get()), bool(renameState.get()), 
                           [checkMethodBox.get(cm) for cm in checkMethodBox.curselection()], fixOption.get(), parameterVar.get())

        if (exitCode == -1):
            print("Could not open file. Close file and try again.")
        elif (exitCode == -2):
            print("Invalid directory.")
        elif (exitCode == -3):
            print("Invalid argument. Must be an integer greater than 0.")
        elif (exitCode == -4):
            print("Invalid argument. Must be an integer greater than or equal to 0.")



    def selectDirectory():
        potentialDirectory = filedialog.askdirectory()

        # If user actually selected something, set dirAbsolute accordingly
        if (potentialDirectory):
            dirAbsoluteVar.set(potentialDirectory)


    def openResultsFolder():
        if os.path.exists("fileCrawlerResults"):
            os.startfile("fileCrawlerResults")
            return
        
        print("Folder does not exist. Try running a file crawl first.")


    # GUI Window
    root = tk.Tk()
    root.title("File Crawler")
    root.resizable(0, 0)
    root.attributes('-topmost', True)  # keeps root window at top layer

    frame1 = tk.Frame(root)
    frame1.pack()
    frame2 = tk.Frame(root)
    frame2.pack()
    frame3 = tk.Frame(root)
    frame3.pack(side=tk.BOTTOM)
    # frame4 = tk.Frame(frame3)
    # frame4.pack(side=tk.RIGHT)

    fontGeneral = ("Calibri", 15)

    dirAbsoluteVar = tk.StringVar()
    parameterVar = tk.StringVar()

    includeSubFoldersState = tk.IntVar()
    isf = tk.Checkbutton(frame1, text="Subfolders?", variable=includeSubFoldersState, font=fontGeneral)
    
    renameState = tk.IntVar()
    rf = tk.Checkbutton(frame1, text="Modify?", variable=renameState, font=fontGeneral)

    fixOptions = ["...", LIST_ALL, SPACE_FIX, DELETE_OLD_FILES, DELETE_EMPTY_DIRECTORIES_FIX]
    fixOption = tk.StringVar()
    fixOption.set(fixOptions[0])
    fixDropdownMenu = tk.OptionMenu(frame1, fixOption, *fixOptions)
    fixDropdownMenu.config(font=fontGeneral) # set the font of fixDropdownMenu

    checkMethodBox = tk.Listbox(frame1, selectmode="multiple", exportselection=0, height=6)
    
    # dictionaries are ordered as of Python version 3.7
    for ccm in CONCURRENT_CHECK_METHODS.keys(): checkMethodBox.insert(tk.END, ccm)
    for mcm in MISC_CHECK_METHODS.keys(): checkMethodBox.insert(tk.END, mcm)

    for i in range(3):  # hard code. Just select the first 3.
        checkMethodBox.select_set(i)

    parameterEntry = tk.Entry(frame1, textvariable=parameterVar, width=7, font=fontGeneral)

    dirHeaderLabel = tk.Label(frame2, text = "Directory: ", font=fontGeneral)
    dirLabel = tk.Label(frame2, textvariable=dirAbsoluteVar, font=fontGeneral)

    browseButton = tk.Button(frame3, text="Browse", command=selectDirectory, font=fontGeneral)
    okayButton = tk.Button(frame3, text="Execute", command=launchController, font=fontGeneral)
    resultsButton = tk.Button(frame3, text="Results", command=openResultsFolder, font=fontGeneral)



    isf.pack(side=tk.LEFT)
    rf.pack(side=tk.LEFT)
    checkMethodBox.pack(side=tk.LEFT)
    fixDropdownMenu.pack(side=tk.LEFT)
    parameterEntry.pack(side=tk.RIGHT)
    browseButton.pack(side=tk.LEFT)
    resultsButton.pack(side=tk.RIGHT)
    okayButton.pack(side=tk.RIGHT)
    dirHeaderLabel.pack(side=tk.LEFT)
    dirLabel.pack(side=tk.RIGHT)

    root.mainloop()
        


def main():
    view()


if __name__ == "__main__":
    main()
