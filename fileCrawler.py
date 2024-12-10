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
CHARACTER_LIMIT_CHECK = "CharLimit-Check"
BAD_CHARACTER_CHECK = "BadChar-Check"
SPACE_CHECK = "SPC-Check"



def control(dirAbsolute: str, includeSubFolders: bool, renameFiles: bool, fixOption: str, arg: str):
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

    # Set checkMethod and fixMethod functions
    wbm.addCheckMethod(CHARACTER_LIMIT_CHECK, checkMethodsPCPAEC.overCharLimitCheck)        
    wbm.addCheckMethod(BAD_CHARACTER_CHECK, checkMethodsPCPAEC.badCharErrorCheck)
    wbm.addCheckMethod(SPACE_CHECK, checkMethodsPCPAEC.spaceErrorCheck)

    if (renameFiles):
        if fixOption == SPACE_FIX:
            wbm.setFixMethod(SPACE_FIX, checkMethodsPCPAEC.spaceErrorFixExecute)
        elif fixOption == DELETE_OLD_FILES:
            try:
                arg.strip()
                arg = int(arg)
                if (arg <= 0):
                    return -3
            except:
                return -3
            
            wbm.setFixMethod(DELETE_OLD_FILES, checkMethodsPCPAEC.deleteOldFilesExecute) 
            wbm.setFixArg(arg)
        else:
            wbm.setFixMethod(LIST_ALL, checkMethodsPCPAEC.listAll)
    else:
        if fixOption == SPACE_FIX:
            wbm.setFixMethod(SPACE_FIX, checkMethodsPCPAEC.spaceErrorFixLog)
        elif fixOption == DELETE_OLD_FILES:
            # Before execution, ensure the arg given is valid
            try:
                arg.strip()  # remove leading and trailing whitespace
                arg = int(arg)  # attempt to convert to int
                if (arg <= 0):
                    return -3
            except:
                return -3
            
            wbm.setFixMethod(DELETE_OLD_FILES, checkMethodsPCPAEC.deleteOldFilesLog)
            wbm.setFixArg(arg)
        else:
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
        if (renameState.get() and not messagebox.askyesnocancel("Are you sure?", "You have chosen to rename files. This is an IRREVERSIBLE action. Are you sure?")):
            print("Continuing selection.")
            return

        exitCode = control(dirAbsoluteVar.get(), bool(includeSubFoldersState.get()), bool(renameState.get()), fixOption.get(), parameterVar.get())

        if (exitCode == -1):
            print("Could not open file. Close file and try again.")
        elif (exitCode == -2):
            print("Invalid directory.")
        elif (exitCode == -3):
            print("Invalid argument. Must be an integer greater than 0.")



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

    fontGeneral = ("Calibri", 15)

    dirAbsoluteVar = tk.StringVar()
    parameterVar = tk.StringVar()

    includeSubFoldersState = tk.IntVar()
    isf = tk.Checkbutton(frame1, text="Include sub folders?", variable=includeSubFoldersState, font=fontGeneral)
    
    renameState = tk.IntVar()
    rf = tk.Checkbutton(frame1, text="Rename files?", variable=renameState, font=fontGeneral)

    fixOptions = [LIST_ALL, SPACE_FIX, DELETE_OLD_FILES]
    fixOption = tk.StringVar()
    fixOption.set(fixOptions[0])
    fixDropdownMenu = tk.OptionMenu(frame1, fixOption, *fixOptions)
    fixDropdownMenu.config(font=fontGeneral) # set the font of fixDropdownMenu

    parameterEntry = tk.Entry(frame1, textvariable=parameterVar, width=7, font=fontGeneral)

    dirHeaderLabel = tk.Label(frame2, text = "Directory: ", font=fontGeneral)
    dirLabel = tk.Label(frame2, textvariable=dirAbsoluteVar, font=fontGeneral)

    browseButton = tk.Button(frame3, text="Browse", command=selectDirectory, font=fontGeneral)
    resultsButton = tk.Button(frame3, text="Open results", command=openResultsFolder, font=fontGeneral)
    okayButton = tk.Button(frame3, text="Execute", command=launchController, font=fontGeneral)


    isf.pack(side=tk.LEFT)
    rf.pack(side=tk.LEFT)
    fixDropdownMenu.pack(side=tk.LEFT)
    parameterEntry.pack(side=tk.RIGHT)
    browseButton.pack(side=tk.LEFT)
    okayButton.pack(side=tk.LEFT)
    resultsButton.pack(side=tk.RIGHT)
    dirHeaderLabel.pack(side=tk.LEFT)
    dirLabel.pack(side=tk.RIGHT)

    root.mainloop()
        


def main():
    view()


if __name__ == "__main__":
    main()
