import tkinter as tk
from tkinter import filedialog, messagebox
import os
from workbookManager import WorkbookManager
from datetime import datetime
import checkMethodsPCPAEC


def control(dirAbsolute: str, includeSubFolders: bool, renameFiles: bool, fixOption: str):
    # Create fileCrawlerResults if does not exist. If it does, just pass
    try:
        os.mkdir("fileCrawlerResults")
    except:
        pass

    workbookPathName = "fileCrawlerResults\\" + dirAbsolute.split("/")[-1] + "-" \
    + datetime.today().strftime("%Y-%m-%d") + ".xlsx"
    print("\nCreating " + workbookPathName + "...")

    # Errors if this file already exists and is currently opened
    # Also solved by just having more unique file names, such as including the current time of execution    
    try:
        fileHandler = open(workbookPathName, 'w')
        fileHandler.close()
    except PermissionError:
        return -1

    # Initialize objects
    wbm = WorkbookManager(workbookPathName)
    checkMethodsPCPAEC.setWorkBookManager(wbm)

    # Set checkMethod and fixMethod functions
    if (renameFiles):
        wbm.addCheckMethod("CharLimit-Check", checkMethodsPCPAEC.overCharLimitCheck)        
        wbm.addCheckMethod("BadChar-Check", checkMethodsPCPAEC.badCharErrorCheck)
        wbm.addCheckMethod("SPC-Check", checkMethodsPCPAEC.spaceErrorCheck)
        
        if fixOption == "SPC-Fix":
            wbm.setFixMethod("SPC-Fix", checkMethodsPCPAEC.spaceErrorFixExecute)
        else:
            wbm.setFixMethod("List-All", checkMethodsPCPAEC.listAllLog)
    else:
        wbm.addCheckMethod("CharLimit-Check", checkMethodsPCPAEC.overCharLimitCheck)        
        wbm.addCheckMethod("BadChar-Check", checkMethodsPCPAEC.badCharErrorCheck)
        wbm.addCheckMethod("SPC-Check", checkMethodsPCPAEC.spaceErrorCheck)

        if fixOption == "SPC-Fix":
            wbm.setFixMethod("SPC-Fix", checkMethodsPCPAEC.spaceErrorFixLog)
        else:
            wbm.setFixMethod("List-All", checkMethodsPCPAEC.listAllLog)

    
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
        if (not dirAbsolute.get()):
            print("Invalid directory.")
            return
        
        if (renameState.get() and not messagebox.askyesnocancel("Are you sure?", "You have chosen to rename files. This is an IRREVERSIBLE action. Are you sure?")):
            print("Continuing selection.")
            return

        # root.destroy()
        exitCode = control(dirAbsolute.get(), bool(includeSubFoldersState.get()), bool(renameState.get()), fixOption.get())

        if (exitCode == -1):
            print("Could not open file. Close file and try again.")
    

    def selectDirectory():
        potentialDirectory = filedialog.askdirectory()

        # If user actually selected something, set dirAbsolute accordingly
        if (potentialDirectory):
            dirAbsolute.set(potentialDirectory)


    # GUI Window
    root = tk.Tk()
    root.title("File Crawler")
    root.resizable(0, 0)

    frame1 = tk.Frame(root)
    frame1.pack()
    frame2 = tk.Frame(root)
    frame2.pack()
    frame3 = tk.Frame(root)
    frame3.pack(side=tk.BOTTOM)

    dirAbsolute = tk.StringVar()

    includeSubFoldersState = tk.IntVar()
    isf = tk.Checkbutton(frame1, text="Include sub folders?", variable=includeSubFoldersState)
    
    renameState = tk.IntVar()
    rf = tk.Checkbutton(frame1, text="Rename files?", variable=renameState)

    fixOptions = ["List-All", "SPC-Fix"]
    fixOption = tk.StringVar()
    fixOption.set("List-All")
    fixDropdownMenu = tk.OptionMenu(frame2, fixOption, *fixOptions)

    browseButton = tk.Button(frame2, text="BROWSE", command=selectDirectory)
    
    okayButton = tk.Button(frame2, text="OK", command=launchController)
    dirHeaderLabel = tk.Label(frame3, text = "Directory: ")
    dirLabel = tk.Label(frame3, textvariable=dirAbsolute)

    isf.pack(side=tk.LEFT)
    rf.pack(side=tk.RIGHT)
    fixDropdownMenu.pack(side=tk.LEFT)
    browseButton.pack(side=tk.LEFT)
    okayButton.pack(side=tk.RIGHT)
    dirHeaderLabel.pack(side=tk.LEFT)
    dirLabel.pack(side=tk.RIGHT)
    
    root.mainloop()
        


def main():
    view()


if __name__ == "__main__":
    main()
