import tkinter as tk
from tkinter import filedialog
import os
from workbookManager import WorkbookManager
from datetime import datetime
from defs import *
from sys import argv


def control(dirAbsolute:str, includeSubfolders:bool, modify:bool, selectedFindProcedures:list[str], selectedFixProcedure:str, unprocessedArg:str, excludedDirs:set[str]):
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

    # print("\nCreating " + workbookPathName + "...")
    if (includeSubfolders):
        if (not excludedDirs):
            wbm.folderCrawl(os.walk(dirAbsolute), [])

        else:
            # Run checks on excludedDirs
            for exDir in excludedDirs:
                # Check dirAbsolute is not an exDir. Check that exDir is a child to dirAbsolute
                if (exDir == dirAbsolute) or not (exDir.startswith(dirAbsolute)):
                    return -4

                # Check that no exDir is a subfolder to another exDir
                for exDir2 in excludedDirs:
                    if (exDir != exDir2) and (exDir.startswith(exDir2)):
                        return -4
                

            walkResults = []

            for subDirAbsolute, subDirFolders, subDirFiles in os.walk(dirAbsolute):
                isDirIncluded = True
                subDirAbsolute = subDirAbsolute.replace("\\", "/")
                
                for exDir in excludedDirs:
                    if subDirAbsolute.startswith(exDir):  # exclude by subDirAbsolute to be precise
                        subDirFolders[:] = []  # stop traversal of this folder and its subfolders
                        isDirIncluded = False
                        break

                if isDirIncluded:
                    walkResults.append((subDirAbsolute, subDirFolders, subDirFiles))
            
            wbm.folderCrawl(walkResults, excludedDirs)
    
    else:
        # mimic os.walk()'s output but only for the current directory
        dirFolders = []
        dirFiles = []
        
        for item in os.listdir(dirAbsolute):
            if os.path.isfile(os.path.join(dirAbsolute, item)):
                dirFiles.append(item)
            else:
                dirFolders.append(item)

        wbm.folderCrawl([(dirAbsolute, dirFolders, dirFiles)], [])

    wbm.close()
    os.startfile(workbookPathName)
    return 0



def view(isAdmin: bool):
    def launchController():
        if modifyState.get() and not tk.messagebox.askyesnocancel("Allow Modify?", "You have chosen to modify items. This is an IRREVERSIBLE action. Are you sure?"):
            # print("Aborted. Continuing selection.")
            return

        root.title(FILE_CRAWLER + ": CURRENTLY RUNNING...")
        exitCode = control(dirAbsoluteVar.get(), bool(includeSubFoldersState.get()), bool(modifyState.get()), 
                           [findListbox.get(fm) for fm in findListbox.curselection()],
                           fixListbox.get(tk.ACTIVE) if isAdmin else NULL_OPTION,
                           parameterVar.get(),
                           excludedDirs)
        root.title(FILE_CRAWLER)
        
        errorMessage = ""
        if (exitCode == 0):
            return
        elif (exitCode == -1):
            errorMessage = "Could not open file. Close file and try again."
        elif (exitCode == -2):
            errorMessage = "Invalid directory."
        elif (exitCode == -3):
            errorMessage = "Invalid argument."
        elif (exitCode == -4):
            errorMessage = "Invalid excluded directories."
        else:
            errorMessage = "An error has occurred."

        tk.messagebox.showerror("Error: " + str(exitCode), errorMessage)


    def selectDirectory():
        potentialDirectory = filedialog.askdirectory(title="Browse to SELECT",
                                                     initialdir=dirAbsoluteVar.get(),
                                                     mustexist=True)

        if (potentialDirectory):
            dirAbsoluteVar.set(potentialDirectory)


    def excludeDirectory():
        potentialExcludedDirectory = filedialog.askdirectory(title="Browse to EXCLUDE",
                                                             initialdir=dirAbsoluteVar.get(),
                                                             mustexist=True)

        if (not potentialExcludedDirectory) or (potentialExcludedDirectory in excludedDirs):
            return

        excludedDirs.add(potentialExcludedDirectory)
        excludeListbox.insert(tk.END, potentialExcludedDirectory)
        
        newHeight = len(excludedDirs) + 1
        excludeListbox.config(height = newHeight)
        root.geometry("{}x{}".format(rootWidth, rootHeight + (newHeight * listboxHeightMultiplier)))


    def removeExcludedDirectory():
        toBeRemovedIndex = excludeListbox.curselection()

        if not toBeRemovedIndex:
            return
        
        excludedDirs.remove(excludeListbox.get(tk.ACTIVE)) # Only supports single removal for now
        excludeListbox.delete(toBeRemovedIndex)

        newHeight = len(excludedDirs) + 1
        excludeListbox.config(height = newHeight)
        root.geometry("{}x{}".format(rootWidth, rootHeight + (newHeight * listboxHeightMultiplier)))
      

        
    def openResultsDirectory():
        if os.path.exists(RESULTS_DIRECTORY):
            os.startfile(RESULTS_DIRECTORY)
        else:
            tk.messagebox.showinfo("Directory DNE", "Directory \"" + RESULTS_DIRECTORY + "\" does not exist. Try executing a file crawl first.")

    
    excludedDirs = set()

    #
    listboxHeight = max(len(FIND_PROCEDURES.keys()), len(FIX_PROCEDURES.keys())) +1
    listboxHeightMultiplier = 17

    # root window stuff
    root = tk.Tk()
    root.title(FILE_CRAWLER)
    root.resizable(0, 0)
    rootWidth = 500 if isAdmin else 300
    rootHeight = (listboxHeight * listboxHeightMultiplier) + (365 if isAdmin else 310)
    root.geometry("{}x{}".format(rootWidth, rootHeight))

    if isAdmin: root.attributes('-topmost', True)  # keeps root window at top layer

    root.bind('<Control-Key-w>', lambda e: root.destroy())
    root.bind('<Control-Key-W>', lambda e: root.destroy())
        
    frames = []
    for i in range(9):
        frames.append(tk.Frame(root, bd=0, relief=tk.SOLID))
        frames[i].pack(fill="x", padx=10, pady=3)


    # aesthetic/layout variables
    fontType = "None"
    fontSize = 15
    fontGeneral = (fontType, fontSize)
    fontSmall = (fontType, int(fontSize/3*2))
    # listboxHeight defined above
    listboxWidth = int(rootWidth/15) if isAdmin else int(rootWidth/10)
    finalButtonsWidth = 20 if isAdmin else 10 # HARD CODED

    # data variables
    dirAbsoluteVar = tk.StringVar()
    parameterVar = tk.StringVar()
    includeSubFoldersState = tk.IntVar()
    modifyState = tk.IntVar()


    # widgets
    browseButton = tk.Button(frames[0], text="Browse to Select", command=selectDirectory, font=fontGeneral, width=rootWidth)
    browseButton.pack()

    dirHeaderLabel = tk.Label(frames[1], text = "Directory:", font=fontGeneral)
    dirLabel = tk.Label(frames[1], textvariable=dirAbsoluteVar, font=fontSmall, anchor="e") 
    dirHeaderLabel.pack(side=tk.LEFT)
    dirLabel.pack(side=tk.LEFT)

    excludeButton = tk.Button(frames[2], text="Browse to Exclude", command=excludeDirectory, font=fontGeneral, width=rootWidth)
    excludeScrollbar = tk.Scrollbar(frames[2], orient=tk.HORIZONTAL)
    excludeListbox = tk.Listbox(frames[2], exportselection=0, width=rootWidth, height=0, xscrollcommand=excludeScrollbar.set)
    excludeScrollbar.config(command=excludeListbox.xview)
    excludeListbox.bind("<Double-Button-1>", lambda _: removeExcludedDirectory()) # double left click
    excludeListbox.bind("<Button-3>", lambda _: removeExcludedDirectory()) # right click
    excludeButton.pack()
    excludeListbox.pack()
    excludeScrollbar.pack()
    
    findLabel = tk.Label(frames[3], text="Find", font=fontGeneral)
    if isAdmin:
        fixLabel = tk.Label(frames[3], text="Fix", font=fontGeneral)
        fixLabel.pack(side=tk.RIGHT, padx=(0, rootWidth/5))
        findLabel.pack(side=tk.LEFT, padx=(rootWidth/5, 0))
    else:
        findLabel.pack()
    frames[3].pack(fill="x", padx=10, pady=(3, 0))  # inadvertently packed twice to have less y padding

    findListbox = tk.Listbox(frames[4], selectmode="multiple", exportselection=0, width=listboxWidth, height=listboxHeight)
    for findProcedureName in FIND_PROCEDURES.keys():
        findListbox.insert(tk.END, findProcedureName)
    findListbox.select_set(0)
    findListbox.config(font=fontSmall)
    if isAdmin:
        fixListbox = tk.Listbox(frames[4], exportselection=0, width=listboxWidth, height=listboxHeight)
        for fixProcedureName in FIX_PROCEDURES.keys():
            fixListbox.insert(tk.END, fixProcedureName)
        fixListbox.select_set(0)
        fixListbox.config(font=fontSmall)
        fixListbox.pack(side=tk.RIGHT)
        findListbox.pack(side=tk.LEFT)
    else:
        findListbox.pack()
    frames[4].pack(fill="x", padx=10, pady=(0, 3))

    if isAdmin:
        parameterLabel = tk.Label(frames[5], text="Parameter:", font=fontGeneral)
        argumentEntry = tk.Entry(frames[5], textvariable=parameterVar, width=rootWidth, font=fontSmall)
        parameterLabel.pack(side=tk.LEFT)
        argumentEntry.pack(side=tk.LEFT)

    includeSubfoldersCheckbutton = tk.Checkbutton(frames[6], text="Include Subfolders", variable=includeSubFoldersState, font=fontGeneral)
    includeSubfoldersCheckbutton.pack()
    if isAdmin:
        modifyCheckbutton = tk.Checkbutton(frames[6], text="Allow Modify", variable=modifyState, font=fontGeneral)
        modifyCheckbutton.pack(padx=(0, 50))

    executeButton = tk.Button(frames[7], text="Execute", command=launchController, width=finalButtonsWidth, font=fontGeneral)
    executeButton.pack()
    frames[7].configure(width=rootWidth/2)
    frames[7].pack(side=tk.LEFT, expand=True)

    resultsButton = tk.Button(frames[8], text="Results", command=openResultsDirectory, width=finalButtonsWidth, font=fontGeneral)
    resultsButton.pack()
    frames[8].configure(width=rootWidth/2)
    frames[8].pack(side=tk.LEFT, expand=True)


    root.mainloop()
        

def main():
    view(True)


if __name__ == "__main__":
    main()
