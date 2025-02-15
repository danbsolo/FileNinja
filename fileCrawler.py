import tkinter as tk
from tkinter import filedialog
import os
from workbookManager import WorkbookManager
from datetime import datetime
from defs import *
from sys import exit
from idlelib.tooltip import Hovertip
import threading


def control(dirAbsolute:str, includeSubfolders:bool, modify:bool, selectedFindProcedures:list[str], selectedFixProcedures:list[str], unprocessedArg:str, excludedDirs:set[str]):
    if (not dirAbsolute): return -2
    
    # make it all backslashes, not forward slashes. This is to make it homogenous with os.walk output
    dirAbsolute = dirAbsolute.replace("/", "\\")


    # Create RESULTS directory name if does not exist
    try: os.mkdir(RESULTS_DIRECTORY)
    except: pass

    # >>> "FolderName-YYYY-mm-dd-HH-MM-SS.xlsx"
    workbookPathName = RESULTS_DIRECTORY + "\\" + dirAbsolute.split("\\")[-1] + "-" \
    + datetime.today().strftime("%Y-%m-%d-%H-%M-%S") + ".xlsx"

    # Initialize objects
    wbm = WorkbookManager(workbookPathName)
    setWorkbookManager(wbm)

    ## Errors if this file already exists and is currently opened
    # try:
    #    fileHandler = open(workbookPathName, 'w')
    #    fileHandler.close()
    # except PermissionError:
    #    return -1

    
    # Set findProcedures and fixProcedure
    if modify and len(selectedFixProcedures) > 1:
        return -5


    for findProcedureName in selectedFindProcedures:
        wbm.addFindProcedure(FIND_PROCEDURES[findProcedureName])


    splitArgs = unprocessedArg.split("/")
    splitArgsLen = len(splitArgs)
    currentArg = 0

    for fixProcedureName in selectedFixProcedures:
        arg = None
        
        if FIX_PROCEDURES[fixProcedureName].validatorFunction:
            if currentArg >= splitArgsLen:
                return -6
            
            arg = splitArgs[currentArg]
            currentArg += 1
        
        if (not wbm.addFixProcedure(FIX_PROCEDURES[fixProcedureName], modify, arg)):
            return -3


    # print(splitArgs)
    # print(wbm.fixSheetsDict, wbm.fixProcedureArgs, wbm.fixProcedureFunctions, sep="\n"*2, end="\n"*3)
    # return 0


    #if selectedFixProcedure != NULL_OPTION:
    #    fixProcedureObject = FIX_PROCEDURES[selectedFixProcedure]
    #
    #    if not wbm.setFixArg(fixProcedureObject, unprocessedArg):
    #        return -3
    #    
    #    wbm.setFixProcedure(fixProcedureObject, modify)
                    
    wbm.styleSummarySheet(dirAbsolute, includeSubfolders, modify)    

    # print("\nCreating " + workbookPathName + "...")
    if (includeSubfolders):
        if (not excludedDirs):
            wbm.folderCrawl(os.walk(dirAbsolute), [])

        else:
            # Run checks on excludedDirs
            for i in range(len(excludedDirs)):
                excludedDirs[i] = excludedDirs[i].replace("/", "\\")

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
                # subDirAbsolute = subDirAbsolute.replace("\\", "/")
                
                ### NOTE: Should it ignore hidden folders. i.e. folders that begin with "."?

                for exDir in excludedDirs:
                    if subDirAbsolute.startswith(exDir):  # exclude by subDirAbsolute to be precise
                        isDirIncluded = False
                        break

                if isDirIncluded:
                    walkResults.append((subDirAbsolute, subDirFolders, subDirFiles))
                else:
                    subDirFolders[:] = []  # stop traversal of this folder and its subfolders
            
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
    def launchControllerWorker():
        exitCodeList.append("FillerText")
        
        exitCode = control(dirAbsoluteVar.get(), bool(includeSubFoldersState.get()), bool(modifyState.get()), 
                    [findListbox.get(fm) for fm in findListbox.curselection()],
                    [fixListbox.get(fm) for fm in fixListbox.curselection()] if isAdmin else [],
                    parameterVar.get(),
                    excludedDirs)

        exitCodeList.append(exitCode)

    def schedule_check(t):
        root.after(500, check_if_done, t)


    def check_if_done(t):
        # If the thread has finished, re-enable the button and show a message.
        if not t.is_alive():
            root.title(FILE_CRAWLER)
            
            executeButton["state"] = "normal"

            exitCodeList.pop(0)
            exitCode = exitCodeList.pop(0)
            
            
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
            elif (exitCode == -5):
                errorMessage = "Invalid settings. Cannot run multiple Fix Procedures when modify is checked."
            elif (exitCode == -6):
                errorMessage = "Invalid arguments. Separate with \"/\""
            else:
                errorMessage = "An error has occurred."

            tk.messagebox.showerror("Error: " + str(exitCode), errorMessage)
            
        else:
            # Otherwise check again the specified number of milliseconds.
            schedule_check(t)
    
    def launchController():
        if modifyState.get() and not tk.messagebox.askyesnocancel("Allow Modify?", "You have chosen to modify items. This is an IRREVERSIBLE action. Are you sure?"):
            # print("Aborted. Continuing selection.")
            return
            
        root.title(FILE_CRAWLER + ": CURRENTLY RUNNING...")

        executeButton["state"] = "disabled"
        executionThread = threading.Thread(target=launchControllerWorker)
        executionThread.daemon = True  # When the main thread closes, this daemon thread will also close alongside it
        executionThread.start()

        schedule_check(executionThread)


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

        excludedDirs.append(potentialExcludedDirectory)
        excludeListbox.insert(tk.END, potentialExcludedDirectory)
        
        newHeight = len(excludedDirs) + 1
        excludeListbox.config(height = newHeight)
        root.geometry("{}x{}".format(rootWidth, rootHeight + (newHeight * listboxHeightMultiplier)))


    def removeExcludedDirectory():
        selection = excludeListbox.curselection()

        if not selection:
            return
        
        toBeRemovedIndex = selection[0]
        excludeListbox.delete(toBeRemovedIndex)
        del excludedDirs[toBeRemovedIndex]      

        newHeight = len(excludedDirs) + 1
        excludeListbox.config(height = newHeight)
        root.geometry("{}x{}".format(rootWidth, rootHeight + (newHeight * listboxHeightMultiplier)))
      

        
    def openResultsDirectory():
        if os.path.exists(RESULTS_DIRECTORY):
            os.startfile(RESULTS_DIRECTORY)
        else:
            tk.messagebox.showinfo("Directory DNE", "Directory \"" + RESULTS_DIRECTORY + "\" does not exist. Try executing a file crawl first.")


    def closeWindow():
        if not exitCodeList:
            root.destroy()
        else:
            # Force close
            exit()



    exitCodeList = []  # super hacky
    excludedDirs = []

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

    # The following line of code breaks Hovertips. I do not know why.
    # if isAdmin: root.attributes('-topmost', True)  # keeps root window at top layer

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
    tooltipHoverDelay = 0

    # data variables
    dirAbsoluteVar = tk.StringVar()
    parameterVar = tk.StringVar()
    includeSubFoldersState = tk.IntVar()
    modifyState = tk.IntVar()


    # widgets
    browseButton = tk.Button(frames[0], text="Browse to Select", command=selectDirectory, font=fontGeneral, width=rootWidth)
    browseButton.pack()
    browseTip = Hovertip(browseButton, "Browse to select a directory.", hover_delay=tooltipHoverDelay)

    dirHeaderLabel = tk.Label(frames[1], text = "Directory:", font=fontGeneral)
    dirLabel = tk.Label(frames[1], textvariable=dirAbsoluteVar, font=fontSmall, anchor="e") 
    dirHeaderLabel.pack(side=tk.LEFT)
    dirLabel.pack(side=tk.LEFT)
    dirHeaderTip = Hovertip(dirHeaderLabel, "Currently selected directory.", hover_delay=tooltipHoverDelay)
    

    excludeButton = tk.Button(frames[2], text="Browse to Exclude", command=excludeDirectory, font=fontGeneral, width=rootWidth)
    excludeScrollbar = tk.Scrollbar(frames[2], orient=tk.HORIZONTAL)
    excludeListbox = tk.Listbox(frames[2], exportselection=0, width=rootWidth, height=0, xscrollcommand=excludeScrollbar.set)
    excludeScrollbar.config(command=excludeListbox.xview)
    excludeListbox.bind("<Double-Button-1>", lambda _: removeExcludedDirectory()) # double left click
    excludeListbox.bind("<Button-3>", lambda _: removeExcludedDirectory()) # right click
    excludeButton.pack()
    excludeListbox.pack()
    excludeScrollbar.pack()
    excludeTip = Hovertip(excludeButton, "Browse to exclude subfolders of currently selected directory.", hover_delay=tooltipHoverDelay)
    
    findLabel = tk.Label(frames[3], text="Find", font=fontGeneral)
    if isAdmin:
        fixLabel = tk.Label(frames[3], text="Fix", font=fontGeneral)
        fixLabel.pack(side=tk.RIGHT, padx=(0, rootWidth/5))
        findLabel.pack(side=tk.LEFT, padx=(rootWidth/5, 0))
        fixTip = Hovertip(fixLabel, "Run a Fix procedure.\nCheck the HELPME.txt file for more info.", hover_delay=tooltipHoverDelay)
    else:
        findLabel.pack()
    frames[3].pack(fill="x", padx=10, pady=(3, 0))  # inadvertently packed twice to have less y padding
    findTip = Hovertip(findLabel, "Run a Find procedure.\nCheck the HELPME.txt file for more info.", hover_delay=tooltipHoverDelay)

    findListbox = tk.Listbox(frames[4], selectmode="multiple", exportselection=0, width=listboxWidth, height=listboxHeight)
    for findProcedureName in FIND_PROCEDURES.keys():
        findListbox.insert(tk.END, findProcedureName)
    findListbox.select_set(0)
    findListbox.config(font=fontSmall)
    if isAdmin:
        fixListbox = tk.Listbox(frames[4], selectmode="multiple", exportselection=0, width=listboxWidth, height=listboxHeight)
        for fixProcedureName in FIX_PROCEDURES.keys():
            fixListbox.insert(tk.END, fixProcedureName)
        # fixListbox.select_set(0)
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
        parameterTip = Hovertip(parameterLabel, "Input a number, string, etc. Required for some Fix procedures.", hover_delay=tooltipHoverDelay)

    includeSubfoldersCheckbutton = tk.Checkbutton(frames[6], text="Include Subfolders", variable=includeSubFoldersState, font=fontGeneral)
    includeSubfoldersCheckbutton.pack()
    if isAdmin:
        modifyCheckbutton = tk.Checkbutton(frames[6], text="Allow Modify", variable=modifyState, font=fontGeneral)
        modifyCheckbutton.pack(padx=(0, 50))
        modifyTip = Hovertip(modifyCheckbutton, "Unless you understand the consequences of this feature, leave this off.", hover_delay=tooltipHoverDelay)
    includeSubfoldersTip = Hovertip(includeSubfoldersCheckbutton, "Turn on to also delve into all subfolders, other than those excluded.", hover_delay=tooltipHoverDelay)

    executeButton = tk.Button(frames[7], text="Execute", command=launchController, width=finalButtonsWidth, font=fontGeneral)
    executeButton.pack()
    frames[7].configure(width=rootWidth/2)
    frames[7].pack(side=tk.LEFT, expand=True)
    executeTip = Hovertip(executeButton, "Execute the program.", hover_delay=tooltipHoverDelay)

    resultsButton = tk.Button(frames[8], text="Results", command=openResultsDirectory, width=finalButtonsWidth, font=fontGeneral)
    resultsButton.pack()
    frames[8].configure(width=rootWidth/2)
    frames[8].pack(side=tk.LEFT, expand=True)
    resultsTip = Hovertip(resultsButton, "Open folder containing all excel files of previous executions.", hover_delay=tooltipHoverDelay)


    root.protocol("WM_DELETE_WINDOW", closeWindow)
    

    root.mainloop()
        

def main():
    view(True)


if __name__ == "__main__":
    main()
