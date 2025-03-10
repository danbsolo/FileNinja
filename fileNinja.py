import tkinter as tk
from tkinter import filedialog
import os
from workbookManager import WorkbookManager
from datetime import datetime
from defs import *
from sys import exit
from idlelib.tooltip import Hovertip
import threading


def control(dirAbsolute:str, includeSubfolders:bool, allowModify:bool, selectedFindProcedures:list[str], selectedFixProcedures:list[str], argUnprocessed:str, excludedDirs:set[str]):
    if (not dirAbsolute): return -2

    # If multiple fix procedures are selected and allowModify is checked, exit
    if allowModify and len(selectedFixProcedures) > 1:
        return -5
    
    # make it all backslashes, not forward slashes. This is to make it homogenous with os.walk() output
    dirAbsolute = dirAbsolute.replace("/", "\\")

    # Run checks on excludedDirs if relevant
    if (includeSubfolders and excludedDirs):
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

    # Create RESULTS directory if it does not exist
    try: os.mkdir(RESULTS_DIRECTORY)
    except: pass

    # >>> "FolderName-YYYY-mm-dd-HH-MM-SS.xlsx"
    workbookPathName = RESULTS_DIRECTORY + "\\" + dirAbsolute.split("\\")[-1] + "-" \
    + datetime.today().strftime("%Y-%m-%d-%H-%M-%S") + ".xlsx"

    # Initialize objects
    wbm = WorkbookManager(workbookPathName)
    setWorkbookManager(wbm)
    
    # Add selected findProcedures
    for findProcedureName in selectedFindProcedures:
        wbm.addFindProcedure(FIND_PROCEDURES[findProcedureName])

    # Add selected fixProcedures
    argsList = argUnprocessed.split("/")
    argsListLength = len(argsList)
    currentArg = 0

    for fixProcedureName in selectedFixProcedures:
        arg = None
        
        if FIX_PROCEDURES[fixProcedureName].validatorFunction:
            if currentArg >= argsListLength:
                return -6
            
            arg = argsList[currentArg]
            currentArg += 1
        
        if (not wbm.addFixProcedure(FIX_PROCEDURES[fixProcedureName], allowModify, arg)):
            return -3


    wbm.initiateCrawl(dirAbsolute, includeSubfolders, allowModify, excludedDirs)

    wbm.close()
    os.startfile(workbookPathName)
    return 0



def view(isAdmin: bool):
    def launchControllerWorker():        
        currentState.set(102)  # 102 == The HTTP response code for "Still processing"

        exitCode = control(dirAbsoluteVar.get(), bool(includeSubFoldersState.get()), bool(modifyState.get()), 
                    [findListbox.get(fm) for fm in findListbox.curselection()],
                    [fixListbox.get(fm) for fm in fixListbox.curselection()] if isAdmin else [],
                    parameterVar.get(),
                    excludedDirs)

        currentState.set(exitCode)

    def scheduleCheckIfDone(t):
        root.after(500, checkIfDone, t)

    def checkIfDone(t):
        # If the thread has finished
        if not t.is_alive():
            root.title(FILE_NINJA)
            executeButton.config(text="Execute", state="normal")

            exitCode = currentState.get()
            
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
                exitCode = -999
                errorMessage = "An error has occurred."

            tk.messagebox.showerror(f"Error: {str(exitCode)}", errorMessage)
            
        else:
            # Otherwise check again after the specified number of milliseconds.
            scheduleCheckIfDone(t)
    

    def launchController():
        if modifyState.get() and not tk.messagebox.askyesnocancel("Allow Modify?", "You have chosen to modify items. This is an IRREVERSIBLE action. Are you sure?"):
            return
            
        root.title(FILE_NINJA + ": RUNNING...")

        executeButton.config(text="RUNNING....", state="disabled")

        executionThread = threading.Thread(target=launchControllerWorker)
        executionThread.daemon = True  # When the main thread closes, this daemon thread will also close alongside it
        executionThread.start()

        scheduleCheckIfDone(executionThread)


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
            tk.messagebox.showinfo("Directory DNE", "Directory \"" + RESULTS_DIRECTORY + "\" does not exist in current directory. Try executing the program first.")


    def closeWindow():
        if currentState.get() == 102:
            exit()  # Force close
        else:
            root.destroy()
            

    def onSelectFixlistbox(e):
        for selectedIndex in fixListbox.curselection():            
            # If the value is an empty string (which is just there for the spacing), deselect it immediately
            if fixListbox.get(selectedIndex) == "":
                fixListbox.selection_clear(selectedIndex)

    
    def openHelpMe():
        helpMeAdmin = "HELPME-Admin.txt"
        helpMeLite = "HELPME-Lite.txt"

        # try local
        if isAdmin:
            if os.path.exists(helpMeAdmin + "kek"):
                os.startfile(helpMeAdmin)
                return
            elif os.path.exists(HCS_ASSETS_PATH + helpMeAdmin):
                os.startfile(HCS_ASSETS_PATH + helpMeAdmin)
                return
        # Do not open helpMeLite as Admin, even if it's available and helpMeAdmin isn't
        # try HCS
        else:
            if os.path.exists(helpMeLite):
                os.startfile(helpMeLite)
                return
            elif os.path.exists(HCS_ASSETS_PATH + helpMeLite):
                os.startfile(HCS_ASSETS_PATH + helpMeLite)
                return

        tk.messagebox.showinfo("HELPME DNE", "HELPME file does not exist.")

        

    listboxHeight = max(len(FIND_PROCEDURES_DISPLAY), len(FIX_PROCEDURES_DISPLAY)) + 1
    listboxHeightMultiplier = 17

    # root window stuff
    root = tk.Tk()
    root.title(FILE_NINJA)
    root.resizable(0, 0)
    rootWidth = 500 if isAdmin else 310
    rootHeight = (listboxHeight * listboxHeightMultiplier) + (345 if isAdmin else 310)
    root.geometry("{}x{}".format(rootWidth, rootHeight))

    # The following line of code breaks Hovertips. It just does.
    # if isAdmin: root.attributes('-topmost', True)  # keeps root window at top layer
        
    frames = []
    for i in range(9):
        frames.append(tk.Frame(root, bd=0, relief=tk.SOLID))
        frames[i].pack(fill="x", padx=10, pady=3)

    # aesthetic/layout variables
    fontType = "None"
    fontSize = 15
    fontGeneral = (fontType, fontSize)
    fontSmall = (fontType, int(fontSize/3*2))
    listboxWidth = int(rootWidth/15) if isAdmin else int(rootWidth/10)  # listboxHeight defined above
    finalButtonsWidth = 20 if isAdmin else 10  # HARD CODED WIDTH
    tooltipHoverDelay = 0

    # data variables
    dirAbsoluteVar = tk.StringVar()
    parameterVar = tk.StringVar()
    includeSubFoldersState = tk.IntVar(value=1)
    modifyState = tk.IntVar(value=0)
    currentState = tk.IntVar(value=0)
    excludedDirs = []


    # widget layout
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
    for findProcedureName in FIND_PROCEDURES_DISPLAY:
        findListbox.insert(tk.END, findProcedureName)
    findListbox.select_set(0)
    findListbox.config(font=fontSmall)

    if isAdmin:
        fixListbox = tk.Listbox(frames[4], selectmode="multiple", exportselection=0, width=listboxWidth, height=listboxHeight)
        for fixProcedureName in FIX_PROCEDURES_DISPLAY:
            fixListbox.insert(tk.END, fixProcedureName)
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
    includeSubfoldersCheckbutton.pack(side=tk.LEFT)
    if isAdmin:
        modifyCheckbutton = tk.Checkbutton(frames[6], text="Allow Modify", variable=modifyState, font=fontGeneral)
        modifyCheckbutton.pack(side=tk.LEFT)  # , padx=(0, 50)
    helpMeButton = tk.Button(frames[6], text="HELPME", command=openHelpMe, width=rootWidth, font=fontGeneral)
    helpMeButton.pack(side=tk.LEFT)

    executeButton = tk.Button(frames[7], text="Execute", command=launchController, width=finalButtonsWidth, font=fontGeneral)
    executeButton.pack()
    frames[7].configure(width=rootWidth/2)
    frames[7].pack(side=tk.LEFT, expand=True)

    resultsButton = tk.Button(frames[8], text="Results", command=openResultsDirectory, width=finalButtonsWidth, font=fontGeneral)
    resultsButton.pack()
    frames[8].configure(width=rootWidth/2)
    frames[8].pack(side=tk.LEFT, expand=True)


    # tool tips
    browseTip = Hovertip(browseButton, "Browse to select a directory.", hover_delay=tooltipHoverDelay)  
    dirHeaderTip = Hovertip(dirHeaderLabel, "Currently selected directory.", hover_delay=tooltipHoverDelay)
    excludeTip = Hovertip(excludeButton, "Browse to exclude subfolders of currently selected directory.", hover_delay=tooltipHoverDelay)
    findTip = Hovertip(findLabel, "Run a Find procedure.\nCheck the HELPME.txt file for more info.", hover_delay=tooltipHoverDelay)
    if isAdmin: 
        fixTip = Hovertip(fixLabel, "Run a Fix procedure.\nCheck the HELPME.txt file for more info.", hover_delay=tooltipHoverDelay)
        parameterTip = Hovertip(parameterLabel, "Input a number, string, etc. Required for some Fix procedures.", hover_delay=tooltipHoverDelay)
        modifyTip = Hovertip(modifyCheckbutton, "Unless you understand the consequences of this feature, leave this off.", hover_delay=tooltipHoverDelay)
    includeSubfoldersTip = Hovertip(includeSubfoldersCheckbutton, "Turn on to also delve into all subfolders, other than those excluded.", hover_delay=tooltipHoverDelay)
    helpMeTip = Hovertip(helpMeButton, "Open HELPME file.", hover_delay=tooltipHoverDelay)
    executeTip = Hovertip(executeButton, "Execute the program.", hover_delay=tooltipHoverDelay)
    resultsTip = Hovertip(resultsButton, "Open folder containing all excel files of previous executions.", hover_delay=tooltipHoverDelay)


    # bindings
    excludeListbox.bind("<Double-Button-1>", lambda _: removeExcludedDirectory()) # double left click
    excludeListbox.bind("<Button-3>", lambda _: removeExcludedDirectory()) # right click
    if isAdmin: fixListbox.bind("<<ListboxSelect>>", onSelectFixlistbox)
    root.bind('<Control-Key-w>', lambda e: root.destroy())
    root.bind('<Control-Key-W>', lambda e: root.destroy())
    root.protocol("WM_DELETE_WINDOW", closeWindow)


    # set icon image (if available)
    logoPath = HCS_ASSETS_PATH + "File-Ninja-Logo-Square.png"
    if os.path.exists(logoPath):
        logoImg = tk.PhotoImage(file=logoPath)
        root.iconphoto(False, logoImg)


    root.mainloop()
        

def main():
    view(False)


if __name__ == "__main__":
    main()
