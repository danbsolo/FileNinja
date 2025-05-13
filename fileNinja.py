import tkinter as tk
from tkinter import filedialog
import os
from workbookManager import WorkbookManager
from datetime import datetime
from procedureDefinitions import *
import procedureFunctions
import findFunctions
import fixFunctions
from defs import *
from sys import exit
from idlelib.tooltip import Hovertip
import threading
import traceback
import filesScannedSharedVar


def control(dirAbsolute:str, includeSubfolders:bool, allowModify:bool, includeHiddenFiles:bool, addRecommendations:bool, selectedFindProcedures:list[str], selectedFixProcedures:list[str], argUnprocessed:str, excludedDirs:set[str]):
    if (not dirAbsolute): return -2

    # If multiple fix procedures are selected and allowModify is checked, exit
    if allowModify and len(selectedFixProcedures) > 1:
        return -5
    
    # If allowModify and addRecommendations are both turned on, even if it would be of no consequence, exit
    if allowModify and addRecommendations:
        return -7
    
    
    
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
    procedureFunctions.setWorkbookManager(wbm)
    findFunctions.setWorkbookManager(wbm)
    fixFunctions.setWorkbookManager(wbm)
    
    # Add selected findProcedures
    for findProcedureName in selectedFindProcedures:
        wbm.addFindProcedure(FIND_PROCEDURES[findProcedureName], addRecommendations)

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
        
        if (not wbm.addFixProcedure(FIX_PROCEDURES[fixProcedureName], allowModify, addRecommendations, arg)):
            return -3


    try:
        wbm.initiateCrawl(dirAbsolute, includeSubfolders, allowModify, includeHiddenFiles, addRecommendations, excludedDirs)
        wbm.close()
        os.startfile(workbookPathName)
    except Exception as e:
        return traceback.format_exc()
    
    return 0



def view(isAdmin: bool):
    def launchControllerWorker():        
        currentState.set(102)  # 102 == The HTTP response code for "Still processing"

        exitStatus = control(dirAbsoluteVar.get(), bool(includeSubFoldersState.get()), bool(modifyState.get()), bool(includeHiddenFilesState.get()), bool(addRecommendationsState.get()),
                    [findListbox.get(fm) for fm in findListbox.curselection()],
                    [fixListbox.get(fm) for fm in fixListbox.curselection()] if isAdmin else [],
                    parameterVar.get(),
                    excludedDirs)

        currentState.set(exitStatus)

    def scheduleCheckIfDone(t):
        root.after(500, checkIfDone, t)

    def checkIfDone(t):
        # If the thread has finished
        if not t.is_alive():
            root.title(FILE_NINJA)
            executeButton.config(text="Execute", state="normal")
            filesScannedSharedVar.FILES_SCANNED = 0
            filesScannedSharedVar.TOTAL_FILES = 0

            try:
                exitStatus = int(currentState.get())
            except:
                exitStatus = currentState.get()

            errorMessage = ""
            if (exitStatus == 0):
                return
            elif (exitStatus == -1):
                errorMessage = "Could not open file. Close file and try again."
            elif (exitStatus == -2):
                errorMessage = "Invalid directory."
            elif (exitStatus == -3):
                errorMessage = "Invalid argument."
            elif (exitStatus == -4):
                errorMessage = "Invalid excluded directories."
            elif (exitStatus == -5):
                errorMessage = "Invalid settings. Cannot run multiple Fix Procedures when modify is checked."
            elif (exitStatus == -6):
                errorMessage = "Invalid arguments. Separate with \"/\""
            elif (exitStatus == -7):
                errorMessage = "Invalid settings. Cannot run recommendations and modifications simultaneously."
            else:
                # In this case, "exitStatus" is the text from "traceback.format_exc()", before we change it to -999
                errorMessage = f"TAKE A SCREENSHOT -- an error has occured.\n\n{exitStatus}"
                exitStatus = -999

            tk.messagebox.showerror(f"Error: {str(exitStatus)}", errorMessage)
        else:
            # Otherwise check again after the specified number of milliseconds.
            filesScanned = filesScannedSharedVar.FILES_SCANNED
            root.title(f"{FILE_NINJA} ({filesScanned})")
            executeButton.config(text=f"{filesScanned} files")
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

    
    def openReadMe():
        # try local
        if isAdmin:
            if os.path.exists(README_ADMIN):
                os.startfile(README_ADMIN)
                return
            # try HCS
            elif os.path.exists(HCS_ASSETS_PATH + README_ADMIN):
                os.startfile(HCS_ASSETS_PATH + README_ADMIN)
                return
            
        # Do not open readMeLite as Admin, even if it's available and readMeAdmin isn't
        else:
            if os.path.exists(README_LITE):
                os.startfile(README_LITE)
                return
            elif os.path.exists(HCS_ASSETS_PATH + README_LITE):
                os.startfile(HCS_ASSETS_PATH + README_LITE)
                return

        tk.messagebox.showinfo("README DNE", "README file does not exist.")

    def changeColorMode():
        global onDefaultColorMode
        if onDefaultColorMode:
            onDefaultColorMode = False
            changeChildrenColor(root, "#202124", "white", "gray20", "gray20")
        else:
            onDefaultColorMode = True
            changeChildrenColor(root, originalBg, originalFg, originalActiveBg, originalSelectColor)


    def changeChildrenColor(widget, bgColor, fgColor, activeBgColor, selectColor):
        widgetAttributes = widget.keys()

        if "bg" in widgetAttributes:
            widget.config(bg=bgColor)
        if "fg" in widgetAttributes:
            widget.config(fg=fgColor)
        if "activebackground" in widgetAttributes:
            widget.config(activebackground=activeBgColor)
        if "activeforeground" in widgetAttributes:
            widget.config(activeforeground=fgColor)
        if "selectcolor" in widgetAttributes:
            widget.config(selectcolor=selectColor)

        for child in widget.winfo_children():
            changeChildrenColor(child, bgColor, fgColor, activeBgColor,selectColor)
        

    listboxHeight = max(len(FIND_PROCEDURES_DISPLAY), len(FIX_PROCEDURES_DISPLAY)) + 1
    listboxHeightMultiplier = 17

    # root window stuff
    root = tk.Tk()
    root.title(FILE_NINJA)
    root.resizable(0, 0)
    rootWidth = 500 if isAdmin else 345
    rootHeight = (listboxHeight * listboxHeightMultiplier) + (385 if isAdmin else 310) # 345, 310
    root.geometry("{}x{}".format(rootWidth, rootHeight))

    # The following line of code breaks Hovertips. It just does.
    # if isAdmin: root.attributes('-topmost', True)  # keeps root window at top layer
        
    frames = []
    for i in range(10):
        frames.append(tk.Frame(root, bd=0, relief=tk.SOLID))
        frames[i].pack(fill="x", padx=10, pady=3)

    # aesthetic/layout variables
    fontType = "None"
    fontSize = 15
    fontGeneral = (fontType, fontSize)
    fontSmall = (fontType, int(fontSize/3*2))
    listboxWidth = int(rootWidth/15) if isAdmin else int(rootWidth/10)  # listboxHeight defined above
    finalButtonsWidth = 20 if isAdmin else 12  # HARD CODED WIDTH
    tooltipHoverDelay = 0

    # data variables
    dirAbsoluteVar = tk.StringVar()
    parameterVar = tk.StringVar()
    includeSubFoldersState = tk.IntVar(value=1)
    modifyState = tk.IntVar(value=0)
    includeHiddenFilesState = tk.IntVar(value=isAdmin)
    addRecommendationsState = tk.IntVar(value=0)
    currentState = tk.StringVar(value=0)
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

    findListbox = tk.Listbox(frames[4], selectmode=tk.MULTIPLE, exportselection=0, width=listboxWidth, height=listboxHeight)
    for findProcedureName in FIND_PROCEDURES_DISPLAY:
        findListbox.insert(tk.END, findProcedureName)
    findListbox.select_set(0)
    findListbox.config(font=fontSmall)

    if isAdmin:
        fixListbox = tk.Listbox(frames[4], selectmode=tk.MULTIPLE, exportselection=0, width=listboxWidth, height=listboxHeight)
        for fixProcedureName in FIX_PROCEDURES_DISPLAY:
            fixListbox.insert(tk.END, fixProcedureName)
        fixListbox.config(font=fontSmall)
        fixListbox.pack(side=tk.RIGHT)
        findListbox.pack(side=tk.LEFT)
    else:
        findListbox.pack()
    frames[4].pack(fill="x", padx=10, pady=(0, 3))

    if isAdmin:
        parameterLabel = tk.Label(frames[5], text="Parameter#:", font=fontGeneral)
        argumentEntry = tk.Entry(frames[5], textvariable=parameterVar, width=rootWidth, font=fontSmall)
        parameterLabel.pack(side=tk.LEFT)
        argumentEntry.pack(side=tk.LEFT)

    includeSubfoldersCheckbutton = tk.Checkbutton(frames[6], text="Include Subdirectories", variable=includeSubFoldersState, font=fontGeneral)
    includeSubfoldersCheckbutton.pack(side=tk.LEFT)
    if isAdmin:
        modifyCheckbutton = tk.Checkbutton(frames[6], text="Allow Modify", variable=modifyState, font=fontGeneral)
        modifyCheckbutton.pack(side=tk.LEFT)  # , padx=(0, 50)
    readMeButton = tk.Button(frames[6], text="README", command=openReadMe, width=rootWidth, font=fontGeneral)
    readMeButton.pack(side=tk.LEFT)

    if isAdmin:
        includeHiddenFilesCheckbutton = tk.Checkbutton(frames[7], text="Include Hidden Files", variable=includeHiddenFilesState, font=fontGeneral)
        includeHiddenFilesCheckbutton.pack(side=tk.LEFT)
        addRecommendationsButton = tk.Checkbutton(frames[7], text="Add Recommendations~", variable=addRecommendationsState, font=fontGeneral)
        addRecommendationsButton.pack(side=tk.LEFT)

    executeButton = tk.Button(frames[8], text="Execute", command=launchController, width=finalButtonsWidth, font=fontGeneral)
    executeButton.pack()
    frames[8].configure(width=rootWidth/2)
    frames[8].pack(side=tk.LEFT, expand=True)

    resultsButton = tk.Button(frames[9], text="Results", command=openResultsDirectory, width=finalButtonsWidth, font=fontGeneral)
    resultsButton.pack()
    frames[9].configure(width=rootWidth/2)
    frames[9].pack(side=tk.LEFT, expand=True)


    # tool tips
    browseTip = Hovertip(browseButton, "Browse to select a directory.", hover_delay=tooltipHoverDelay)  
    dirHeaderTip = Hovertip(dirHeaderLabel, "Currently selected directory.", hover_delay=tooltipHoverDelay)
    excludeTip = Hovertip(excludeButton, "Browse to exclude subdirectories of currently selected directory.", hover_delay=tooltipHoverDelay)
    findTip = Hovertip(findLabel, "Run a Find procedure.", hover_delay=tooltipHoverDelay)
    includeSubfoldersTip = Hovertip(includeSubfoldersCheckbutton, "Dive into all subdirectories, other than those excluded.", hover_delay=tooltipHoverDelay)
    readMeTip = Hovertip(readMeButton, "Open README file.", hover_delay=tooltipHoverDelay)
    executeTip = Hovertip(executeButton, "Execute program.", hover_delay=tooltipHoverDelay)
    resultsTip = Hovertip(resultsButton, "Open results folder.", hover_delay=tooltipHoverDelay)
    if isAdmin: 
        fixTip = Hovertip(fixLabel, "Run a Fix procedure.", hover_delay=tooltipHoverDelay)
        parameterTip = Hovertip(parameterLabel, "\"#\" = requires argument input.\nInput a number, string, etc. Required for some procedures.", hover_delay=tooltipHoverDelay)
        modifyTip = Hovertip(modifyCheckbutton, "Unless you understand the consequences of this option, leave this off.", hover_delay=tooltipHoverDelay)
        includeHiddenFilesTip = Hovertip(includeHiddenFilesCheckbutton, "Include hidden files in Find procedure output. Fix procedures ignore hidden files by default.", hover_delay=tooltipHoverDelay)
        addRecommendationsTip = Hovertip(addRecommendationsButton, "\"~\" = has recommendation option.\nAdd recommendations to some procedures.", hover_delay=tooltipHoverDelay)
    

    # color mode stuff
    global onDefaultColorMode
    onDefaultColorMode = True
    originalBg = browseButton.cget("bg")
    originalFg = browseButton.cget("fg")
    originalActiveBg = browseButton.cget("activebackground")
    originalSelectColor = includeSubfoldersCheckbutton.cget("selectcolor")


    # bindings
    excludeListbox.bind("<Double-Button-1>", lambda _: removeExcludedDirectory()) # double left click
    excludeListbox.bind("<Button-3>", lambda _: removeExcludedDirectory()) # right click
    if isAdmin: fixListbox.bind("<<ListboxSelect>>", onSelectFixlistbox)
    root.bind('<Control-Key-w>', lambda e: root.destroy())
    root.bind('<Control-Key-W>', lambda e: root.destroy())
    root.bind("<Button-2>", lambda _: changeColorMode()) # middle click
    root.protocol("WM_DELETE_WINDOW", closeWindow)


    # set icon image (if available)
    if os.path.exists(LOGO_PATH):
        logoImg = tk.PhotoImage(file=LOGO_PATH)
        root.iconphoto(False, logoImg)


    # premature call for default dark mode
    changeColorMode()


    root.mainloop()
        


def main():
    view(True)



if __name__ == "__main__":
    main()
