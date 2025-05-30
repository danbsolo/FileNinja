import tkinter as tk
from tkinter import filedialog
import os
from defs import *
import sys
from idlelib.tooltip import Hovertip
import threading
import filesScannedSharedVar
import control
import json



def launchView(isAdmin: bool):
    def launchControllerWorker():        
        currentState.set(102)  # 102 == The HTTP response code for "Still processing"

        exitStatus = control.launchController(dirAbsoluteVar.get(), bool(includeSubdirectoriesState.get()), bool(allowModifyState.get()), bool(includeHiddenFilesState.get()), bool(addRecommendationsState.get()),
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
            if exitStatus in EXIT_STATUS_CODES:
                if exitStatus == 0:
                    return
                errorMessage = EXIT_STATUS_CODES[exitStatus]
            else:
                # In this case, "exitStatus" is a tuple, where the first element is -999 and the second element is the output from "traceback.format_exc()"
                errorMessage = exitStatus[1]
                exitStatus = exitStatus[0]

            tk.messagebox.showerror(f"Error: {str(exitStatus)}", errorMessage)
        else:
            # Otherwise check again after the specified number of milliseconds.
            filesScanned = filesScannedSharedVar.FILES_SCANNED
            root.title(f"{FILE_NINJA} ({filesScanned})")
            executeButton.config(text=f"{filesScanned} files")
            scheduleCheckIfDone(t)
    

    def launchController():            
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


    def queryExcludeDirectory():
        potentialExcludedDirectory = filedialog.askdirectory(title="Browse to EXCLUDE",
                                                             initialdir=dirAbsoluteVar.get(),
                                                             mustexist=True)

        if (not potentialExcludedDirectory) or (potentialExcludedDirectory in excludedDirs):
            return
        
        excludeDirectory(potentialExcludedDirectory)

    def excludeDirectory(directoryToBeExcluded):
        excludedDirs.append(directoryToBeExcluded)
        excludeListbox.insert(tk.END, directoryToBeExcluded)
        
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
            sys.exit()  # Force close
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
        if "insertbackground" in widgetAttributes:
            widget.config(insertbackground=fgColor)

        for child in widget.winfo_children():
            changeChildrenColor(child, bgColor, fgColor, activeBgColor,selectColor)


    def saveSettingsIntoJSON():
        settings = {
            "dirAbsolute": dirAbsoluteVar.get(),
            "includeSubdirectories": bool(includeSubdirectoriesState.get()),
            "allowModify": bool(allowModifyState.get()),
            "includeHiddenFiles": bool(includeHiddenFilesState.get()),
            "addRecommendations": bool(addRecommendationsState.get()),
            "selectedFindProcedures": [findListbox.get(fm) for fm in findListbox.curselection()],
            "selectedFixProcedures": [fixListbox.get(fm) for fm in fixListbox.curselection()] if isAdmin else [],
            "argUnprocessed": parameterVar.get(),
            "excludedDirs": excludedDirs
        }

        # differentiate between running an exe or python
        if getattr(sys, 'frozen', False):
            isExe = True
            appDir = os.path.dirname(sys.executable)
        else:
            isExe = False
            appDir = os.path.dirname(os.path.abspath(__file__))

        jsonFilepath = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON files", "*.json")],
            title="Save File-Ninja settings as...",
            initialdir=appDir
        )

        if not jsonFilepath: return

        with open(jsonFilepath, "w") as f:
            json.dump(settings, f, indent=4)

        if not tk.messagebox.askyesno("Create .bat?", "Create corresponding .bat file?"):
            return
        
        pathSansFile = os.path.dirname(jsonFilepath)
        jsonFilename = os.path.basename(jsonFilepath)
        filename, _ = os.path.splitext(jsonFilename)

        # TODO: Probably get name of the current file executing rather than hard-coding
        if isExe:
            command = ""
            appPath = f"{appDir}\\File-Ninja-Control.exe"
        else:
            command = "python "
            appPath = f"{appDir}\\control.py"

        with open(f"{pathSansFile}\\{filename}.bat", "w") as f:
            f.write(
            f'@echo off\n\
set "batchFilename=%~n0"\n\
{command}"{appPath}" "%batchFilename%.json"\n\
pause')



    def loadSettingsFromJSON():
        filePath = filedialog.askopenfilename(title="Load File-Ninja settings",
                            filetypes=[("JSON files", "*.json")])
        
        if not filePath: return

        with open(filePath, "r") as f:
            settings = json.load(f)
        
         
        dirAbsoluteVar.set(settings["dirAbsolute"])
        includeSubdirectoriesState.set(settings["includeSubdirectories"])
        allowModifyState.set(settings["allowModify"])
        includeHiddenFilesState.set(settings["includeHiddenFiles"]) 
        addRecommendationsState.set(settings["addRecommendations"])

        findListbox.selection_clear(0, tk.END)
        for i in range(findListbox.size()):
            item = findListbox.get(i)
            if item in settings["selectedFindProcedures"]:
                findListbox.selection_set(i)

        fixListbox.selection_clear(0, tk.END)
        for i in range(fixListbox.size()):
            item = fixListbox.get(i)
            if item in settings["selectedFixProcedures"]:
                fixListbox.selection_set(i)

        parameterVar.set(settings["argUnprocessed"])
        
        excludeListbox.delete(0, tk.END)
        for item in settings["excludedDirs"]:
            excludeDirectory(item)
        


    listboxHeight = max(len(FIND_PROCEDURES_DISPLAY), len(FIX_PROCEDURES_DISPLAY)) + 1
    listboxHeightMultiplier = 17

    # root window stuff
    root = tk.Tk()
    root.title(FILE_NINJA)
    root.resizable(0, 0)
    rootWidth = 500 if isAdmin else 345
    rootHeight = (listboxHeight * listboxHeightMultiplier) + (425 if isAdmin else 310)
    root.geometry("{}x{}".format(rootWidth, rootHeight))

    # The following line of code breaks Hovertips. It just does.
    # if isAdmin: root.attributes('-topmost', True)  # keeps root window at top layer
        
    frames = []
    for i in range(11):
        frames.append(tk.Frame(root, bd=0, relief=tk.SOLID))
        frames[i].pack(fill="x", padx=10, pady=3)

    # aesthetic/layout variables
    fontType = "None"
    fontSize = 15
    fontGeneral = (fontType, fontSize)
    fontSmall = (fontType, int(fontSize/3*2))
    listboxWidth = int(rootWidth/15) if isAdmin else int(rootWidth/10)  # listboxHeight defined above
    finalButtonsWidth = 20 if isAdmin else 12  # HARD CODED WIDTH
    qolButtonsWidth = 12 # ALSO HARD CODED
    tooltipHoverDelay = 0

    # data variables
    dirAbsoluteVar = tk.StringVar()
    parameterVar = tk.StringVar()
    includeSubdirectoriesState = tk.IntVar(value=1)
    allowModifyState = tk.IntVar(value=0)
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
    
    excludeButton = tk.Button(frames[2], text="Browse to Exclude", command=queryExcludeDirectory, font=fontGeneral, width=rootWidth)
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

    includeSubfoldersCheckbutton = tk.Checkbutton(frames[6], text="Include Subdirectories", variable=includeSubdirectoriesState, font=fontGeneral)
    includeSubfoldersCheckbutton.pack(side=tk.LEFT)
    if isAdmin:
        modifyCheckbutton = tk.Checkbutton(frames[6], text="Allow Modify", variable=allowModifyState, font=fontGeneral)
        modifyCheckbutton.pack(side=tk.LEFT)  # , padx=(0, 50)
    else:
        readMeButton = tk.Button(frames[6], text="README", command=openReadMe, font=fontGeneral)
        readMeButton.pack(side=tk.LEFT)

    if isAdmin:
        includeHiddenFilesCheckbutton = tk.Checkbutton(frames[7], text="Include Hidden Files", variable=includeHiddenFilesState, font=fontGeneral)
        includeHiddenFilesCheckbutton.pack(side=tk.LEFT)
        addRecommendationsButton = tk.Checkbutton(frames[7], text="Add Recommendations~", variable=addRecommendationsState, font=fontGeneral)
        addRecommendationsButton.pack(side=tk.LEFT)

    if isAdmin:
        readMeButton = tk.Button(frames[8], text="README", command=openReadMe, width=qolButtonsWidth, font=fontGeneral)
        saveSettingsButton = tk.Button(frames[8], text="Save Settings", command=saveSettingsIntoJSON, width=qolButtonsWidth, font=fontGeneral)
        loadSettingsButton = tk.Button(frames[8], text="Load Settings", command=loadSettingsFromJSON, width=qolButtonsWidth, font=fontGeneral)
        readMeButton.pack(padx=(17, 0), side=tk.LEFT)
        saveSettingsButton.pack(padx=(10, 0), side=tk.LEFT)
        loadSettingsButton.pack(padx=(10, 0), side=tk.LEFT)

    executeButton = tk.Button(frames[9], text="Execute", command=launchController, width=finalButtonsWidth, font=fontGeneral)
    executeButton.pack()
    frames[9].configure(width=rootWidth/2)
    frames[9].pack(side=tk.LEFT, expand=True)

    resultsButton = tk.Button(frames[10], text="Results", command=openResultsDirectory, width=finalButtonsWidth, font=fontGeneral)
    resultsButton.pack()
    frames[10].configure(width=rootWidth/2)
    frames[10].pack(side=tk.LEFT, expand=True)


    # tool tips
    browseTip = Hovertip(browseButton, "Browse to select a directory.", hover_delay=tooltipHoverDelay)  
    dirHeaderTip = Hovertip(dirHeaderLabel, "Currently selected directory.", hover_delay=tooltipHoverDelay)
    excludeTip = Hovertip(excludeButton, "Browse to exclude subdirectories of currently selected directory.", hover_delay=tooltipHoverDelay)
    findTip = Hovertip(findLabel, "Run a Find procedure.", hover_delay=tooltipHoverDelay)
    includeSubdirectoriesTip = Hovertip(includeSubfoldersCheckbutton, "Dive into all subdirectories, other than those excluded.", hover_delay=tooltipHoverDelay)
    readMeTip = Hovertip(readMeButton, "Open README file.", hover_delay=tooltipHoverDelay)
    executeTip = Hovertip(executeButton, "Execute program.", hover_delay=tooltipHoverDelay)
    resultsTip = Hovertip(resultsButton, "Open results folder.", hover_delay=tooltipHoverDelay)
    if isAdmin: 
        fixTip = Hovertip(fixLabel, "Run a Fix procedure.", hover_delay=tooltipHoverDelay)
        parameterTip = Hovertip(parameterLabel, "\"#\" = requires argument input.\nInput a number, string, etc. Required for some procedures.", hover_delay=tooltipHoverDelay)
        allowModifyTip = Hovertip(modifyCheckbutton, "Unless you understand the consequences of this option, leave this off.", hover_delay=tooltipHoverDelay)
        includeHiddenFilesTip = Hovertip(includeHiddenFilesCheckbutton, "Include hidden files in Find procedure output. Fix procedures ignore hidden files by default.", hover_delay=tooltipHoverDelay)
        addRecommendationsTip = Hovertip(addRecommendationsButton, "\"~\" = has recommendation option.\nAdd recommendations to some procedures.", hover_delay=tooltipHoverDelay)
        saveSettingsTip = Hovertip(saveSettingsButton, "Save settings into a JSON file.", hover_delay=tooltipHoverDelay)
        loadSettingsTip = Hovertip(loadSettingsButton, "Load settings from a JSON file.", hover_delay=tooltipHoverDelay)


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
