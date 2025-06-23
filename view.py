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
import common



def launchView(isAdmin: bool):
    def launchControllerWorker():
        nonlocal currentStatusPair

        # Double check that the user wants to allow modifications
        if allowModifyState.get() and not tk.messagebox.askyesnocancel(
            "Confirmation", 
            "Modifications are irreversible. By proceeding with the Fix procedure(s), you are confirming the action has been requested by the data owner and that there are no identified Litigation Holds (LIT-HOLD) or Access to Information and Privacy (ATIP) requests for in scope data. Proceed?"):
            currentStatusPair = (STATUS_SUCCESSFUL, None)
            return
        
        currentStatusPair = (STATUS_RUNNING, None)
        currentStatusPair = control.launchController(dirAbsoluteVar.get(), bool(includeSubdirectoriesState.get()), bool(allowModifyState.get()), bool(includeHiddenFilesState.get()), bool(addRecommendationsState.get()),
                    [findListbox.get(fm) for fm in findListbox.curselection()],
                    [fixListbox.get(fm) for fm in fixListbox.curselection()] if isAdmin else [],
                    parameterVar.get(),
                    excludedDirs,
                    excludedExtensionsVar.get())

    def scheduleCheckIfDone(t):
        root.after(500, checkIfDone, t)

    def checkIfDone(t):
        # If the thread has finished
        if not t.is_alive():
            root.title(FILE_NINJA)
            executeButton.config(text="Execute", state="normal")
            filesScannedSharedVar.FILES_SCANNED = 0

            nonlocal currentStatusPair
            exitCode, _ = currentStatusPair
            if exitCode == STATUS_SUCCESSFUL:
                return          
            tk.messagebox.showerror(f"Error: {exitCode}", common.interpretError(currentStatusPair))

            currentStatusPair = (STATUS_IDLE, None)
        
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
        if currentStatusPair[0] == STATUS_RUNNING:
            sys.exit()  # Force close
        else:
            root.destroy()  # End gracefully
            

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


    def changeToDarkMode(parentWidget):
        changeChildrenColor(parentWidget, "#202124", "white", "gray20", "gray20")

    def changeToLightMode(parentWidget):
        changeChildrenColor(parentWidget, "#F1F1F1", "#000000", "#F1F1F1", "#FFFFFF")
        
    def changeColorMode(parentWidget):
        nonlocal onDarkMode
        if onDarkMode:
            onDarkMode = False
            changeToLightMode(parentWidget)
        else:
            onDarkMode = True
            changeToDarkMode(parentWidget)

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
            DIR_ABSOLUTE_KEY: dirAbsoluteVar.get(),
            EXCLUDED_DIRS_KEY: excludedDirs,
            EXCLUDED_EXTENSIONS_KEY: excludedExtensionsVar.get(),
            SELECTED_FIND_PROCEDURES_KEY: [findListbox.get(fm) for fm in findListbox.curselection()],
            SELECTED_FIX_PROCEDURES_KEY: [fixListbox.get(fm) for fm in fixListbox.curselection()] if isAdmin else [],
            ARG_UNPROCESSED_KEY: parameterVar.get(),
            INCLUDE_SUBDIRECTORIES_KEY: bool(includeSubdirectoriesState.get()),
            ENABLE_MODIFICATIONS_KEY: bool(allowModifyState.get()),
            INCLUDE_HIDDEN_FILES_KEY: bool(includeHiddenFilesState.get()),
            ADD_RECOMMENDATIONS_KEY: bool(addRecommendationsState.get()),
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
            title="Save settings as...",
        )

        if not jsonFilepath: return

        with open(jsonFilepath, "w") as f:
            json.dump(settings, f, indent=4)

        if not tk.messagebox.askyesno("Create .bat?", "Create corresponding .bat file?"):
            return
        
        pathSansFile = os.path.dirname(jsonFilepath)
        jsonFilename = os.path.basename(jsonFilepath)
        filename, _ = os.path.splitext(jsonFilename)
        
        appPath = f"{appDir}\\{os.path.basename(sys.argv[0])}"  # File-Ninja-Control.exe/.py
        with open(f"{pathSansFile}\\{filename}.bat", "w") as f:
            f.write(
            f'@echo off\n\
set "batchFilename=%~n0"\n\
{"" if isExe else "python " }"{appPath}" "%batchFilename%.json"\n\
pause')


    def loadSettingsFromJSON():
        filePath = filedialog.askopenfilename(title="Load File-Ninja settings",
                            filetypes=[("JSON files", "*.json")])
        
        if not filePath: return

        settingsPair = common.loadSettingsFromJSON(filePath)
        if settingsPair[0] != STATUS_SUCCESSFUL:
            tk.messagebox.showerror("Invalid JSON file", common.interpretError(settingsPair))
            return
        settings = settingsPair[1]
         
        dirAbsoluteVar.set(settings[DIR_ABSOLUTE_KEY])
        includeSubdirectoriesState.set(settings[INCLUDE_SUBDIRECTORIES_KEY])
        allowModifyState.set(settings[ENABLE_MODIFICATIONS_KEY])
        includeHiddenFilesState.set(settings[INCLUDE_HIDDEN_FILES_KEY]) 
        addRecommendationsState.set(settings[ADD_RECOMMENDATIONS_KEY])
        excludedExtensionsVar.set(settings[EXCLUDED_EXTENSIONS_KEY])

        findListbox.selection_clear(0, tk.END)
        for i in range(findListbox.size()):
            item = findListbox.get(i)
            if item in settings[SELECTED_FIND_PROCEDURES_KEY]:
                findListbox.selection_set(i)

        fixListbox.selection_clear(0, tk.END)
        for i in range(fixListbox.size()):
            item = fixListbox.get(i)
            if item in settings[SELECTED_FIX_PROCEDURES_KEY]:
                fixListbox.selection_set(i)

        parameterVar.set(settings[ARG_UNPROCESSED_KEY])
        
        excludeListbox.delete(0, tk.END)
        for item in settings[EXCLUDED_DIRS_KEY]:
            excludeDirectory(item)


    def openAdvancedOptionsWindow():
        nonlocal advancedOptionsWindow

        # If this window is already open, lift it to the forefront
        if advancedOptionsWindow and advancedOptionsWindow.winfo_exists():
            advancedOptionsWindow.deiconify()
            advancedOptionsWindow.lift()
            return

        advancedOptionsWindow = tk.Toplevel(root)
        advancedOptionsWindow.title(f"Advanced Options")
        advancedOptionsWindow.resizable(0, 0)
        # advancedOptionsWindow.geometry("500x200")
        advancedOptionsWindow.geometry("+{}+{}".format(root.winfo_pointerx()+rootWidth//2, root.winfo_pointery()))

        if logoImg:
            advancedOptionsWindow.iconphoto(False, logoImg)

        frame1 = tk.Frame(advancedOptionsWindow)
        frame1.pack(fill="x", padx=10, pady=3)
        frame2 = tk.Frame(advancedOptionsWindow)
        frame2.pack(fill="x", padx=10, pady=3)

        includeSubdirectoriesCheckbutton = tk.Checkbutton(frame1, text="Include Subdirectories", variable=includeSubdirectoriesState, font=fontGeneral)
        includeHiddenFilesCheckbutton = tk.Checkbutton(frame1, text="Include Hidden Files", variable=includeHiddenFilesState, font=fontGeneral)
        includeSubdirectoriesCheckbutton.pack(side=tk.TOP)
        includeHiddenFilesCheckbutton.pack(side=tk.TOP)

        excludedExtensionsLabel = tk.Label(frame2, text="Extensions to Exclude:", font=fontGeneral)
        excludedExtensionsEntry = tk.Entry(frame2, textvariable=excludedExtensionsVar, width=45, font=fontSmall)
        excludedExtensionsLabel.pack(side=tk.LEFT)
        excludedExtensionsEntry.pack(side=tk.LEFT)
        
        Hovertip(includeSubdirectoriesCheckbutton, "Dive into all subdirectories, other than those excluded.", hover_delay=tooltipHoverDelay) # includeSubdirectoriesTip
        Hovertip(includeHiddenFilesCheckbutton, "Include hidden files in Find procedure output. Fix procedures always ignore hidden files.", hover_delay=tooltipHoverDelay) # includeHiddenFilesTip
        Hovertip(excludedExtensionsLabel, "Comma separated list of extensions to exclude from scan.", hover_delay=tooltipHoverDelay)

        if onDarkMode:
            changeToDarkMode(advancedOptionsWindow)
        else:
            changeToLightMode(advancedOptionsWindow)



    listboxHeight = max(len(FIND_PROCEDURES_DISPLAY), len(FIX_PROCEDURES_DISPLAY)) + 1
    listboxHeightMultiplier = 17

    # root window stuff
    root = tk.Tk()
    root.title(FILE_NINJA)
    root.resizable(0, 0)
    rootWidth = 500 if isAdmin else 365
    rootHeight = (listboxHeight * listboxHeightMultiplier) + (420 if isAdmin else 295)
    root.geometry("{}x{}".format(rootWidth, rootHeight))

    advancedOptionsWindow = None

    # The following line of code breaks Hovertips. It just does.
    # if isAdmin: root.attributes('-topmost', True)  # keeps root window at top layer
    
    frames = []
    for i in range(10):
        frames.append(tk.Frame(root)) # bd=0, relief=tk.SOLID
        frames[i].pack(fill="x", padx=10, pady=3)

    # aesthetic/layout variables
    fontType = "None"
    fontSize = 15
    fontGeneral = (fontType, fontSize)
    fontSmall = (fontType, int(fontSize/3*2))
    listboxWidth = int(rootWidth/15) if isAdmin else int(rootWidth/10)  # listboxHeight defined above
    characterHalfRootWidth = 20 if isAdmin else 14  # HARD CODED WIDTH
    characterThirdRootWidth = 12 # ALSO HARD CODED
    tooltipHoverDelay = 0

    # data variables
    dirAbsoluteVar = tk.StringVar()
    parameterVar = tk.StringVar()
    excludedExtensionsVar = tk.StringVar(value=".one, .onepkg, .onetoc2, .onebak, .shp, .dbf")
    includeSubdirectoriesState = tk.IntVar(value=1)
    allowModifyState = tk.IntVar(value=0)
    includeHiddenFilesState = tk.IntVar(value=0)
    addRecommendationsState = tk.IntVar(value=0)
    excludedDirs = []
    currentStatusPair = (STATUS_IDLE, None)


    # widget layout
    browseButton = tk.Button(frames[0], text="Browse to Select", command=selectDirectory, font=fontGeneral, width=rootWidth)
    browseButton.pack()
    
    dirHeaderLabel = tk.Label(frames[1], text = "Directory:", font=fontGeneral)
    dirLabel = tk.Label(frames[1], textvariable=dirAbsoluteVar, font=fontSmall, anchor="e")
    dirHeaderLabel.pack(side=tk.LEFT)
    dirLabel.pack(side=tk.LEFT)
    
    excludeDirButton = tk.Button(frames[2], text="Browse to Exclude", command=queryExcludeDirectory, font=fontGeneral, width=rootWidth)
    excludeScrollbar = tk.Scrollbar(frames[2], orient=tk.HORIZONTAL)
    excludeListbox = tk.Listbox(frames[2], exportselection=0, width=rootWidth, height=0, xscrollcommand=excludeScrollbar.set)
    excludeScrollbar.config(command=excludeListbox.xview)
    excludeDirButton.pack()
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
        frames[4].pack(fill="x", padx=10, pady=(0, 3))
    else:
        findListbox.pack()
        frames[4].pack(pady=0)

    if isAdmin:
        parameterLabel = tk.Label(frames[5], text="Parameter#:", font=fontGeneral)
        argumentEntry = tk.Entry(frames[5], textvariable=parameterVar, width=rootWidth, font=fontSmall)
        parameterLabel.pack(side=tk.LEFT)
        argumentEntry.pack(side=tk.LEFT)
    else:
        frames[5].pack(pady=0)

    addRecommendationsCheckbutton = tk.Checkbutton(frames[6], text="Add Recommendations~", variable=addRecommendationsState, font=fontGeneral)
    addRecommendationsCheckbutton.pack(side=tk.LEFT)
    if isAdmin:
        allowModifyCheckbutton = tk.Checkbutton(frames[6], text="Enable Modifications", variable=allowModifyState, font=fontGeneral)
        allowModifyCheckbutton.pack(padx=(0, 0), side=tk.LEFT)  # , padx=(0, 50)
    else:
        readMeButton = tk.Button(frames[6], text="README", command=openReadMe, font=fontGeneral)
        readMeButton.pack(side=tk.LEFT)
    frames[6].pack(pady=(3, 0))

    if isAdmin:
        advancedOptionsButton = tk.Button(frames[7], text="Advanced Options", command=openAdvancedOptionsWindow, width=characterHalfRootWidth, font=fontGeneral)
        readMeButton = tk.Button(frames[7], text="README", command=openReadMe, width=characterHalfRootWidth, font=fontGeneral)
        advancedOptionsButton.pack(padx=(6, 5), side=tk.LEFT)
        readMeButton.pack(padx=(5, 6), side=tk.LEFT)
    else:
        frames[7].pack(pady=0)

    if isAdmin:
        saveSettingsButton = tk.Button(frames[8], text="Save Settings", command=saveSettingsIntoJSON, width=characterThirdRootWidth, font=fontGeneral)
        loadSettingsButton = tk.Button(frames[8], text="Load Settings", command=loadSettingsFromJSON, width=characterThirdRootWidth, font=fontGeneral)
        resultsButton = tk.Button(frames[8], text="Results", command=openResultsDirectory, width=characterThirdRootWidth, font=fontGeneral)
        saveSettingsButton.pack(padx=(17, 0), side=tk.LEFT)
        loadSettingsButton.pack(padx=(10, 0), side=tk.LEFT)
        resultsButton.pack(padx=(10, 0), side=tk.LEFT)
    else:
        frames[8].pack(pady=0)
    
    if isAdmin:
        executeButton = tk.Button(frames[9], text="Execute", command=launchController, width=rootWidth, font=fontGeneral)
        executeButton.pack()
        frames[9].pack(expand=True)
    else:
        executeButton = tk.Button(frames[9], text="Execute", command=launchController, width=characterHalfRootWidth, font=fontGeneral)
        resultsButton = tk.Button(frames[9], text="Results", command=openResultsDirectory, width=characterHalfRootWidth, font=fontGeneral)
        executeButton.pack(padx=(0, 20), side=tk.LEFT)
        resultsButton.pack(side=tk.LEFT)
    

    # tool tips
    Hovertip(browseButton, "Browse to select a directory.", hover_delay=tooltipHoverDelay) # browseTip
    Hovertip(dirHeaderLabel, "Currently selected directory.", hover_delay=tooltipHoverDelay) # dirHeaderTip
    Hovertip(excludeDirButton, "Browse to exclude subdirectories of currently selected directory.", hover_delay=tooltipHoverDelay) # excludeTip
    Hovertip(findLabel, "Run a Find procedure.", hover_delay=tooltipHoverDelay) # findTip
    Hovertip(readMeButton, "Open README file.", hover_delay=tooltipHoverDelay) # readMeTip
    Hovertip(addRecommendationsCheckbutton, "~ -> has recommendation option.\nAdd recommendations to some procedures.", hover_delay=tooltipHoverDelay) # addRecommendationsTip
    Hovertip(executeButton, "Execute program.", hover_delay=tooltipHoverDelay) # executeTip
    Hovertip(resultsButton, "Open results folder.", hover_delay=tooltipHoverDelay) # resultsTip
    if isAdmin: 
        Hovertip(fixLabel, "Run a Fix procedure.", hover_delay=tooltipHoverDelay) # fixTip
        Hovertip(parameterLabel, "# -> requires argument input.\nInput a number, string, etc. Required for some procedures.", hover_delay=tooltipHoverDelay) # parameterTip
        Hovertip(allowModifyCheckbutton, "Unless you understand the consequences of this option, leave this off.", hover_delay=tooltipHoverDelay) # allowModifyTip
        Hovertip(advancedOptionsButton, "Access advanced options.", hover_delay=tooltipHoverDelay) # advancedOptionsTip
        Hovertip(saveSettingsButton, "Save settings into a JSON file.", hover_delay=tooltipHoverDelay) # saveSettingsTip
        Hovertip(loadSettingsButton, "Load settings from a JSON file.", hover_delay=tooltipHoverDelay) # loadSettingsTip


    # bindings
    excludeListbox.bind("<Double-Button-1>", lambda _: removeExcludedDirectory()) # double left click
    excludeListbox.bind("<Button-3>", lambda _: removeExcludedDirectory()) # right click
    if isAdmin: fixListbox.bind("<<ListboxSelect>>", onSelectFixlistbox)
    root.bind('<Control-Key-w>', lambda _: root.destroy())
    root.bind('<Control-Key-W>', lambda _: root.destroy())
    root.bind("<Button-2>", lambda _: changeColorMode(root)) # middle click
    root.protocol("WM_DELETE_WINDOW", closeWindow)

    # set icon image (if available)
    logoImg = None
    if os.path.exists(LOGO_PATH):
        logoImg = tk.PhotoImage(file=LOGO_PATH)
        root.iconphoto(False, logoImg)

    # change to dark mode as a de facto default. Note: tkinter's "option_add()" is too inconsistent to use instead
    onDarkMode = False
    changeColorMode(root)


    root.mainloop()
