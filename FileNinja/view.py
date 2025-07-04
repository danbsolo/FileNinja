import tkinter as tk
from tkinter import filedialog
import os
from FileNinjaSuite.FileNinja.defs import *
import sys
from FileNinjaSuite.FileNinja import filesScannedSharedVar
from FileNinjaSuite.FileNinja import control
import json
from FileNinjaSuite.FileNinja import common
from FileNinjaSuite.Shared import guiController, sharedViewHelpers
#from FileNinjaSuite.FileChop import fileChop


def launchView(isAdmin: bool):
    def threadAliveFunction():
        executeButton.config(text=f"{filesScannedSharedVar.FILES_SCANNED} files")

    def threadDoneFunction():
        filesScannedSharedVar.FILES_SCANNED = 0

    def confirmEnableModificationsAndInitiateControllerThread():
        # Double check that the user wants to enable modifications
        if enableModificationsState.get() and not tk.messagebox.askyesnocancel(
            "Confirm Enabling of Modifications", 
            "Modifications are irreversible. By proceeding with the Fix procedure(s), you are confirming the action has been requested by the data owner and that there are no identified Litigation Holds (LIT-HOLD) or Access to Information and Privacy (ATIP) requests for in scope data. Proceed?"):
            return
        
        guiC.setThreadVars(executeButton, threadAliveFunction, threadDoneFunction)
        guiC.initiateControllerThread(lambda:
                                        control.launchController(dirAbsoluteVar.get(), bool(includeSubdirectoriesState.get()), bool(enableModificationsState.get()), bool(includeHiddenFilesState.get()), bool(addRecommendationsState.get()),
                                        [findListbox.get(fm) for fm in findListbox.curselection()],
                                        [fixListbox.get(fm) for fm in fixListbox.curselection()] if isAdmin else [],
                                        parameterVar.get(),
                                        excludedDirs,
                                        excludedExtensionsVar.get()))


    def selectDirectory():
        potentialDirectory = filedialog.askdirectory(title="Browse to SELECT",
                                                     initialdir=dirAbsoluteVar.get(),
                                                     mustexist=True)

        if (potentialDirectory):
            dirAbsoluteVar.set(potentialDirectory)


    def selectExcludeDirectory():
        potentialExcludedDirectory = filedialog.askdirectory(title="Browse to EXCLUDE",
                                                             initialdir=dirAbsoluteVar.get(),
                                                             mustexist=True)

        if (not potentialExcludedDirectory) or (potentialExcludedDirectory in excludedDirs):
            return
        
        commenceExcludeDirectory(potentialExcludedDirectory)

    def commenceExcludeDirectory(directoryToBeExcluded):
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


    def saveSettingsIntoJSON():
        settings = {
            DIR_ABSOLUTE_KEY: dirAbsoluteVar.get(),
            EXCLUDED_DIRS_KEY: excludedDirs,
            EXCLUDED_EXTENSIONS_KEY: excludedExtensionsVar.get(),
            SELECTED_FIND_PROCEDURES_KEY: [findListbox.get(fm) for fm in findListbox.curselection()],
            SELECTED_FIX_PROCEDURES_KEY: [fixListbox.get(fm) for fm in fixListbox.curselection()] if isAdmin else [],
            ARG_UNPROCESSED_KEY: parameterVar.get(),
            INCLUDE_SUBDIRECTORIES_KEY: bool(includeSubdirectoriesState.get()),
            ENABLE_MODIFICATIONS_KEY: bool(enableModificationsState.get()),
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
        enableModificationsState.set(settings[ENABLE_MODIFICATIONS_KEY])
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
            commenceExcludeDirectory(item)


    def openAdvancedOptionsWindow():
        nonlocal advancedOptionsWindow

        # If this window is already open, lift it to the forefront
        if advancedOptionsWindow and advancedOptionsWindow.winfo_exists():
            advancedOptionsWindow.deiconify()
            advancedOptionsWindow.lift()
            return

        advancedOptionsWindow = tk.Toplevel(root)
        advancedOptionsWindow.title(f"{TITLE} Advanced Options")
        advancedOptionsWindow.resizable(0, 0)
        # advancedOptionsWindow.geometry("500x200")
        advancedOptionsWindow.geometry("+{}+{}".format(root.winfo_pointerx()+rootWidth//2, root.winfo_pointery()))

        if logoImg := guiC.getLogoIcon():
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
        excludedExtensionsEntry.icursor(tk.END)
        excludedExtensionsEntry.xview_moveto(1.0)
        excludedExtensionsLabel.pack(side=tk.LEFT)
        excludedExtensionsEntry.pack(side=tk.LEFT)
        
        guiC.createHoverTips({
            includeSubdirectoriesCheckbutton: "Dive into all subdirectories, other than those excluded.",
            includeHiddenFilesCheckbutton: "Include hidden files in Find procedure output. Fix procedures always ignore hidden files.",
            excludedExtensionsLabel: "Comma separated list of extensions to exclude from scan."
            }
        )

        if guiC.isOnDarkMode:
            sharedViewHelpers.changeToDarkMode(advancedOptionsWindow)
        else:
            sharedViewHelpers.changeToLightMode(advancedOptionsWindow)



    listboxHeight = max(len(FIND_PROCEDURES_DISPLAY), len(FIX_PROCEDURES_DISPLAY)) + 1
    listboxHeightMultiplier = 17

    # root window stuff
    root = tk.Tk()
    root.title(TITLE)
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
    fontType = None
    fontSize = 15
    fontGeneral = (fontType, fontSize)
    fontSmall = (fontType, int(fontSize/3*2))
    listboxWidth = int(rootWidth/15) if isAdmin else int(rootWidth/10)  # listboxHeight defined above
    characterHalfRootWidth = 20 if isAdmin else 14  # HARD CODED WIDTH
    characterThirdRootWidth = 12 # ALSO HARD CODED

    # data variables
    dirAbsoluteVar = tk.StringVar()
    parameterVar = tk.StringVar()
    excludedExtensionsVar = tk.StringVar(
        value=".shp, .dbf, .shx, .sbd, .sbx, .spx, .sbn, .qpj, .atx, .cpg, .prj, .gbd, .gdbtablx, .gdbtable, .freelist, .horizon, .gdbindexes, .one, .onepkg, .onetoc2, .onebak,")
    includeSubdirectoriesState = tk.IntVar(value=1)
    enableModificationsState = tk.IntVar(value=0)
    includeHiddenFilesState = tk.IntVar(value=0)
    addRecommendationsState = tk.IntVar(value=0)
    excludedDirs = []

    # widget layout
    browseButton = tk.Button(frames[0], text="Browse to Select", command=selectDirectory, font=fontGeneral, width=rootWidth)
    browseButton.pack()
    
    dirHeaderLabel = tk.Label(frames[1], text = "Directory:", font=fontGeneral)
    dirLabel = tk.Label(frames[1], textvariable=dirAbsoluteVar, font=fontSmall, anchor="e")
    dirHeaderLabel.pack(side=tk.LEFT)
    dirLabel.pack(side=tk.LEFT)
    
    excludeDirButton = tk.Button(frames[2], text="Browse to Exclude", command=selectExcludeDirectory, font=fontGeneral, width=rootWidth)
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
        enableModificationsCheckbutton = tk.Checkbutton(frames[6], text="Enable Modifications", variable=enableModificationsState, font=fontGeneral)
        enableModificationsCheckbutton.pack(padx=(0, 0), side=tk.LEFT)  # , padx=(0, 50)
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
        executeButton = tk.Button(frames[9], text="Execute", command=confirmEnableModificationsAndInitiateControllerThread, width=rootWidth, font=fontGeneral)
        executeButton.pack()
        frames[9].pack(expand=True)
    else:
        executeButton = tk.Button(frames[9], text="Execute", command=confirmEnableModificationsAndInitiateControllerThread, width=characterHalfRootWidth, font=fontGeneral)
        resultsButton = tk.Button(frames[9], text="Results", command=openResultsDirectory, width=characterHalfRootWidth, font=fontGeneral)
        executeButton.pack(padx=(0, 20), side=tk.LEFT)
        resultsButton.pack(side=tk.LEFT)
    
    # tool tips
    hoverTipDictionary = {
        browseButton: "Browse to select a directory.",
        dirHeaderLabel: "Currently selected directory.",
        excludeDirButton: "Browse to exclude subdirectories of currently selected directory.",
        findLabel: "Run a Find procedure.",
        readMeButton: "Open README file.",
        addRecommendationsCheckbutton: "~ -> has recommendation option.\nAdd recommendations to some procedures.",
        executeButton: "Execute program.",
        resultsButton: "Open results folder."
    }
    if isAdmin:
        hoverTipDictionary.update({
            fixLabel: "Run a Fix procedure.",
            parameterLabel: "# -> requires argument input.\nInput a number, string, etc. Required for some procedures.",
            enableModificationsCheckbutton: "Unless you understand the consequences of this option, leave this off.",
            advancedOptionsButton: "Access advanced options.",
            saveSettingsButton: "Save settings into a JSON file.",
            loadSettingsButton: "Load settings from a JSON file.",
        }) 

    # bindings
    excludeListbox.bind("<Double-Button-1>", lambda _: removeExcludedDirectory()) # double left click
    excludeListbox.bind("<Button-3>", lambda _: removeExcludedDirectory()) # right click
    if isAdmin: fixListbox.bind("<<ListboxSelect>>", onSelectFixlistbox)

    # guiController stuff
    guiC = guiController.GUIController(root, TITLE)
    guiC.standardInitialize()
    guiC.createHoverTips(hoverTipDictionary)


    root.mainloop()
