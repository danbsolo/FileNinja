from workbookManager import WorkbookManager
from datetime import datetime
import procedureHelpers
import procedureFunctions
import traceback
import tkinter as tk
from defs import *



def launchController(dirAbsolute:str, includeSubfolders:bool, allowModify:bool, includeHiddenFiles:bool, addRecommendations:bool, selectedFindProcedures:list[str], selectedFixProcedures:list[str], argUnprocessed:str, excludedDirs:set[str]):
    if (not dirAbsolute): return -2

    # If no procedures are selected, exit
    if len(selectedFindProcedures + selectedFixProcedures) == 0:
        return -8

    # If multiple fix procedures are selected and allowModify is checked, exit
    if allowModify and len(selectedFixProcedures) > 1:
        return -5
    
    # If allowModify and addRecommendations are both turned on, even if it would be of no consequence, exit
    # If allowModify and includeHiddenFiles are both turned on, exit
    if allowModify and (addRecommendations or includeHiddenFiles):
        return -7
    
    # Double check that the user wants to allow modifications
    if allowModify and not tk.messagebox.askyesnocancel("Allow Modify?", "You have chosen to modify items. This is an IRREVERSIBLE action. Are you sure?"):
        return 0
    
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
    procedureHelpers.setWorkbookManager(wbm)
    procedureFunctions.setWorkbookManager(wbm)


    # If argUnprocessed is just an empty string or *just* whitespace
    if not argUnprocessed.strip():
        argsList = []
    else:
        argsList = argUnprocessed.split("/")
    
    argsListLength = len(argsList)
    currentArg = 0


    # If arguments were inputted, each procedure needs to be accounted for via "/"s
    if argsList:
        for procedureName in selectedFindProcedures + selectedFixProcedures:
            arg = None

            if ALL_PROCEDURES[procedureName].validatorFunction:
                if currentArg >= argsListLength:
                    return -6

                arg = argsList[currentArg]
                currentArg += 1
            
            if (not wbm.addProcedure(ALL_PROCEDURES[procedureName], allowModify, addRecommendations, arg)):
                return -3
    else:
        for procedureName in selectedFindProcedures + selectedFixProcedures:

            # This procs if a procedure doesn't *have* a default value
            # For find procedures, since they're all available in Lite, this really shouldn't ever proc
            if (not wbm.addProcedure(ALL_PROCEDURES[procedureName], allowModify, addRecommendations, None)):
                return -3

    try:
        wbm.initiateCrawl(dirAbsolute, includeSubfolders, allowModify, includeHiddenFiles, addRecommendations, excludedDirs)
        wbm.close()
        os.startfile(workbookPathName)
    except Exception as e:
        return (-999, f"TAKE A SCREENSHOT — a fatal error has occurred.\n\n{traceback.format_exc()}")
    
    return 0
