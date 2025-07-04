from FileNinjaSuite.FileNinja.workbookManager import WorkbookManager
from datetime import datetime
from FileNinjaSuite.FileNinja import procedureHelpers
from FileNinjaSuite.FileNinja import procedureFunctions
from FileNinjaSuite.FileNinja.defs import *
from FileNinjaSuite.FileNinja import common
import traceback
import re


def launchController(dirAbsolute:str, includeSubdirectories:bool, enableModifications:bool, includeHiddenFiles:bool, addRecommendations:bool, selectedFindProcedures:list[str], selectedFixProcedures:list[str], argUnprocessed:str, excludedDirs:set[str], excludedExtensionsUnprocessed:str):
    # If dirAbsolute value was not selected (empty string)
    if (not dirAbsolute): return (-2, None)

    # If dirAbsolute path is non-existent
    if (not os.path.exists(dirAbsolute)): return (-9, None)

    # If no procedures are selected, exit
    if len(selectedFindProcedures + selectedFixProcedures) == 0:
        return (-8, None)

    # If multiple fix procedures are selected and enableModifications is checked, exit
    if enableModifications and len(selectedFixProcedures) > 1:
        return (-5, None)
    
    # If enableModifications and addRecommendations are both turned on, even if it would be of no consequence, exit
    # If enableModifications and includeHiddenFiles are both turned on, exit
    if enableModifications and (addRecommendations or includeHiddenFiles):
        return (-7, None)
    
    # make it all backslashes, not forward slashes. This is to make it homogenous with os.walk() output
    dirAbsolute = dirAbsolute.replace("/", "\\")

    #
    if includeSubdirectories:
        # If excludedDirs are specified, run checks accordingly
        if excludedDirs:
            for i in range(len(excludedDirs)):
                if (not os.path.exists(excludedDirs[i])):
                    return (-4, None)
                
                excludedDirs[i] = excludedDirs[i].replace("/", "\\")

            for exDir in excludedDirs:
                # Check dirAbsolute is not an exDir. Check that exDir is a child to dirAbsolute
                if (exDir == dirAbsolute) or not (exDir.startswith(dirAbsolute)):
                    return (-4, None)

                # Check that no exDir is a subfolder to another exDir
                for exDir2 in excludedDirs:
                    if (exDir != exDir2) and (exDir.startswith(exDir2)):
                        return (-4, None)
    else:
        # If subdirectories are not included but excluded dirs are specified, clear it
        if excludedDirs:
            excludedDirs.clear()

    # Check excluded extensions
    excludedExtensions = None
    if excludedExtensionsUnprocessed:
        # Remove trailing and leading whitespace
        # Then, remove trailing and leading commas so the noExtension ("") isn't added accidentally
        # Split by comma
        excludedExtensions = excludedExtensionsUnprocessed.strip().strip(",").split(",")
        for i in range(len(excludedExtensions)):
            excludedExtensions[i] = excludedExtensions[i].strip().lower()
            currentExt = excludedExtensions[i]
            if not bool(re.fullmatch(r"\.[a-zA-Z0-9]+", currentExt)) and currentExt != "":
                print(f"|{currentExt}| is invalid")
                return (-13, None)
            
        excludedExtensions = set(excludedExtensions)

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

    # preCrawlInitiation
    wbm.createMetaSheets(addRecommendations)

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
                    return (-6, None)

                arg = argsList[currentArg]
                currentArg += 1
            
            if (not wbm.addProcedure(ALL_PROCEDURES[procedureName], enableModifications, addRecommendations, arg)):
                return (-3, None)
    else:
        for procedureName in selectedFindProcedures + selectedFixProcedures:

            # This procs if a procedure doesn't *have* a default value
            # For find procedures, since they're all available in Lite, this really shouldn't ever proc
            if (not wbm.addProcedure(ALL_PROCEDURES[procedureName], enableModifications, addRecommendations, None)):
                return (-3, None)

    try:
        wbm.initiateCrawl(dirAbsolute, includeSubdirectories, enableModifications, includeHiddenFiles, addRecommendations, excludedDirs, excludedExtensions)
        wbm.close()
        os.startfile(workbookPathName)
    except:
        return (STATUS_UNEXPECTED, f"{traceback.format_exc()}")
    
    return (STATUS_SUCCESSFUL, None)


def launchControllerFromSettings(settings):
    # "Enable Modifications" is always set to False
    return launchController(
        settings[DIR_ABSOLUTE_KEY],
        settings[INCLUDE_SUBDIRECTORIES_KEY],
        False,
        settings[INCLUDE_HIDDEN_FILES_KEY],
        settings[ADD_RECOMMENDATIONS_KEY],
        settings[SELECTED_FIND_PROCEDURES_KEY],
        settings[SELECTED_FIX_PROCEDURES_KEY],
        settings[ARG_UNPROCESSED_KEY],
        settings[EXCLUDED_DIRS_KEY],
        settings[EXCLUDED_EXTENSIONS_KEY]
    )


def launchControllerFromJSON(jsonFilePath):
    settingsPair = common.loadSettingsFromJSON(jsonFilePath)
    status = settingsPair[0]
    result = settingsPair[1]

    if status == STATUS_SUCCESSFUL:
        return launchControllerFromSettings(result)
    else:
        return settingsPair
