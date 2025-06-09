from workbookManager import WorkbookManager
from datetime import datetime
import procedureHelpers
import procedureFunctions
import traceback
from defs import *
import common


def launchController(dirAbsolute:str, includeSubdirectories:bool, allowModify:bool, includeHiddenFiles:bool, addRecommendations:bool, selectedFindProcedures:list[str], selectedFixProcedures:list[str], argUnprocessed:str, excludedDirs:set[str]):
    # If dirAbsolute value was not selected (empty string)
    if (not dirAbsolute): return (-2, None)

    # If dirAbsolute path is non-existent
    if (not os.path.exists(dirAbsolute)): return (-9, None)

    # If no procedures are selected, exit
    if len(selectedFindProcedures + selectedFixProcedures) == 0:
        return (-8, None)

    # If multiple fix procedures are selected and allowModify is checked, exit
    if allowModify and len(selectedFixProcedures) > 1:
        return (-5, None)
    
    # If allowModify and addRecommendations are both turned on, even if it would be of no consequence, exit
    # If allowModify and includeHiddenFiles are both turned on, exit
    if allowModify and (addRecommendations or includeHiddenFiles):
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
                    return (-6, None)

                arg = argsList[currentArg]
                currentArg += 1
            
            if (not wbm.addProcedure(ALL_PROCEDURES[procedureName], allowModify, addRecommendations, arg)):
                return (-3, None)
    else:
        for procedureName in selectedFindProcedures + selectedFixProcedures:

            # This procs if a procedure doesn't *have* a default value
            # For find procedures, since they're all available in Lite, this really shouldn't ever proc
            if (not wbm.addProcedure(ALL_PROCEDURES[procedureName], allowModify, addRecommendations, None)):
                return (-3, None)

    try:
        wbm.initiateCrawl(dirAbsolute, includeSubdirectories, allowModify, includeHiddenFiles, addRecommendations, excludedDirs)
        wbm.close()
        os.startfile(workbookPathName)
    except Exception as e:
        return (STATUS_UNEXPECTED, f"{traceback.format_exc()}")
    
    return (STATUS_SUCCESSFUL, None)


def launchControllerFromSettings(settings):
    # settings[ALLOW_MODIFY_KEY]
    return launchController(
        settings[DIR_ABSOLUTE_KEY],
        settings[INCLUDE_SUBDIRECTORIES_KEY],
        False,
        settings[INCLUDE_HIDDEN_FILES_KEY],
        settings[ADD_RECOMMENDATIONS_KEY],
        settings[SELECTED_FIND_PROCEDURES_KEY],
        settings[SELECTED_FIX_PROCEDURES_KEY],
        settings[ARG_UNPROCESSED_KEY],
        settings[EXCLUDED_DIRS_KEY]
    )


def launchControllerFromJSON(jsonFilePath):
    settingsPair = common.loadSettingsFromJSON(jsonFilePath)
    status = settingsPair[0]
    result = settingsPair[1]

    if status == STATUS_SUCCESSFUL:
        return launchControllerFromSettings(result)
    else:
        return settingsPair
