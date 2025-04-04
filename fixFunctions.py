from sys import maxsize as MAXSIZE
from getFileOwner import getOwnerCatch
from procedureFunctions import *


# Used by spaceFolderFixModify or searchAndReplaceFolderModify (not logs)
FOLDER_RENAMES = []


def fixfolderModifyPost(ws):
    frListLength = len(FOLDER_RENAMES)

    # in reverse order so the directoryOfFolder is never invalid; deepest folders are renamed first    
    for i in range(frListLength-1, -1, -1):
        row = i+1
        directoryOfFolder = FOLDER_RENAMES[i][0]
        oldFolderName = FOLDER_RENAMES[i][1]
        newFolderName = FOLDER_RENAMES[i][2]
        oldDirAbsolute = f"{directoryOfFolder}\\{oldFolderName}"
        # not correct directory necesssarily cause parent folders may get edited
        newDirAbsolute = f"{directoryOfFolder}\\{newFolderName}"

        for j in range(i+1, frListLength):
            otherDirectoryOfFolder = FOLDER_RENAMES[j][0]

            if otherDirectoryOfFolder.startswith(oldDirAbsolute):
                FOLDER_RENAMES[j][0] = f"{newDirAbsolute}{otherDirectoryOfFolder[len(oldDirAbsolute): ]}"
            else:
                break

        ws.write(row, wbm.ITEM_COL, oldFolderName, wbm.errorFormat)

        try:
            os.rename(addLongPathPrefix(oldDirAbsolute), addLongPathPrefix(newDirAbsolute))
            ws.write(row, wbm.OUTCOME_COL, newFolderName, wbm.modifyFormat)
            # Since row is manually tracked here, do not need to call wbm.incrementRow()
            wbm.incrementFileCount(ws)
        except PermissionError:
            ws.write(row, wbm.OUTCOME_COL, "MODIFICATION FAILED. PERMISSION ERROR.", wbm.errorFormat)
            wbm.incrementFileCount(ws)
        except Exception as e:
            ws.write(row, wbm.OUTCOME_COL, f"MODIFICATION FAILED. {e}", wbm.errorFormat)
            wbm.incrementFileCount(ws)
    
    # Now write the directories. The directoryOfFolder would be updated if a parent folder was renamed
    for i in range(frListLength):
        directoryOfFolder = FOLDER_RENAMES[i][0]
        newFolderName = FOLDER_RENAMES[i][2]
        newDirAbsolute = f"{directoryOfFolder}\\{newFolderName}"

        ws.write(i+1, wbm.DIR_COL, newDirAbsolute, wbm.dirFormat)

    FOLDER_RENAMES.clear()



def setWorkbookManager(newManager: WorkbookManager):
    # Globally declare the WorkbookManager object
    global wbm
    wbm = newManager



def spaceFolderFixHelper(oldFolderName) -> str:
    if (" " not in oldFolderName) and ("--" not in oldFolderName):
        return
    return "-".join(oldFolderName.replace("-", " ").split())


def spaceFolderFixLog(dirAbsolute, dirFolders, dirFiles, ws, arg):
    oldFolderName = dirAbsolute[dirAbsolute.rfind("\\") +1:]
    newFolderName = spaceFolderFixHelper(oldFolderName)

    if (not newFolderName): return
    
    wbm.writeDir(ws, dirAbsolute, wbm.dirFormat)
    wbm.writeItem(ws, oldFolderName, wbm.errorFormat)
    wbm.writeOutcomeAndIncrement(ws, newFolderName, wbm.logFormat)


def spaceFolderFixModify(dirAbsolute, dirFolders, dirFiles, ws, arg):
    lastSlashIndex = dirAbsolute.rfind("\\")
    oldFolderName = dirAbsolute[lastSlashIndex +1:]
    newFolderName = spaceFolderFixHelper(oldFolderName)

    if (not newFolderName): return

    directoryOfFolder = dirAbsolute[0:lastSlashIndex]
    FOLDER_RENAMES.append([directoryOfFolder, oldFolderName, newFolderName])


def spaceFileFixHelper(oldItemName) -> str:
    # Also fixes double dashes, even if no space is present
    if (" " not in oldItemName) and ("--" not in oldItemName):
        return

    lastPeriodIndex = oldItemName.rfind(".")

    # Replace '-' characters with ' ' to make the string homogenous for the upcoming split()
    # split() automatically removes leading, trailing, and excess middle whitespace
    newItemNameSansExt = "-".join(oldItemName[0:lastPeriodIndex].replace("-", " ").split())

    # This works because of a double oversight that fixes itself lol
    # TODO: Should I fix? Ya, probably.
    return newItemNameSansExt + oldItemName[lastPeriodIndex:]


def spaceFileFixLog(_1:str, _2:str, _3:str, oldItemName:str, ws, _4):
    newItemName = spaceFileFixHelper(oldItemName)
    if (not newItemName): return False

    wbm.writeItem(ws, oldItemName)
    wbm.writeOutcomeAndIncrement(ws, newItemName, wbm.logFormat)
    return True
    

def spaceFileFixModify(longFileAbsolute:str, longDirAbsolute:str, dirAbsolute:str, oldItemName:str, ws, _):
    newItemName = spaceFileFixHelper(oldItemName)
    if (not newItemName): return False

    wbm.writeItem(ws, oldItemName)

    # Log newItemName and rename file
    try:
        os.rename(longFileAbsolute, joinDirToFileName(longDirAbsolute, newItemName))
        wbm.writeOutcomeAndIncrement(ws, newItemName, wbm.modifyFormat)
    except PermissionError:
        wbm.writeOutcomeAndIncrement(ws, "MODIFICATION FAILED. PERMISSION ERROR.", wbm.errorFormat)
    except Exception as e:
        wbm.writeOutcomeAndIncrement(ws, f"MODIFICATION FAILED. {e}", wbm.errorFormat)
    return True


def deleteOldFilesStart(arg, ws):
    writeOwnerHeader(ws)
    
    global DAYS_LOWER_BOUND
    DAYS_LOWER_BOUND = arg[0]
    # Double-checking that this value is usable. Dire consequences if not.
    if (DAYS_LOWER_BOUND <= 0):
        raise Exception("DeleteOldFiles lower bound argument cannot be less than 1.")

    global DAYS_UPPER_BOUND
    if len(arg) == 2:
        DAYS_UPPER_BOUND = arg[1]
    else:
        DAYS_UPPER_BOUND = MAXSIZE


def deleteOldFilesLowerboundHelper(longFileAbsolute: str, arg) -> int:
    """Note that a file that is 23 hours and 59 minutes old is still considered 0 days old."""

    # Get date of file. This *can* error virtue of the library functions, hence try/except
    try: fileDate = datetime.fromtimestamp(os.path.getatime(longFileAbsolute))
    except: return -1

    fileDaysAgo = (TODAY - fileDate).days

    if (DAYS_LOWER_BOUND <= fileDaysAgo): return fileDaysAgo
    else: return 0


def deleteOldFilesLog(longFileAbsolute:str, longDirAbsolute:str, dirAbsolute:str, itemName:str, ws, arg):
    daysOld = deleteOldFilesLowerboundHelper(longFileAbsolute, arg)

    # Either it's actually 0 days old or the fileDate is not within the cutOffDate range. Either way, don't flag.
    # If it's greater than the upperbound, exit
    if (daysOld == 0 or daysOld >= DAYS_UPPER_BOUND):
        return False

    wbm.writeItem(ws, itemName)

    if (daysOld == -1):
        wbm.writeOutcome(ws, "UNABLE TO READ DATE.", wbm.errorFormat)
        wbm.incrementRow(ws)
        return 2
    else:
        wbm.writeAuxiliary(ws, getOwnerCatch(longFileAbsolute))
        wbm.writeOutcomeAndIncrement(ws, daysOld, wbm.logFormat)
    return True


###
def deleteOldFilesRecommendLog(longFileAbsolute:str, longDirAbsolute:str, dirAbsolute:str, itemName:str, ws, arg):
    daysOld = deleteOldFilesLowerboundHelper(longFileAbsolute, arg)

    if (daysOld == 0): return False

    if (daysOld == -1):
        wbm.writeItem(ws, itemName)
        wbm.writeOutcome(ws, "UNABLE TO READ DATE.", wbm.errorFormat)
        wbm.incrementRow(ws)
        return 2
    else:
        # Unlike the other variants, this will still log (and flag) items above the UPPER_BOUND threshold
        if daysOld >= DAYS_UPPER_BOUND:
            dynamicFormat = wbm.warningStrongFormat
        else:
            dynamicFormat = wbm.warningWeakFormat

        wbm.writeItem(ws, itemName, dynamicFormat)
        wbm.writeAuxiliary(ws, getOwnerCatch(longFileAbsolute))
        wbm.writeOutcomeAndIncrement(ws, daysOld, dynamicFormat)
    return True
###


def deleteOldFilesModify(longFileAbsolute:str, longDirAbsolute:str, dirAbsolute:str, itemName:str, ws, arg):
    daysOld = deleteOldFilesLowerboundHelper(longFileAbsolute, arg)

    # Either it's actually 0 days old or the fileDate is not within the cutOffDate range. Either way, don't flag.
    # If it's greater than the upperbound, exit
    if (daysOld == 0 or daysOld >= DAYS_UPPER_BOUND):
        return False

    wbm.writeItem(ws, itemName)

    if (daysOld == -1):
        wbm.writeOutcome(ws, "UNABLE TO READ DATE.", wbm.errorFormat) 
        wbm.incrementRow(ws)
        return 2
    else:
        try:
            wbm.writeAuxiliary(ws, getOwnerCatch(longFileAbsolute))
            os.remove(longFileAbsolute)
            wbm.writeOutcomeAndIncrement(ws, daysOld, wbm.modifyFormat)
        except PermissionError:
            wbm.writeOutcomeAndIncrement(ws, "FAILED TO DELETE. PERMISSION ERROR.", wbm.errorFormat)
        except Exception as e:
            wbm.writeOutcomeAndIncrement(ws, f"FAILED TO DELETE. {e}", wbm.errorFormat)
    return True


def deleteEmptyDirectoriesLog(_, dirFolders, dirFiles, ws, arg):
    tooFewAmount = arg[0]

    # If even 1 folder exists, this isn't empty
    if len(dirFolders) != 0: return False

    # If equal to tooFewAmount or less, then this folder needs to be at least flagged
    fileAmount = len(dirFiles)
    if fileAmount <= tooFewAmount:
        wbm.writeOutcomeAndIncrement(ws, fileAmount, wbm.logFormat)
        return True
    return False


###
def deleteEmptyDirectoriesRecommendLog(dirAbsolute, dirFolders, dirFiles, ws, arg):
    tooFewAmount = arg[0]

    if len(dirFolders) != 0: return False

    fileAmount = len(dirFiles)
    if fileAmount <= tooFewAmount:
        
        ## HARD CODED RECOMMENDATIONS
        # Dynamic Format
        if fileAmount <= 1:
            dynamicFormat = wbm.warningStrongFormat
        elif fileAmount == 2:
            dynamicFormat = wbm.warningWeakFormat
        else:
            dynamicFormat = wbm.logFormat

        # wbm.writeItem(ws, "", dynamicFormat)
        wbm.writeOutcomeAndIncrement(ws, fileAmount, dynamicFormat)
        return True
    return False
###
    

def deleteEmptyDirectoriesModify(dirAbsolute, dirFolders, dirFiles, ws, arg):
    tooFewAmount = arg[0]

    if len(dirFolders) != 0: return False

    fileAmount = len(dirFiles)
    if fileAmount <= tooFewAmount:
        # If it specifically has 0 files, delete the folder
        if (fileAmount == 0):
            try: 
                os.rmdir(addLongPathPrefix(dirAbsolute))
                wbm.writeOutcomeAndIncrement(ws, fileAmount, wbm.modifyFormat)
            except Exception as e:
                wbm.writeOutcomeAndIncrement(ws, f"0 FILES. UNABLE TO DELETE. {e}", wbm.errorFormat)
            return True
        # Otherwise, just flag as usual
        else:
            wbm.writeOutcomeAndIncrement(ws, fileAmount, wbm.logFormat)
            return True
    return False


def searchAndReplaceFolderHelper(oldFolderName:str, arg):
    newFolderName = oldFolderName
    for argPair in arg:
        toBeReplaced, replacer = argPair
        newFolderName = newFolderName.replace(toBeReplaced, replacer)

    if (oldFolderName == newFolderName): return
    return newFolderName


def searchAndReplaceFolderLog(dirAbsolute, dirFolders, dirFiles, ws, arg):
    oldFolderName = dirAbsolute[dirAbsolute.rfind("\\") +1:]

    if not (newFolderName := searchAndReplaceFolderHelper(oldFolderName, arg)):
        return
    
    wbm.writeDir(ws, dirAbsolute, wbm.dirFormat)
    wbm.writeItem(ws, oldFolderName, wbm.errorFormat)
    wbm.writeOutcomeAndIncrement(ws, newFolderName, wbm.logFormat)


def searchAndReplaceFolderModify(dirAbsolute, dirFolders, dirFiles, ws, arg):
    lastSlashIndex = dirAbsolute.rfind("\\")
    oldFolderName = dirAbsolute[lastSlashIndex +1:]

    if not (newFolderName := searchAndReplaceFolderHelper(oldFolderName, arg)): return
    
    directoryOfFolder = dirAbsolute[0:lastSlashIndex]
    FOLDER_RENAMES.append([directoryOfFolder, oldFolderName, newFolderName])


def searchAndReplaceFileHelper(oldItemName:str, arg):
    lastPeriodIndex = oldItemName.rfind(".")
    if lastPeriodIndex == -1:
        extension = ""
        oldItemNameSansExt = oldItemName[0:]
    else:
        extension = oldItemName[lastPeriodIndex:]
        oldItemNameSansExt = oldItemName[0:lastPeriodIndex]
    
    # Order of argument pairs given matters.
    newItemNameSansExt = oldItemNameSansExt
    for argPair in arg:
        toBeReplaced, replacer = argPair
        newItemNameSansExt = newItemNameSansExt.replace(toBeReplaced, replacer)

    if (oldItemNameSansExt == newItemNameSansExt): return
    return newItemNameSansExt + extension


def searchAndReplaceFileLog(longFileAbsolute:str, longDirAbsolute:str, _:str, oldItemName:str, ws, arg):
    if not (newItemName := searchAndReplaceFileHelper(oldItemName, arg)): return False

    wbm.writeItem(ws, oldItemName, wbm.errorFormat)
    wbm.writeOutcomeAndIncrement(ws, newItemName, wbm.logFormat)
    return True


def searchAndReplaceFileModify(longFileAbsolute:str, longDirAbsolute:str, dirAbsolute:str, oldItemName:str, ws, arg):
    if not (newItemName := searchAndReplaceFileHelper(oldItemName, arg)): return False

    wbm.writeItem(ws, oldItemName, wbm.errorFormat)

    try:
        os.rename(longFileAbsolute, joinDirToFileName(longDirAbsolute, newItemName))
        wbm.writeOutcomeAndIncrement(ws, newItemName, wbm.modifyFormat)
    except PermissionError:
        wbm.writeOutcomeAndIncrement(ws, "MODIFICATION FAILED. PERMISSION ERROR.", wbm.errorFormat)
    except Exception as e:
        wbm.writeOutcomeAndIncrement(ws, f"MODIFICATION FAILED. {e}", wbm.errorFormat)
    return True


def deleteEmptyFilesStart(_, ws):
    writeOwnerHeader(ws)


def deleteEmptyFilesLog(longFileAbsolute:str, longDirAbsolute:str, dirAbsolute:str, itemName:str, ws, _):
    try:
        fileSize = os.path.getsize(longFileAbsolute) # Bytes
    except PermissionError:
        wbm.writeItem(ws, itemName)
        wbm.writeOutcome(ws, "UNABLE TO READ. PERMISSION ERROR.", wbm.errorFormat)
        wbm.incrementRow(ws)
        return 2
    except Exception as e:
        wbm.writeItem(ws, itemName)
        wbm.writeOutcome(ws, f"UNABLE TO READ. {e}", wbm.errorFormat)
        wbm.incrementRow(ws)
        return 2

    if fileSize == 0:
        wbm.writeItem(ws, itemName)
        wbm.writeAuxiliary(ws, getOwnerCatch(longFileAbsolute))
        wbm.writeOutcomeAndIncrement(ws, "", wbm.logFormat)
        return True
    return False


###
def deleteEmptyFilesRecommendLog(longFileAbsolute:str, longDirAbsolute:str, dirAbsolute:str, itemName:str, ws, _):
    try:
        fileSize = os.path.getsize(longFileAbsolute)
    except PermissionError:
        wbm.writeItem(ws, itemName)
        wbm.writeOutcome(ws, "UNABLE TO READ. PERMISSION ERROR.", wbm.errorFormat)
        wbm.incrementRow(ws)
        return 2
    except Exception as e:
        wbm.writeItem(ws, itemName)
        wbm.writeOutcome(ws, f"UNABLE TO READ. {e}", wbm.errorFormat)
        wbm.incrementRow(ws)
        return 2

    if fileSize == 0:
        wbm.writeItem(ws, itemName, wbm.warningStrongFormat)
        wbm.writeAuxiliary(ws, getOwnerCatch(longFileAbsolute))
        wbm.writeOutcomeAndIncrement(ws, "", wbm.warningStrongFormat)
        return True
    return False
###


def deleteEmptyFilesModify(longFileAbsolute:str, longDirAbsolute:str, dirAbsolute:str, itemName:str, ws, _):
    """Glitch exists in that the current excel file will be considered empty.
    However, despite claiming so, the program does not actually delete it.'"""

    try:
        fileSize = os.path.getsize(longFileAbsolute)
    except PermissionError:
        wbm.writeItem(ws, itemName)
        wbm.writeOutcome(ws, "UNABLE TO READ. PERMISSION ERROR.", wbm.errorFormat)
        wbm.incrementRow(ws)
        return 2
    except Exception as e:
        wbm.writeItem(ws, itemName)
        wbm.writeOutcome(ws, f"UNABLE TO READ. {e}", wbm.errorFormat)
        wbm.incrementRow(ws)
        return 2

    # Stage for deletion
    if fileSize == 0:
        wbm.writeItem(ws, itemName)

        try:
            wbm.writeAuxiliary(ws, getOwnerCatch(longFileAbsolute))
            os.remove(longFileAbsolute)
            wbm.writeOutcomeAndIncrement(ws, "", wbm.modifyFormat)
        except PermissionError:
            wbm.writeOutcomeAndIncrement(ws, "FAILED TO DELETE. PERMISSION ERROR.", wbm.errorFormat)
        except Exception as e:
            wbm.writeOutcomeAndIncrement(ws, f"FAILED TO DELETE. {e}", wbm.errorFormat)
        return True
    return False
