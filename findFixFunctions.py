from typing import Set
import string
import os
import time
from datetime import datetime, timedelta
from workbookManager import WorkbookManager


# Used by badCharFind().
# Includes ' ' (space) as there is a separate method for finding that error
PERMISSIBLE_CHARACTERS = set(string.ascii_letters + string.digits + "- ")
CHARACTER_LIMIT = 200

# Used by deleteOldFilesGeneral()
TODAY = datetime.now()

# Used by fileTypeMisc() and fileTypePost()
FILE_EXTENSION_COUNT = {}
FILE_EXTENSION_TOTAL_SIZE = {}

# Used by duplicateFileMisc() and duplicateFilePost()
FILES_AND_PATHS = {}



def setWorkbookManager(newManager: WorkbookManager):
    # Globally declare the WorkboookManager object
    global wbm
    wbm = newManager
    

def listAll(_:str, itemName:str, ws):
    wbm.writeInCell(ws, wbm.ITEM_COL, itemName, rowIncrement=1)
    return False

def spaceFind(_:str, itemName:str, ws) -> bool:
    if " " in itemName: 
        wbm.writeInCell(ws, wbm.ITEM_COL, itemName, wbm.errorFormat, 1, 1)
        return True
    return False
    

def overCharLimitFind(dirAbsolute:str, itemName:str, ws) -> bool:
    absoluteItemLength = len(dirAbsolute + "/" + itemName)
    if (absoluteItemLength > CHARACTER_LIMIT):
        wbm.writeInCell(ws, wbm.ITEM_COL, itemName, wbm.errorFormat)
        wbm.writeInCell(ws, wbm.ERROR_COL, "{} > {}".format(absoluteItemLength, CHARACTER_LIMIT), rowIncrement=1, fileIncrement=1)
        return True
    return False

def badCharFind(_:str, itemName:str, ws) -> Set[str]:
    """Does not check for space characters."""
    
    badChars = set()

    # If no extension (aka, no period), lastPeriodIndex will equal -1
    lastPeriodIndex = itemName.rfind(".")

    if (lastPeriodIndex == -1): itemNameLength = len(itemName)
    else: itemNameLength = lastPeriodIndex

    for i in range(itemNameLength):
        if itemName[i] not in PERMISSIBLE_CHARACTERS:
            badChars.add(itemName[i])
        
        # double-dash error
        elif itemName[i:i+2] == "--":
            badChars.add(itemName[i])

    # if any bad characters were found
    if (badChars):
        wbm.writeInCell(ws, wbm.ITEM_COL, itemName, wbm.errorFormat)
        wbm.writeInCell(ws, wbm.ERROR_COL, "".join(badChars), rowIncrement=1, fileIncrement=1)
        return True
    return False


def spaceFixHelper(oldItemName) -> str:
    if (" " not in oldItemName):
        return

    lastPeriodIndex = oldItemName.rfind(".")

    # Replace '-' characters with ' ' to make the string homogenous for the upcoming split()
    # split() automatically removes leading, trailing, and excess middle whitespace
    newItemNameSansExt = "-".join(oldItemName[0:lastPeriodIndex].replace("-", " ").split())

    # This works because of a double oversight that fixes itself lol
    # TODO: Should I fix? Ya, probably.
    return newItemNameSansExt + oldItemName[lastPeriodIndex:]


def spaceFixLog(_:str, oldItemName:str, ws):
    newItemName = spaceFixHelper(oldItemName)
    if (not newItemName): return

    wbm.writeInCell(ws, wbm.ITEM_COL, oldItemName)
    wbm.writeInCell(ws, wbm.MOD_COL, newItemName, wbm.logFormat, 1, 1)

    
def spaceFixModify(dirAbsolute:str, oldItemName:str, ws):
    newItemName = spaceFixHelper(oldItemName)
    if (not newItemName): return

    wbm.writeInCell(ws, wbm.ITEM_COL, oldItemName)

    # Log newItemName and rename file
    try:
        os.rename(dirAbsolute + "/" + oldItemName, dirAbsolute + "/" + newItemName)
        wbm.writeInCell(ws, wbm.MOD_COL, newItemName, wbm.modifyFormat, 1, 1)
    except PermissionError:
        wbm.writeInCell(ws, wbm.MOD_COL, "FILE LOCKED. MODIFICATION FAILED.", wbm.errorFormat, 1, 0)
    except OSError:
        wbm.writeInCell(ws, wbm.MOD_COL, "OS ERROR. MODIFICATION FAILED.", wbm.errorFormat, 1, 0)


def deleteOldFilesHelper(fullFilePath: str) -> int:
    """Note that a file that is 23 hours and 59 minutes old is still considered 0 days old."""
    
    daysTooOld = wbm.fixArg
    # Could double-check that this value is usable each time. Dire consequences if not.
    #  if (daysTooOld <= 0): return -1
    
    # gets date of file. This *can* error virtue of the library functions, hence try/except
    try: fileDate = datetime.fromtimestamp(os.path.getatime(fullFilePath))
    except: return -1

    fileDaysAgo = (TODAY - fileDate).days

    if (fileDaysAgo >= daysTooOld): return fileDaysAgo
    else: return 0


def deleteOldFilesLog(dirAbsolute:str, itemName:str, ws):
    fullFilePath =  dirAbsolute + "\\" + itemName
    daysOld = deleteOldFilesHelper(fullFilePath)

    # Either it's actually 0 days old or the fileDate is after the cutOffDate. Either way, don't flag.         
    if (daysOld == 0): return

    wbm.writeInCell(ws, wbm.ITEM_COL, itemName)

    if (daysOld == -1):
        wbm.writeInCell(ws, wbm.MOD_COL, "UNABLE TO READ DATE", wbm.errorFormat, 1, 1) 
    elif len(fullFilePath) > CHARACTER_LIMIT:
        wbm.writeInCell(ws, wbm.MOD_COL, "{} days but violates charLimit".format(daysOld), wbm.logFormat, 1, 1)
    else:
        wbm.writeInCell(ws, wbm.MOD_COL, "{}".format(daysOld), wbm.logFormat, 1, 1)    


def deleteOldFilesModify(dirAbsolute:str, itemName:str, ws):
    fullFilePath =  dirAbsolute + "\\" + itemName
    daysOld = deleteOldFilesHelper(fullFilePath)

    if (daysOld == 0): return

    wbm.writeInCell(ws, wbm.ITEM_COL, itemName)

    if (daysOld == -1):
        wbm.writeInCell(ws, wbm.MOD_COL, "UNABLE TO READ DATE", wbm.errorFormat, 1, 1) 
    # If over CHARACTER_LIMIT characters, do not delete as it is not backed up
    elif len(fullFilePath) > CHARACTER_LIMIT:
        wbm.writeInCell(ws, wbm.MOD_COL, "{} days but violates charLimit".format(daysOld), wbm.logFormat, 1, 1)
    else:
        try:
            os.remove(fullFilePath)
            wbm.writeInCell(ws, wbm.MOD_COL, "{}".format(daysOld), wbm.modifyFormat, 1, 1)
        except:
            wbm.writeInCell(ws, wbm.MOD_COL, "FAILED TO DELETE", wbm.errorFormat, 1, 1)
            
    

def fileExtensionConcurrent(dirAbsolute:str, itemName:str, _):
    try: fileSize = os.path.getsize(dirAbsolute+"\\"+itemName) / 1000_000  # Bytes / 1000_000 = MBs
    except: return False

    lastPeriodIndex = itemName.rfind(".")

    if lastPeriodIndex == -1: extension = ""
    else:
        extension = itemName[lastPeriodIndex:]
        extension = extension.lower()

    if extension in FILE_EXTENSION_COUNT:
        FILE_EXTENSION_COUNT[extension] += 1
        FILE_EXTENSION_TOTAL_SIZE[extension] += fileSize
    else:
        FILE_EXTENSION_COUNT[extension] = 1
        FILE_EXTENSION_TOTAL_SIZE[extension] = fileSize
    return False


def fileExtensionPost(ws):
    ws.write(0, 0, "Extensions", wbm.headerFormat)
    ws.write(0, 1, "Count", wbm.headerFormat)
    ws.write(0, 2, "Avg Size (MB)", wbm.headerFormat)

    row = 1
    for extension in sorted(FILE_EXTENSION_COUNT.keys()):
        ws.write_string(row, 0, extension)
        ws.write_number(row, 1, FILE_EXTENSION_COUNT[extension])
        ws.write_number(row, 2, round(FILE_EXTENSION_TOTAL_SIZE[extension] / FILE_EXTENSION_COUNT[extension], 1))
        wbm.incrementFileCount(ws)
        row += 1

    ws.freeze_panes(1, 0)
    ws.autofit()


def duplicateFileConcurrent(dirAbsolute:str, fileName:str, _):
    if fileName in FILES_AND_PATHS:
        FILES_AND_PATHS[fileName].add(dirAbsolute)
    else:
        FILES_AND_PATHS[fileName] = set([dirAbsolute])
    return False



def duplicateFilePost(ws):
    ws.write(0, 0, "Files", wbm.headerFormat)
    ws.write(0, 1, "Directories", wbm.headerFormat)
    
    row = 1
    for fileName in sorted(FILES_AND_PATHS.keys()):
        # if a fileName is seen more than once, it's a duplicate
        if len(FILES_AND_PATHS[fileName]) > 1:
            ws.write_string(row, 0, fileName, wbm.errorFormat)
            wbm.incrementFileCount(ws)
            
            for path in FILES_AND_PATHS[fileName]:
                ws.write_string(row, 1, path, wbm.dirColFormat)
                row += 1

    ws.freeze_panes(1, 0)
    ws.autofit()



def deleteEmptyDirectoriesLog(dirAbsolute, dirFolders, dirFiles, ws):
    tooFewAmount = wbm.fixArg

    # If even 1 folder exists, this isn't empty
    if len(dirFolders) != 0: return

    # If equal to tooFewAmount or less, then this folder needs to be at least flagged
    fileAmount = len(dirFiles)
    if fileAmount <= tooFewAmount:
        wbm.writeInCell(ws, wbm.ERROR_COL, "{}".format(fileAmount), wbm.logFormat, 1, 1)


def deleteEmptyDirectoriesModify(dirAbsolute, dirFolders, dirFiles, ws):
    tooFewAmount = wbm.fixArg

    if len(dirFolders) != 0: return

    fileAmount = len(dirFiles)
    if fileAmount <= tooFewAmount:
        # If it specifically has 0 files, delete the folder
        if (fileAmount == 0):
            try: 
                os.rmdir(dirAbsolute)
            except:
                wbm.writeInCell(ws, wbm.MOD_COL, "0 FILES. COULD NOT DELETE", wbm.errorFormat, 1, 1)
                return
            wbm.writeInCell(ws, wbm.MOD_COL, "{}".format(fileAmount), wbm.modifyFormat, 1, 1)
        # Otherwise, just flag as usual
        else:
            wbm.writeInCell(ws, wbm.MOD_COL, "{}".format(fileAmount), wbm.logFormat, 1, 1)


# TODO: This.
def searchAndReplaceHelper(oldItemName:str):
    pass


def searchAndReplaceLog(_:str, oldItemName:str, ws):
    toBeReplaced, replacer = wbm.fixArg

    lastPeriodIndex = oldItemName.rfind(".")
    if lastPeriodIndex == -1:
        extension = ""
        oldItemNameSansExt = oldItemName[0:]
    else:
        extension = oldItemName[lastPeriodIndex:]
        oldItemNameSansExt = oldItemName[0:lastPeriodIndex]
    
    newItemNameSansExt = oldItemNameSansExt.replace(toBeReplaced, replacer)

    if (oldItemNameSansExt == newItemNameSansExt): return

    newItemName = newItemNameSansExt + extension

    wbm.writeInCell(ws, wbm.ITEM_COL, oldItemName, wbm.errorFormat)
    wbm.writeInCell(ws, wbm.MOD_COL, newItemName, wbm.logFormat, 1, 1)    


def searchAndReplaceModify(dirAbsolute:str, oldItemName:str, ws):
    toBeReplaced, replacer = wbm.fixArg

    lastPeriodIndex = oldItemName.rfind(".")
    if lastPeriodIndex == -1:
        extension = ""
        oldItemNameSansExt = oldItemName[0:]
    else:
        extension = oldItemName[lastPeriodIndex:]
        oldItemNameSansExt = oldItemName[0:lastPeriodIndex]
    
    newItemNameSansExt = oldItemNameSansExt.replace(toBeReplaced, replacer)

    if (oldItemNameSansExt == newItemNameSansExt): return
    
    newItemName = newItemNameSansExt + extension

    wbm.writeInCell(ws, wbm.ITEM_COL, oldItemName, wbm.errorFormat)

    try:
        os.rename(dirAbsolute + "/" + oldItemName, dirAbsolute + "/" + newItemName)
        wbm.writeInCell(ws, wbm.MOD_COL, newItemName, wbm.modifyFormat, 1, 1)
    except PermissionError:
        wbm.writeInCell(ws, wbm.MOD_COL, "FILE LOCKED. MODIFICATION FAILED.", wbm.errorFormat, 1, 0)
    except OSError:
        # This will happen if attempting to rename to an empty string
        wbm.writeInCell(ws, wbm.MOD_COL, "OS ERROR. MODIFICATION FAILED.", wbm.errorFormat, 1, 0)
