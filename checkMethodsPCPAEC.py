from typing import Set
import string
import os
import time
from datetime import datetime, timedelta
from workbookManager import WorkbookManager


# Used by badCharErrorCheck(). This includes space as a permissible character as it is checked by spaceErrorCheck() instead
PERMISSIBLE_CHARACTERS = set(string.ascii_letters + string.digits + "-. ")
CHARACTER_LIMIT = 200

# Used by deleteOldFilesGeneral()
TODAY = datetime.now()

# Used by fileTypeMisc() and fileTypePost()
FILE_EXTENSION_COUNT = {}
FILE_EXTENSION_TOTAL_SIZE = {}

# Used by duplicateFileMisc() and duplicateFilePost()
FILES_AND_PATHS = {}


# Declare a global variable within a function
# ~ Usually a bad idea, but here, it makes sense
def setWorkBookManager(newManager: WorkbookManager):
    global wbm
    wbm = newManager


def spaceErrorCheck(_:str, itemName:str, ws) -> bool:
    if " " in itemName:
        wbm.writeInCell(ws, wbm.ITEM_COL, itemName, wbm.fileErrorFormat, rowIncrement=1, fileIncrement=1)
        # no need to write in the error column as this won't vary between errors found
        return True
    
    return False


def overCharLimitCheck(dirAbsolute:str, itemName:str, ws) -> bool:
    absoluteItemLength = len(dirAbsolute + "/" + itemName)
    if (absoluteItemLength > CHARACTER_LIMIT):
        wbm.writeInCell(ws, wbm.ITEM_COL, itemName, wbm.fileErrorFormat)
        wbm.writeInCell(ws, wbm.ERROR_COL, "{} > {}".format(absoluteItemLength, CHARACTER_LIMIT), rowIncrement=1, fileIncrement=1)
        return True
    
    return False


def badCharErrorCheck(_:str, itemName:str, ws) -> Set[str]:
    """Does not check for SPC characters nor extra periods."""
    
    badChars = set()
    itemNameLength = len(itemName)

    for i in range(itemNameLength):
        if itemName[i] not in PERMISSIBLE_CHARACTERS:
            # Not necessary to set errorPresent=True because we're returning non-empty badChars instead
            badChars.add(itemName[i])
        
        # double dash error
        elif itemName[i:i+2] == "--":
            badChars.add(itemName[i])

    # write to own sheet here
    if (badChars):
        wbm.writeInCell(ws, wbm.ITEM_COL, itemName, wbm.fileErrorFormat)
        wbm.writeInCell(ws, wbm.ERROR_COL, "".join(badChars), rowIncrement=1, fileIncrement=1)
        
    return badChars


def spaceErrorGeneral(oldItemName) -> str:
    # if no spaces, just leave
    if (" " not in oldItemName):
        return
    
    # Replace "-" characters with " " to make the string homogenous for the upcoming split()
    # split() automatically removes leading, trailing, and excess middle whitespace
    newItemName = oldItemName.replace("-", " ").split()
    newItemName = "-".join(newItemName)

    # For if there's a dash to the left of the file extension period
    lastPeriodIndex = newItemName.rfind(".")
    # If lastPeriodIndex isn't the very first character and there actually is a period
    if (lastPeriodIndex > 0 and newItemName[lastPeriodIndex -1] == "-"):
        newItemName = newItemName[0:lastPeriodIndex-1] + newItemName[lastPeriodIndex:]

    return newItemName


def spaceErrorFixLog(_:str, oldItemName:str, ws):
    newItemName = spaceErrorGeneral(oldItemName)
    if (not newItemName): return

    # write the new name in the appropriate field
    wbm.writeInCell(ws, wbm.ITEM_COL, oldItemName)
    wbm.writeInCell(ws, wbm.RENAME_COL, newItemName, wbm.showRenameFormat, 1, 1)

    
def spaceErrorFixExecute(dirAbsolute:str, oldItemName:str, ws):
    newItemName = spaceErrorGeneral(oldItemName)
    if (not newItemName): return

    wbm.writeInCell(ws, wbm.ITEM_COL, oldItemName)

    # Log newItemName and rename file
    try:
        os.rename(dirAbsolute + "/" + oldItemName, dirAbsolute + "/" + newItemName)
        wbm.writeInCell(ws, wbm.RENAME_COL, newItemName, wbm.renameFormat, 1, 1)
    except PermissionError:
        wbm.writeInCell(ws, wbm.RENAME_COL, "FILE LOCKED. RENAME FAILED.", wbm.fileErrorFormat, 1, 0)
    except OSError:
        wbm.writeInCell(ws, wbm.RENAME_COL, "OS ERROR. RENAME FAILED.", wbm.fileErrorFormat, 1, 0)


def deleteOldFilesGeneral(fullFilePath: str) -> int:
    """Note that a file that is 23 hours and 59 minutes old is still considered 0 days old."""
    
    daysTooOld = wbm.fixArg
    # Could double-check that this value is usable. Dire consequences if not.
    #  if (daysTooOld <= 0): return -1
    
    # gets date of file. This *can* error virtue of the library functions, hence try/except
    try:
        fileDate = datetime.fromtimestamp(os.path.getatime(fullFilePath))
    except:
        return -1

    fileDaysAgo = (TODAY - fileDate).days

    if (fileDaysAgo >= daysTooOld): return fileDaysAgo
    else: return 0


def deleteOldFilesLog(dirAbsolute:str, itemName:str, ws):
    fullFilePath =  dirAbsolute + "\\" + itemName
    daysOld = deleteOldFilesGeneral(fullFilePath)

    if (daysOld == 0): return

    wbm.writeInCell(ws, wbm.ITEM_COL, itemName)

    if (daysOld == -1):
        wbm.writeInCell(ws, wbm.RENAME_COL, "UNABLE TO READ DATE", wbm.fileErrorFormat, 1, 1) 
    elif len(fullFilePath) > CHARACTER_LIMIT:
        wbm.writeInCell(ws, wbm.RENAME_COL, "{} days ago, but violates charLimit".format(daysOld), wbm.showRenameFormat, 1, 1)
    else:
        wbm.writeInCell(ws, wbm.RENAME_COL, "{} days ago".format(daysOld), wbm.showRenameFormat, 1, 1)    


def deleteOldFilesExecute(dirAbsolute:str, itemName:str, ws):
    fullFilePath =  dirAbsolute + "\\" + itemName
    daysOld = deleteOldFilesGeneral(fullFilePath)

    # Either it's actually 0 days old or the fileDate is after the cutOffDate. Either way, don't flag.         
    if (daysOld == 0):
        return

    wbm.writeInCell(ws, wbm.ITEM_COL, itemName)

    if (daysOld == -1):
        wbm.writeInCell(ws, wbm.RENAME_COL, "UNABLE TO READ DATE", wbm.fileErrorFormat, 1, 1) 
    # If over CHARACTER_LIMIT characters, do not delete as it is not backed up
    elif len(fullFilePath) > CHARACTER_LIMIT:
        wbm.writeInCell(ws, wbm.RENAME_COL, "{} days, but violates charLimit".format(daysOld), wbm.showRenameFormat, 1, 1)
    else:
        try:
            os.remove(fullFilePath)
            wbm.writeInCell(ws, wbm.RENAME_COL, "{} days".format(daysOld), wbm.renameFormat, 1, 1)
        except:
            wbm.writeInCell(ws, wbm.RENAME_COL, "FAILED TO DELETE", wbm.fileErrorFormat, 1, 1)


def fileExtensionMisc(dirAbsolute:str, itemName:str):
    _, ext = os.path.splitext(itemName)
    ext = ext.lower()
    fileSize = os.path.getsize(dirAbsolute+"/"+itemName) / 1000_000  # bytes / 1000_000 = mbs

    if ext in FILE_EXTENSION_COUNT:
        FILE_EXTENSION_COUNT[ext] += 1
        FILE_EXTENSION_TOTAL_SIZE[ext] += fileSize
    else:
        FILE_EXTENSION_COUNT[ext] = 1
        FILE_EXTENSION_TOTAL_SIZE[ext] = fileSize


def fileExtensionPost(ws):
    ws.write(0, 0, "Extensions", wbm.headerFormat)
    ws.write(0, 1, "Count", wbm.headerFormat)
    ws.write(0, 2, "Avg Size (KB)", wbm.headerFormat)

    row = 1
    for ext in sorted(FILE_EXTENSION_COUNT.keys()):
        ws.write_string(row, 0, ext)
        ws.write_number(row, 1, FILE_EXTENSION_COUNT[ext])
        ws.write_number(row, 2, round(FILE_EXTENSION_TOTAL_SIZE[ext] / FILE_EXTENSION_COUNT[ext], 1))
        row += 1

    ws.freeze_panes(1, 0)
    ws.autofit()


def duplicateFileMisc(dirAbsolute:str, fileName:str):
    if fileName in FILES_AND_PATHS:
        FILES_AND_PATHS[fileName].append(dirAbsolute)
    else:
        FILES_AND_PATHS[fileName] = [dirAbsolute]


def duplicateFilePost(ws):
    ws.write(0, 0, "Files", wbm.headerFormat)
    ws.write(0, 1, "Directories", wbm.headerFormat)
    
    row = 1
    for itemName in sorted(FILES_AND_PATHS.keys()):
        if len(FILES_AND_PATHS[itemName]) > 1:  # if so, it's duplicated
            ws.write_string(row, 0, itemName, wbm.fileErrorFormat)

            for path in FILES_AND_PATHS[itemName]:
                ws.write_string(row, 1, path, wbm.dirColFormat)
                row += 1

    ws.freeze_panes(1, 0)
    ws.autofit()



def listAll(_:str, itemName:str, ws):
    wbm.writeInCell(ws, wbm.ITEM_COL, itemName, rowIncrement=1, fileIncrement=1)
