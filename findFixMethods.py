from typing import Set
import string
import os
import time
from datetime import datetime, timedelta
from workbookManager import WorkbookManager


# Used by badCharErrorCheck().
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
    # ListAll is a special case in that it wants to return False so as to not increment the errorCount
    # But then needs to write the itemName and increment the row itself
    wbm.writeInCell(ws, wbm.ITEM_COL, itemName, rowIncrement=1)
    return False


def spaceErrorFind(_:str, itemName:str, ws) -> bool:
    if " " in itemName:
        # Nothing is explicitly written in the error column here
        return True
    
    return False


def overCharLimitFind(dirAbsolute:str, itemName:str, ws) -> bool:
    absoluteItemLength = len(dirAbsolute + "/" + itemName)
    if (absoluteItemLength > CHARACTER_LIMIT):
        wbm.writeError(ws, "{} > {}".format(absoluteItemLength, CHARACTER_LIMIT))
        return True
    
    return False


def badCharErrorFind(_:str, itemName:str, ws) -> Set[str]:
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
        wbm.writeError(ws, "".join(badChars))
        return True
        
    return False



def spaceErrorGeneral(oldItemName) -> str:
    if (" " not in oldItemName):
        return

    lastPeriodIndex = oldItemName.rfind(".")

    # Replace '-' characters with ' ' to make the string homogenous for the upcoming split()
    # split() automatically removes leading, trailing, and excess middle whitespace
    newItemNameSansExt = "-".join(oldItemName[0:lastPeriodIndex].replace("-", " ").split())

    return newItemNameSansExt + oldItemName[lastPeriodIndex:]


def spaceErrorFixLog(_:str, oldItemName:str, ws):
    newItemName = spaceErrorGeneral(oldItemName)
    if (not newItemName): return

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
            
    

def fileExtensionMisc(dirAbsolute:str, itemName:str, _):
    # Like other os operations, this may randomly crash. The solution here is to catch and return, ignoring the file completely
    # However, for this function, it's okay. We're just looking for general averages; a couple values is fine to ignore
    # Honestly, since this likes to error sometimes, could just use the ol' rfind(".") method to separate the extension
    try: _, ext = os.path.splitext(itemName)
    except: return

    ext = ext.lower()

    try: fileSize = os.path.getsize(dirAbsolute+"\\"+itemName) / 1000_000  # bytes / 1000_000 = mbs
    except: return

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
        wbm.incrementFileCount(ws)
        row += 1

    ws.freeze_panes(1, 0)
    ws.autofit()


def duplicateFileMisc(dirAbsolute:str, fileName:str, _):
    if fileName in FILES_AND_PATHS:
        FILES_AND_PATHS[fileName].add(dirAbsolute)
    else:
        FILES_AND_PATHS[fileName] = set([dirAbsolute])



def duplicateFilePost(ws):
    ws.write(0, 0, "Files", wbm.headerFormat)
    ws.write(0, 1, "Directories", wbm.headerFormat)
    
    row = 1
    for itemName in sorted(FILES_AND_PATHS.keys()):
        if len(FILES_AND_PATHS[itemName]) > 1:  # if so, it's duplicated
            ws.write_string(row, 0, itemName, wbm.fileErrorFormat)
            wbm.incrementFileCount(ws)
            
            for path in FILES_AND_PATHS[itemName]:
                ws.write_string(row, 1, path, wbm.dirColFormat)
                row += 1

    ws.freeze_panes(1, 0)
    ws.autofit()



def deleteEmptyDirectoriesLog(dirAbsolute, dirFolders, dirFiles, ws):
    tooFewAmount = wbm.fixArg

    # If even 1 folder exists, this isn't empty
    if len(dirFolders) != 0:
        return

    # If equal to tooFewAmount or less, then this folder needs to be at least flagged
    fileAmount = len(dirFiles)
    if fileAmount <= tooFewAmount:
        wbm.writeInCell(ws, wbm.ERROR_COL, "{} files".format(fileAmount), wbm.showRenameFormat, 1, 1)


def deleteEmptyDirectoriesExecute(dirAbsolute, dirFolders, dirFiles, ws):
    tooFewAmount = wbm.fixArg

    if len(dirFolders) != 0:
        return

    fileAmount = len(dirFiles)
    if fileAmount <= tooFewAmount:

        # If it specifically has 0 files, delete the folder
        if (fileAmount == 0):
            try: 
                os.rmdir(dirAbsolute)
            except:
                wbm.writeInCell(ws, wbm.ERROR_COL, "{} FILES. COULD NOT DELETE".format(fileAmount), wbm.fileErrorFormat, 1, 1)
                return
            wbm.writeInCell(ws, wbm.ERROR_COL, "{} FILES".format(fileAmount), wbm.renameFormat, 1, 1)
        else: # Otherwise, just flag as usual
            wbm.writeInCell(ws, wbm.ERROR_COL, "{} files".format(fileAmount), wbm.showRenameFormat, 1, 1)

