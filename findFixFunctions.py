from typing import Set
import string
import os
from datetime import datetime
import hashlib
from workbookManager import WorkbookManager
import mmap


# Used by badCharFind
# Includes ' ' (space) as there is a separate procedure for finding that error
PERMISSIBLE_CHARACTERS = set(string.ascii_letters + string.digits + "- ")
CHARACTER_LIMIT = 200

# Used by oldFileFind
DAYS_TOO_OLD = 365

# Used by oldFileFind and deleteOldFiles
TODAY = datetime.now()

# Used by emptyDirectoryConcurrent
TOO_FEW_AMOUNT = 0

# Used by fileExtension
EXTENSION_COUNT = {}
EXTENSION_TOTAL_SIZE = {}

# Used by duplicateName
NAMES_AND_PATHS = {}

# Used by duplicateContent
HASH_AND_FILES = {}
# MMAP_THRESHOLD = 8 * 1024 * 1024  # 8MB to Bytes
hashFunc = hashlib.new("sha256")
hashFunc.update("".encode())
EMPTY_INPUT_HASH_CODE = hashFunc.hexdigest()


def setWorkbookManager(newManager: WorkbookManager):
    # Globally declare the WorkboookManager object
    global wbm
    wbm = newManager
    

def listAll(_:str, itemName:str, ws) -> bool:
    wbm.writeItemAndIncrement(ws, itemName)
    return False

def spaceFind(_:str, itemName:str, ws) -> bool:
    if " " in itemName: 
        wbm.writeItemAndIncrement(ws, itemName, wbm.errorFormat)
        return True
    return False
    
def overCharLimitFind(dirAbsolute:str, itemName:str, ws) -> bool:
    absoluteItemLength = len(dirAbsolute + "/" + itemName)
    if (absoluteItemLength > CHARACTER_LIMIT):
        wbm.writeItem(ws, itemName, wbm.errorFormat)
        wbm.writeOutcomeAndIncrement(ws, "{} > {}".format(absoluteItemLength, CHARACTER_LIMIT))
        return True
    return False

def badCharFind(_:str, itemName:str, ws) -> bool:
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
        wbm.writeItem(ws, itemName, wbm.errorFormat)
        wbm.writeOutcomeAndIncrement(ws, "".join(badChars))
        return True
    return False

def oldFileFind(dirAbsolute:str, itemName:str, ws):
    try:
        fileDate = datetime.fromtimestamp(os.path.getatime(dirAbsolute + "\\" + itemName))
    except Exception as e:
        wbm.writeItem(ws, itemName, wbm.errorFormat)
        wbm.writeOutcomeAndIncrement(ws, f"UNABLE TO READ DATE. {e}", wbm.errorFormat) 
        return False

    fileDaysAgo = (TODAY - fileDate).days

    if (fileDaysAgo >= DAYS_TOO_OLD):
        wbm.writeItem(ws, itemName, wbm.errorFormat)
        wbm.writeOutcomeAndIncrement(ws, "{} days old".format(fileDaysAgo))
    else:
        return False


def emptyDirectoryConcurrent(dirAbsolute:str, dirFolders, dirFiles, ws):
    if len(dirFolders) == 0 and len(dirFiles) <= TOO_FEW_AMOUNT:
        wbm.writeDirAndIncrement(ws, dirAbsolute, wbm.errorFormat)

def emptyDirectoryPost(ws):
    ws.autofit()


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

    wbm.writeItem(ws, oldItemName)
    wbm.writeOutcomeAndIncrement(ws, newItemName, wbm.logFormat)
    
def spaceFixModify(dirAbsolute:str, oldItemName:str, ws):
    newItemName = spaceFixHelper(oldItemName)
    if (not newItemName): return

    wbm.writeItem(ws, oldItemName)

    # Log newItemName and rename file
    try:
        os.rename(dirAbsolute + "/" + oldItemName, dirAbsolute + "/" + newItemName)
        wbm.writeOutcomeAndIncrement(ws, newItemName, wbm.modifyFormat)
    except PermissionError:
        wbm.writeOutcome(ws, "MODIFICATION FAILED. PERMISSION ERROR.", wbm.errorFormat)
        wbm.incrementRow(ws)
    except Exception as e:
        wbm.writeOutcome(ws, f"MODIFICATION FAILED. {e}", wbm.errorFormat)
        wbm.incrementRow(ws)
        

def deleteOldFilesHelper(fullFilePath: str) -> int:
    """Note that a file that is 23 hours and 59 minutes old is still considered 0 days old."""
    
    daysTooOld = wbm.fixArg
    # Could double-check that this value is usable each time. Dire consequences if not.
    #  if (daysTooOld <= 0): return -1
    
    # Get date of file. This *can* error virtue of the library functions, hence try/except
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

    wbm.writeItem(ws, itemName)

    if (daysOld == -1):
        wbm.writeOutcomeAndIncrement(ws, "UNABLE TO READ DATE", wbm.errorFormat) 
    elif len(fullFilePath) > CHARACTER_LIMIT:
        wbm.writeOutcomeAndIncrement(ws, "{} days but violates charLimit".format(daysOld), wbm.logFormat)    
    else:
        wbm.writeOutcomeAndIncrement(ws, "{}".format(daysOld), wbm.logFormat)

def deleteOldFilesModify(dirAbsolute:str, itemName:str, ws):
    fullFilePath =  dirAbsolute + "\\" + itemName
    daysOld = deleteOldFilesHelper(fullFilePath)

    if (daysOld == 0): return

    wbm.writeItem(ws, itemName)

    if (daysOld == -1):
        wbm.writeOutcomeAndIncrement(ws, "UNABLE TO READ DATE.", wbm.errorFormat) 
    # If over CHARACTER_LIMIT characters, do not delete as it is not backed up
    elif len(fullFilePath) > CHARACTER_LIMIT:
        wbm.writeOutcomeAndIncrement(ws, "{} days but violates charLimit".format(daysOld), wbm.logFormat)
    else:
        try:
            os.remove(fullFilePath)
            wbm.writeOutcomeAndIncrement(ws, "{}".format(daysOld), wbm.modifyFormat)
        except PermissionError:
            wbm.writeOutcomeAndIncrement(ws, "FAILED TO DELETE. PERMISSION ERROR.", wbm.errorFormat)
        except Exception as e:
            wbm.writeOutcomeAndIncrement(ws, f"FAILED TO DELETE. {e}", wbm.errorFormat)
            

def fileExtensionConcurrent(dirAbsolute:str, itemName:str, _):
    try: fileSize = os.path.getsize(dirAbsolute+"/"+itemName) / 1000_000  # Bytes / 1000_000 = MBs
    except: return False

    lastPeriodIndex = itemName.rfind(".")

    if lastPeriodIndex == -1: extension = ""
    else:
        extension = itemName[lastPeriodIndex:]
        extension = extension.lower()

    if extension in EXTENSION_COUNT:
        EXTENSION_COUNT[extension] += 1
        EXTENSION_TOTAL_SIZE[extension] += fileSize
    else:
        EXTENSION_COUNT[extension] = 1
        EXTENSION_TOTAL_SIZE[extension] = fileSize
    return False

def fileExtensionPost(ws):
    ws.write(0, 0, "Extensions", wbm.headerFormat)
    ws.write(0, 1, "Count", wbm.headerFormat)
    ws.write(0, 2, "Avg Size (MB)", wbm.headerFormat)

    row = 1
    for extension in sorted(EXTENSION_COUNT.keys()):
        ws.write_string(row, 0, extension)
        ws.write_number(row, 1, EXTENSION_COUNT[extension])
        ws.write_number(row, 2, round(EXTENSION_TOTAL_SIZE[extension] / EXTENSION_COUNT[extension], 1))
        wbm.incrementFileCount(ws)
        row += 1

    EXTENSION_COUNT.clear()
    EXTENSION_TOTAL_SIZE.clear()
    ws.freeze_panes(1, 0)
    ws.autofit()


def duplicateNameConcurrent(dirAbsolute:str, fileName:str, ws):
    if fileName in NAMES_AND_PATHS:
        NAMES_AND_PATHS[fileName].add(dirAbsolute)
        wbm.incrementFileCount(ws)
        return True
    else:
        NAMES_AND_PATHS[fileName] = set([dirAbsolute])
        return False

def duplicateNamePost(ws):
    ws.write(0, 0, "Files", wbm.headerFormat)
    ws.write(0, 1, "Directories", wbm.headerFormat)
    
    row = 1
    for fileName in sorted(NAMES_AND_PATHS.keys()):
        # if a fileName is seen more than once, it's a duplicate
        if len(NAMES_AND_PATHS[fileName]) > 1:
            ws.write_string(row, 0, fileName, wbm.errorFormat)
            
            for path in NAMES_AND_PATHS[fileName]:
                ws.write_string(row, 1, path, wbm.dirFormat)
                row += 1

    NAMES_AND_PATHS.clear()
    ws.freeze_panes(1, 0)
    ws.autofit()


def duplicateContentHelper(dirAbsolute:str, itemName:str):
    # try: fileSize = os.path.getsize(dirAbsolute+"/"+itemName)  # Bytes
    # except: fileSize = 0 # TODO: check average size of files that cause this
    
    hashFunc = hashlib.new("sha256")
    
    try: 
        with open(dirAbsolute+"/"+itemName, "rb") as file:
            # if fileSize > MMAP_THRESHOLD:
                # mmappedFile = mmap.mmap(file.fileno(), 0, access=mmap.ACCESS_READ)
                # print(fileSize, "B", sep="", end=" ")
                # while chunk := mmappedFile.read(16384):
                    # hashFunc.update(chunk)
            # else:
            # print(fileSize, "B", sep="", end=" ")
            while chunk := file.read(8192):
                hashFunc.update(chunk)

        return hashFunc.hexdigest()
    # FileNotFoundError, PermissionError, OSError, UnicodeDecodeError
    except: return

def duplicateContentConcurrent(dirAbsolute:str, itemName:str, ws):
    hashCode = duplicateContentHelper(dirAbsolute, itemName)

    if (not hashCode or hashCode == EMPTY_INPUT_HASH_CODE):
        return False
    
    if hashCode in HASH_AND_FILES:
        HASH_AND_FILES[hashCode][0].append(itemName)
        HASH_AND_FILES[hashCode][1].append(dirAbsolute)
        wbm.incrementFileCount(ws)
        return True
    else:
        HASH_AND_FILES[hashCode] = ([itemName], [dirAbsolute])
        return False

def duplicateContentPost(ws):
    ws.write(0, 0, "Separator", wbm.headerFormat)
    ws.write(0, 1, "Files", wbm.headerFormat)
    ws.write(0, 2, "Directories", wbm.headerFormat)

    row = 1
    for hashCode in HASH_AND_FILES.keys():
        if (numOfFiles := len(HASH_AND_FILES[hashCode][0])) > 1:
            ws.write(row, 0, "---"*4, wbm.logFormat)

            for i in range(numOfFiles):
                ws.write(row, 1, HASH_AND_FILES[hashCode][0][i], wbm.errorFormat)
                ws.write(row, 2, HASH_AND_FILES[hashCode][1][i], wbm.dirFormat)    
                row += 1
                
    HASH_AND_FILES.clear()
    ws.freeze_panes(1, 0)
    ws.autofit()


def deleteEmptyDirectoriesLog(_, dirFolders, dirFiles, ws):
    tooFewAmount = wbm.fixArg

    # If even 1 folder exists, this isn't empty
    if len(dirFolders) != 0: return

    # If equal to tooFewAmount or less, then this folder needs to be at least flagged
    fileAmount = len(dirFiles)
    if fileAmount <= tooFewAmount:
        wbm.writeOutcomeAndIncrement(ws, "{}".format(fileAmount), wbm.logFormat)

def deleteEmptyDirectoriesModify(dirAbsolute, dirFolders, dirFiles, ws):
    tooFewAmount = wbm.fixArg

    if len(dirFolders) != 0: return

    fileAmount = len(dirFiles)
    if fileAmount <= tooFewAmount:
        # If it specifically has 0 files, delete the folder
        if (fileAmount == 0):
            try: 
                os.rmdir(dirAbsolute)
                wbm.writeOutcomeAndIncrement(ws, "{}".format(fileAmount), wbm.modifyFormat)
            except Exception as e:
                wbm.writeOutcomeAndIncrement(ws, f"0 FILES. UNABLE TO DELETE. {e}", wbm.errorFormat)
                return
        # Otherwise, just flag as usual
        else:
            wbm.writeOutcomeAndIncrement(ws, "{}".format(fileAmount), wbm.logFormat)


def searchAndReplaceHelper(oldItemName:str):
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
    return newItemNameSansExt + extension

def searchAndReplaceLog(_:str, oldItemName:str, ws):
    if not (newItemName := searchAndReplaceHelper(oldItemName)): return

    wbm.writeItem(ws, oldItemName, wbm.errorFormat)
    wbm.writeOutcomeAndIncrement(ws, newItemName, wbm.logFormat)        

def searchAndReplaceModify(dirAbsolute:str, oldItemName:str, ws):
    if not (newItemName := searchAndReplaceHelper(oldItemName)): return

    wbm.writeItem(ws, oldItemName, wbm.errorFormat)

    try:
        os.rename(dirAbsolute + "/" + oldItemName, dirAbsolute + "/" + newItemName)
        wbm.writeOutcomeAndIncrement(ws, newItemName, wbm.modifyFormat)
    except PermissionError:
        wbm.writeOutcome(ws, "MODIFICATION FAILED. PERMISSION ERROR.", wbm.errorFormat)
        wbm.incrementRow(ws)
    except Exception as e:
        # This will happen if attempting to rename to an empty string
        wbm.writeOutcome(ws, f"MODIFICATION FAILED. {e}", wbm.errorFormat)
        wbm.incrementRow(ws)


def deleteEmptyFilesLog(dirAbsolute:str, itemName:str, ws):
    try:
        fileSize = os.path.getsize(dirAbsolute+"\\"+itemName) 
    except PermissionError:
        wbm.writeItem(ws, itemName)
        wbm.writeOutcomeAndIncrement(ws, "UNABLE TO READ. PERMISSION ERROR.", wbm.errorFormat)
        return
    except Exception as e:
        wbm.writeItem(ws, itemName)
        wbm.writeOutcomeAndIncrement(ws, f"UNABLE TO READ. {e}", wbm.errorFormat)
        return

    if fileSize == 0:
        wbm.writeItem(ws, itemName)
        wbm.writeOutcomeAndIncrement(ws, "", wbm.logFormat)
        

def deleteEmptyFilesModify(dirAbsolute:str, itemName:str, ws):
    """Glitch exists in that the current excel file will be considered empty.
    However, despite claiming so, the program does not actually delete it.'"""

    fullFilePath =  dirAbsolute + "\\" + itemName

    try:
        fileSize = os.path.getsize(fullFilePath)  # Bytes
    except PermissionError:
        wbm.writeItem(ws, itemName)
        wbm.writeOutcomeAndIncrement(ws, "UNABLE TO READ. PERMISSION ERROR.", wbm.errorFormat)
        return
    except Exception as e:
        wbm.writeItem(ws, itemName)
        wbm.writeOutcomeAndIncrement(ws, f"UNABLE TO READ. {e}", wbm.errorFormat)
        return

    # Stage for deletion
    if fileSize == 0:
        wbm.writeItem(ws, itemName)

        try:
            os.remove(fullFilePath)
            wbm.writeOutcomeAndIncrement(ws, "", wbm.modifyFormat)
        except PermissionError:
            wbm.writeOutcomeAndIncrement(ws, "FAILED TO DELETE. PERMISSION ERROR.", wbm.errorFormat)
        except Exception as e:
            wbm.writeOutcomeAndIncrement(ws, f"FAILED TO DELETE. {e}", wbm.errorFormat)


def emptyFileFind(dirAbsolute:str, itemName:str, ws):
    try:
        fileSize = os.path.getsize(dirAbsolute+"\\"+itemName)
    except PermissionError:
        wbm.writeItem(ws, itemName)
        wbm.writeOutcomeAndIncrement(ws, "UNABLE TO READ. PERMISSION ERROR.", wbm.errorFormat)
        return False
    except Exception as e:
        wbm.writeItem(ws, itemName)
        wbm.writeOutcomeAndIncrement(ws, f"UNABLE TO READ. {e}", wbm.errorFormat)
        return False
    
    if fileSize == 0:
        wbm.writeItemAndIncrement(ws, itemName, wbm.errorFormat)
        return True