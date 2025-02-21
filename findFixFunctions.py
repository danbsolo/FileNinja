# from typing import Set
import string
import os
from datetime import datetime
import hashlib
from workbookManager import WorkbookManager
from sys import maxsize as MAXSIZE
# import mmap


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

# # Used by duplicateName
# NAMES_AND_PATHS = {}

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
    return 2  # SPECIAL CASE

def spaceFileFind(_:str, itemName:str, ws) -> bool:
    if " " in itemName: 
        wbm.writeItemAndIncrement(ws, itemName, wbm.errorFormat)
        return True
    return False

def spaceFolderFind(dirAbsolute:str, dirFolders, dirFiles, ws):
    folderName = dirAbsolute[dirAbsolute.rfind("\\") +1:]
    
    if " " in folderName:
        wbm.writeDirAndIncrement(ws, dirAbsolute, wbm.errorFormat)


def overCharLimitFind(dirAbsolute:str, itemName:str, ws) -> bool:
    absoluteItemLength = len(dirAbsolute + "/" + itemName)
    if (absoluteItemLength > CHARACTER_LIMIT):
        wbm.writeItem(ws, itemName, wbm.errorFormat)
        wbm.writeOutcomeAndIncrement(ws, "{}".format(absoluteItemLength, CHARACTER_LIMIT))
        return True
    return False


def badCharHelper(s:str) -> set:
    badChars = set()

    for i in range(len(s)):
        if s[i] not in PERMISSIBLE_CHARACTERS:
            badChars.add(s[i])
        
        # double-dash error
        elif s[i:i+2] == "--":
            badChars.add(s[i])

    return badChars


def badCharFileFind(_:str, itemName:str, ws) -> bool:
    # If no extension (aka, no period), lastPeriodIndex will equal -1
    lastPeriodIndex = itemName.rfind(".")

    if (lastPeriodIndex == -1):
        badChars = badCharHelper(itemName)
    else:
        badChars = badCharHelper(itemName[0:lastPeriodIndex])

    # if any bad characters were found
    if (badChars):
        wbm.writeItem(ws, itemName, wbm.errorFormat)
        wbm.writeOutcomeAndIncrement(ws, "".join(badChars))
        return True
    return False


def badCharFolderFind(dirAbsolute:str, dirFolders, dirFiles, ws):
    folderName = dirAbsolute[dirAbsolute.rfind("\\") +1:]
    badChars = badCharHelper(folderName)

    if (badChars):
        wbm.writeDir(ws, dirAbsolute, wbm.errorFormat)
        wbm.writeOutcomeAndIncrement(ws, "".join(badChars))


def oldFileFind(dirAbsolute:str, itemName:str, ws):
    try:
        fileDate = datetime.fromtimestamp(os.path.getatime(dirAbsolute + "\\" + itemName))
    except Exception as e:
        wbm.writeItem(ws, itemName, wbm.errorFormat)
        wbm.writeOutcomeAndIncrement(ws, f"UNABLE TO READ DATE. {e}", wbm.errorFormat) 
        return False

    fileDaysAgoLastAccessed = (TODAY - fileDate).days

    if (fileDaysAgoLastAccessed >= DAYS_TOO_OLD):
        wbm.writeItem(ws, itemName, wbm.errorFormat)
        wbm.writeOutcomeAndIncrement(ws, "{}".format(fileDaysAgoLastAccessed))
        return True
    else:
        return False


def emptyDirectoryConcurrent(dirAbsolute:str, dirFolders, dirFiles, ws):
    if len(dirFolders) == 0 and len(dirFiles) <= TOO_FEW_AMOUNT:
        wbm.writeDirAndIncrement(ws, dirAbsolute, wbm.errorFormat)



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

def spaceFixLog(_:str, oldItemName:str, ws, _2):
    newItemName = spaceFixHelper(oldItemName)
    if (not newItemName): return

    wbm.writeItem(ws, oldItemName)
    wbm.writeOutcomeAndIncrement(ws, newItemName, wbm.logFormat)
    
def spaceFixModify(dirAbsolute:str, oldItemName:str, ws, _2):
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
        

def deleteOldFilesHelper(fullFilePath: str, arg) -> int:
    """Note that a file that is 23 hours and 59 minutes old is still considered 0 days old."""

    daysLowerBound = arg[0]
    
    # NOTE: This is not well-done code since, over the lifetime of an execution, this will always evaluate one or the other.
    if len(arg) == 2:
        daysUpperBound = arg[1]
    else:
        daysUpperBound = MAXSIZE

    # Could double-check that this value is usable each time. Dire consequences if not.
    #  if (daysLowerBound <= 0): return -1
    
    # Get date of file. This *can* error virtue of the library functions, hence try/except
    try: fileDate = datetime.fromtimestamp(os.path.getatime(fullFilePath))
    except: return -1

    fileDaysAgo = (TODAY - fileDate).days

    if (daysLowerBound <= fileDaysAgo) and (fileDaysAgo <= daysUpperBound): return fileDaysAgo
    else: return 0

def deleteOldFilesLog(dirAbsolute:str, itemName:str, ws, arg):
    fullFilePath =  dirAbsolute + "\\" + itemName
    daysOld = deleteOldFilesHelper(fullFilePath, arg)

    # Either it's actually 0 days old or the fileDate is not within the cutOffDate range. Either way, don't flag.         
    if (daysOld == 0): return

    wbm.writeItem(ws, itemName)

    if (daysOld == -1):
        wbm.writeOutcomeAndIncrement(ws, "UNABLE TO READ DATE", wbm.errorFormat) 
    else:
        wbm.writeOutcomeAndIncrement(ws, "{}".format(daysOld), wbm.logFormat)

def deleteOldFilesModify(dirAbsolute:str, itemName:str, ws, arg):
    fullFilePath =  dirAbsolute + "\\" + itemName
    daysOld = deleteOldFilesHelper(fullFilePath, arg)

    if (daysOld == 0): return

    wbm.writeItem(ws, itemName)

    if (daysOld == -1):
        wbm.writeOutcomeAndIncrement(ws, "UNABLE TO READ DATE.", wbm.errorFormat) 
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


# def duplicateNameConcurrent(dirAbsolute:str, fileName:str, ws):
#     if fileName in NAMES_AND_PATHS:
#         NAMES_AND_PATHS[fileName].add(dirAbsolute)
#         wbm.incrementFileCount(ws)
#         return True
#     else:
#         NAMES_AND_PATHS[fileName] = set([dirAbsolute])
#         return False

# def duplicateNamePost(ws):
#     ws.write(0, 0, "Files", wbm.headerFormat)
#     ws.write(0, 1, "Directories", wbm.headerFormat)
    
#     row = 1
#     for fileName in sorted(NAMES_AND_PATHS.keys()):
#         # if a fileName is seen more than once, it's a duplicate
#         if len(NAMES_AND_PATHS[fileName]) > 1:
#             ws.write_string(row, 0, fileName, wbm.errorFormat)
            
#             for path in NAMES_AND_PATHS[fileName]:
#                 ws.write_string(row, 1, path, wbm.dirFormat)
#                 row += 1

#     NAMES_AND_PATHS.clear()
#     ws.freeze_panes(1, 0)
#     ws.autofit()


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
        return 3  # SPECIAL CASE
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


def deleteEmptyDirectoriesLog(_, dirFolders, dirFiles, ws, arg):
    tooFewAmount = arg[0]

    # If even 1 folder exists, this isn't empty
    if len(dirFolders) != 0: return

    # If equal to tooFewAmount or less, then this folder needs to be at least flagged
    fileAmount = len(dirFiles)
    if fileAmount <= tooFewAmount:
        wbm.writeOutcomeAndIncrement(ws, "{}".format(fileAmount), wbm.logFormat)

def deleteEmptyDirectoriesModify(dirAbsolute, dirFolders, dirFiles, ws, arg):
    tooFewAmount = arg[0]

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


def searchAndReplaceHelper(oldItemName:str, arg):
    toBeReplaced, replacer = arg

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

def searchAndReplaceLog(_:str, oldItemName:str, ws, arg):
    if not (newItemName := searchAndReplaceHelper(oldItemName, arg)): return

    wbm.writeItem(ws, oldItemName, wbm.errorFormat)
    wbm.writeOutcomeAndIncrement(ws, newItemName, wbm.logFormat)        

def searchAndReplaceModify(dirAbsolute:str, oldItemName:str, ws, arg):
    if not (newItemName := searchAndReplaceHelper(oldItemName, arg)): return

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


def deleteEmptyFilesLog(dirAbsolute:str, itemName:str, ws, _):
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
        

def deleteEmptyFilesModify(dirAbsolute:str, itemName:str, ws, _):
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
