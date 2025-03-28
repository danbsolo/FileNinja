# from typing import Set
import string
import os
from datetime import datetime
import hashlib
from workbookManager import WorkbookManager
from sys import maxsize as MAXSIZE
# from importlib import import_module
from findOwner import getOwnerCatch
# import mmap


# Used by badCharFind
# Includes ' ' (space) as there is a separate procedure for finding that error
PERMISSIBLE_CHARACTERS = set(string.ascii_letters + string.digits + "- ")
CHARACTER_LIMIT = 200

# Used by oldFileFind
DAYS_TOO_OLD = 1095

# Used by oldFileFind and deleteOldFiles
TODAY = datetime.now()

# Used by emptyDirectoryConcurrent
TOO_FEW_AMOUNT = 0

# Used by fileExtension
EXTENSION_COUNT = {}
EXTENSION_TOTAL_SIZE = {}
TOO_LARGE_SIZE_MB = 100

# # Used by duplicateName
# NAMES_AND_PATHS = {}

# Used by duplicateContent
HASH_AND_FILES = {}
# MMAP_THRESHOLD = 8 * 1024 * 1024  # 8MB to Bytes
hashFunc = hashlib.new("sha256")
hashFunc.update("".encode())
EMPTY_INPUT_HASH_CODE = hashFunc.hexdigest()

# Used by spaceFolderFixModify or searchAndReplaceFolderModify (not logs)
FOLDER_RENAMES = []



def setWorkbookManager(newManager: WorkbookManager):
    # Globally declare the WorkboookManager object
    global wbm
    wbm = newManager

# def importModuleDynamically(moduleName, attributes):
#     moduleLoaded = import_module(moduleName)  # dynamically import module
    
#     for attr in attributes:
#         globals()[attr] = getattr(moduleLoaded, attr)  # globalize all attributes imported

# def importGetOwner():
#     # IS CRASHING
#     return
#     importModuleDynamically("findOwner", ["getOwnerCatch"])  # from findOwner import getOwnerCatch
#     # could add "getOwner" if don't want automatic catch

def writeOwnerHeader(ws):
    # importGetOwner()
    ws.write(0, wbm.AUXILIARY_COL, "Owner", wbm.headerFormat)




def listAll(_1:str, _2:str, itemName:str, ws) -> bool:
    wbm.writeItemAndIncrement(ws, itemName)
    return 2  # SPECIAL CASE


def listAllOwner(longFileAbsolute:str, _:str, itemName:str, ws) -> bool:
    wbm.writeItem(ws, itemName)
    wbm.writeAuxiliaryAndIncrement(ws, getOwnerCatch(longFileAbsolute))

    return 2  # SPECIAL CASE


def spaceFileFind(_:str, _:str, itemName:str, ws) -> bool:
    if " " in itemName: 
        wbm.writeItemAndIncrement(ws, itemName, wbm.errorFormat)
        return True
    return False

def spaceFolderFind(dirAbsolute:str, dirFolders, dirFiles, ws):
    folderName = dirAbsolute[dirAbsolute.rfind("\\") +1:]
    
    if " " in folderName:
        wbm.writeDirAndIncrement(ws, dirAbsolute, wbm.errorFormat)


def overCharLimitFind(_:str, dirAbsolute:str, itemName:str, ws) -> bool:
    absoluteItemLength = len(dirAbsolute + "/" + itemName)
    if (absoluteItemLength > CHARACTER_LIMIT):
        wbm.writeItem(ws, itemName, wbm.errorFormat)
        wbm.writeOutcomeAndIncrement(ws, absoluteItemLength)
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


def badCharFileFind(_1:str, _2:str, itemName:str, ws) -> bool:
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


def addLongPathPrefix(dirAbsolute):
    if dirAbsolute.startswith('\\\\'):
        return '\\\\?\\UNC' + dirAbsolute[1:]
    else:
        return '\\\\?\\' + dirAbsolute


def oldFileFind(longFileAbsolute:str, _:str, itemName:str, ws):
    try:
        fileDate = datetime.fromtimestamp(os.path.getatime(longFileAbsolute + "\\" + itemName))
    except Exception as e:
        wbm.writeItem(ws, itemName, wbm.errorFormat)
        wbm.writeOutcome(ws, f"UNABLE TO READ DATE. {e}", wbm.errorFormat) 
        wbm.incrementRow(ws)
        return 2

    fileDaysAgoLastAccessed = (TODAY - fileDate).days

    if (fileDaysAgoLastAccessed >= DAYS_TOO_OLD):
        wbm.writeAuxiliary(ws, getOwnerCatch(longFileAbsolute))
        wbm.writeItem(ws, itemName, wbm.errorFormat)
        wbm.writeOutcomeAndIncrement(ws, "{}".format(fileDaysAgoLastAccessed))
        return True
    else:
        return False


def emptyDirectoryConcurrent(dirAbsolute:str, dirFolders, dirFiles, ws):
    if len(dirFolders) == 0 and len(dirFiles) <= TOO_FEW_AMOUNT:
        wbm.writeDirAndIncrement(ws, dirAbsolute, wbm.errorFormat)



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

def spaceFileFixLog(_:str, oldItemName:str, ws, _2):
    newItemName = spaceFileFixHelper(oldItemName)
    if (not newItemName): return False

    wbm.writeItem(ws, oldItemName)
    wbm.writeOutcomeAndIncrement(ws, newItemName, wbm.logFormat)
    return True
    
def spaceFileFixModify(dirAbsolute:str, oldItemName:str, ws, _2):
    newItemName = spaceFileFixHelper(oldItemName)
    if (not newItemName): return False

    wbm.writeItem(ws, oldItemName)

    # Log newItemName and rename file
    try:
        os.rename(addLongPathPrefix(dirAbsolute) + "\\" + oldItemName, addLongPathPrefix(dirAbsolute) + "\\" + newItemName)
        wbm.writeOutcomeAndIncrement(ws, newItemName, wbm.modifyFormat)
    except PermissionError:
        wbm.writeOutcomeAndIncrement(ws, "MODIFICATION FAILED. PERMISSION ERROR.", wbm.errorFormat)
    except Exception as e:
        wbm.writeOutcomeAndIncrement(ws, f"MODIFICATION FAILED. {e}", wbm.errorFormat)
    return True
        

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
    try: fileDate = datetime.fromtimestamp(os.path.getatime(addLongPathPrefix(fullFilePath)))
    except: return -1

    fileDaysAgo = (TODAY - fileDate).days

    if (daysLowerBound <= fileDaysAgo) and (fileDaysAgo <= daysUpperBound): return fileDaysAgo
    else: return 0

def deleteOldFilesLog(dirAbsolute:str, itemName:str, ws, arg):
    fullFilePath =  dirAbsolute + "\\" + itemName
    daysOld = deleteOldFilesHelper(fullFilePath, arg)

    # Either it's actually 0 days old or the fileDate is not within the cutOffDate range. Either way, don't flag.         
    if (daysOld == 0): return False

    wbm.writeItem(ws, itemName)

    if (daysOld == -1):
        wbm.writeOutcome(ws, "UNABLE TO READ DATE", wbm.errorFormat)
        wbm.incrementRow(ws)
    else:
        wbm.writeAuxiliary(ws, getOwnerCatch(dirAbsolute))
        wbm.writeOutcomeAndIncrement(ws, daysOld, wbm.logFormat)
    return True

def deleteOldFilesModify(dirAbsolute:str, itemName:str, ws, arg):
    fullFilePath =  dirAbsolute + "\\" + itemName
    daysOld = deleteOldFilesHelper(fullFilePath, arg)

    if (daysOld == 0): return False

    wbm.writeItem(ws, itemName)

    if (daysOld == -1):
        wbm.writeOutcome(ws, "UNABLE TO READ DATE.", wbm.errorFormat) 
        wbm.incrementRow(ws)
    else:
        try:
            wbm.writeAuxiliary(ws, getOwnerCatch(dirAbsolute))
            os.remove(addLongPathPrefix(fullFilePath))
            wbm.writeOutcomeAndIncrement(ws, daysOld, wbm.modifyFormat)
        except PermissionError:
            wbm.writeOutcomeAndIncrement(ws, "FAILED TO DELETE. PERMISSION ERROR.", wbm.errorFormat)
        except Exception as e:
            wbm.writeOutcomeAndIncrement(ws, f"FAILED TO DELETE. {e}", wbm.errorFormat)
    return True

def fileExtensionConcurrent(longFileAbsolute:str, dirAbsolute:str, itemName:str, _):
    try: fileSize = os.path.getsize(addLongPathPrefix(dirAbsolute)+"\\"+itemName) / 1000_000  # Bytes / 1000_000 = MBs
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
    ws.write(0, 3, "Total Size (MB)", wbm.headerFormat)

    row = 1
    for extension in sorted(EXTENSION_COUNT.keys()):
        averageSize = round(EXTENSION_TOTAL_SIZE[extension] / EXTENSION_COUNT[extension], 1)

        # only count as an error if the average size is above a threshold
        if (averageSize >= TOO_LARGE_SIZE_MB):
            ws.write_string(row, 0, extension, wbm.warningFormat)
            wbm.incrementFileCount(ws)
        else:
            ws.write_string(row, 0, extension)

        ws.write_number(row, 1, EXTENSION_COUNT[extension])
        ws.write_number(row, 2, averageSize)
        ws.write_number(row, 3, round(EXTENSION_TOTAL_SIZE[extension], 1))
        row += 1

    EXTENSION_COUNT.clear()
    EXTENSION_TOTAL_SIZE.clear()
    ws.freeze_panes(1, 0)


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
    hashFunc = hashlib.new("sha256")
    
    with open(addLongPathPrefix(dirAbsolute)+"\\"+itemName, "rb") as file:
        while chunk := file.read(8192):
            hashFunc.update(chunk)

    return hashFunc.hexdigest()

def duplicateContentConcurrent(longFileAbsolute:str, dirAbsolute:str, itemName:str, ws):
    try:
        hashCode = duplicateContentHelper(dirAbsolute, itemName)
    except Exception:  # FileNotFoundError, PermissionError, OSError, UnicodeDecodeError
        # Unlike other procedures, this won't print out the error; it'll just assume it's not a duplicate.
        return False
    
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
    ws.write(0, 3, "Owner", wbm.headerFormat)

    row = 1
    for hashCode in HASH_AND_FILES.keys():
        if (numOfFiles := len(HASH_AND_FILES[hashCode][0])) > 1:
            ws.write(row, 0, "---"*4, wbm.logFormat)

            for i in range(numOfFiles):
                ws.write(row, 1, HASH_AND_FILES[hashCode][0][i], wbm.errorFormat)
                ws.write(row, 2, HASH_AND_FILES[hashCode][1][i], wbm.dirFormat)
                ws.write(row, 3, getOwnerCatch(HASH_AND_FILES[hashCode][1][i]))

                
                row += 1
                
    HASH_AND_FILES.clear()
    ws.freeze_panes(1, 0)


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

def searchAndReplaceFileLog(_:str, oldItemName:str, ws, arg):
    if not (newItemName := searchAndReplaceFileHelper(oldItemName, arg)): return False

    wbm.writeItem(ws, oldItemName, wbm.errorFormat)
    wbm.writeOutcomeAndIncrement(ws, newItemName, wbm.logFormat)
    return True

def searchAndReplaceFileModify(dirAbsolute:str, oldItemName:str, ws, arg):
    if not (newItemName := searchAndReplaceFileHelper(oldItemName, arg)): return False

    wbm.writeItem(ws, oldItemName, wbm.errorFormat)

    try:
        os.rename(addLongPathPrefix(dirAbsolute) + "\\" + oldItemName, addLongPathPrefix(dirAbsolute) + "\\" + newItemName)
        wbm.writeOutcomeAndIncrement(ws, newItemName, wbm.modifyFormat)
    except PermissionError:
        wbm.writeOutcomeAndIncrement(ws, "MODIFICATION FAILED. PERMISSION ERROR.", wbm.errorFormat)
    except Exception as e:
        wbm.writeOutcomeAndIncrement(ws, f"MODIFICATION FAILED. {e}", wbm.errorFormat)
    return True

def deleteEmptyFilesLog(dirAbsolute:str, itemName:str, ws, _):
    try:
        fileSize = os.path.getsize(addLongPathPrefix(dirAbsolute)+"\\"+itemName) 
    except PermissionError:
        wbm.writeItem(ws, itemName)
        wbm.writeOutcome(ws, "UNABLE TO READ. PERMISSION ERROR.", wbm.errorFormat)
        wbm.incrementRow(ws)
        return True
    except Exception as e:
        wbm.writeItem(ws, itemName)
        wbm.writeOutcome(ws, f"UNABLE TO READ. {e}", wbm.errorFormat)
        wbm.incrementRow(ws)
        return True

    if fileSize == 0:
        wbm.writeItem(ws, itemName)
        wbm.writeAuxiliary(ws, getOwnerCatch(dirAbsolute))
        wbm.writeOutcomeAndIncrement(ws, "", wbm.logFormat)
        return True

def deleteEmptyFilesModify(dirAbsolute:str, itemName:str, ws, _):
    """Glitch exists in that the current excel file will be considered empty.
    However, despite claiming so, the program does not actually delete it.'"""

    fullFilePath =  dirAbsolute + "\\" + itemName

    try:
        fileSize = os.path.getsize(addLongPathPrefix(fullFilePath))  # Bytes
    except PermissionError:
        wbm.writeItem(ws, itemName)
        wbm.writeOutcome(ws, "UNABLE TO READ. PERMISSION ERROR.", wbm.errorFormat)
        wbm.incrementRow(ws)
        return True
    except Exception as e:
        wbm.writeItem(ws, itemName)
        wbm.writeOutcome(ws, f"UNABLE TO READ. {e}", wbm.errorFormat)
        wbm.incrementRow(ws)
        return True

    # Stage for deletion
    if fileSize == 0:
        wbm.writeItem(ws, itemName)

        try:
            wbm.writeAuxiliary(ws, getOwnerCatch(dirAbsolute))
            os.remove(addLongPathPrefix(fullFilePath))
            wbm.writeOutcomeAndIncrement(ws, "", wbm.modifyFormat)
        except PermissionError:
            wbm.writeOutcomeAndIncrement(ws, "FAILED TO DELETE. PERMISSION ERROR.", wbm.errorFormat)
        except Exception as e:
            wbm.writeOutcomeAndIncrement(ws, f"FAILED TO DELETE. {e}", wbm.errorFormat)
        return True


def emptyFileFind(longFileAbsolute:str, dirAbsolute:str, itemName:str, ws):
    try:
        fileSize = os.path.getsize(addLongPathPrefix(dirAbsolute)+"\\"+itemName)
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
        wbm.writeAuxiliary(ws, getOwnerCatch(dirAbsolute))
        wbm.writeItemAndIncrement(ws, itemName, wbm.errorFormat)
        return True
