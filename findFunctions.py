from procedureFunctions import *


# Used by oldFileFind (1095 days == 3 years old)
DAYS_TOO_OLD = 1095


# Used by emptyDirectoryConcurrent
TOO_FEW_AMOUNT = 0

# Used by fileExtension
EXTENSION_COUNT = {}
EXTENSION_TOTAL_SIZE = {}
TOO_LARGE_SIZE_MB = 100

# Used by duplicateContent
HASH_AND_FILES = {}
# MMAP_THRESHOLD = 8 * 1024 * 1024  # 8MB to Bytes
hashFunc = hashlib.new("sha256")
hashFunc.update("".encode())
EMPTY_INPUT_HASH_CODE = hashFunc.hexdigest()



def setWorkbookManager(newManager: WorkbookManager):
    # Globally declare the WorkbookManager object
    global wbm
    wbm = newManager


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


def spaceFileFind(_1:str, _2:str, itemName:str, ws) -> bool:
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


def oldFileFind(longFileAbsolute:str, _:str, itemName:str, ws):
    try:
        fileDate = datetime.fromtimestamp(os.path.getatime(longFileAbsolute))
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


def fileExtensionConcurrent(longFileAbsolute:str, dirAbsolute:str, itemName:str, _):
    try: fileSize = os.path.getsize(longFileAbsolute) / 1000_000  # Bytes / 1000_000 = MBs
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
            ws.write_string(row, 0, extension, wbm.warningWeakFormat)
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


def duplicateContentHelper(longFileAbsolute:str):    
    hashFunc = hashlib.new("sha256")
    
    with open(longFileAbsolute, "rb") as file:
        while chunk := file.read(8192):
            hashFunc.update(chunk)

    return hashFunc.hexdigest()


def duplicateContentConcurrent(longFileAbsolute:str, dirAbsolute:str, itemName:str, ws):
    try:
        hashCode = duplicateContentHelper(longFileAbsolute)
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


def emptyFileFind(longFileAbsolute:str, dirAbsolute:str, itemName:str, ws):
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
        wbm.writeAuxiliary(ws, getOwnerCatch(dirAbsolute))
        wbm.writeItemAndIncrement(ws, itemName, wbm.errorFormat)
        return True
