from procedureFunctions import *


def setWorkbookManager(newManager: WorkbookManager):
    # Globally declare the WorkbookManager object
    global wbm
    wbm = newManager


def listAll(_1:str, _2:str, itemName:str, ws):
    wbm.incrementRowAndFileCount(ws)
    return (2, ExcelWritePackage(wbm.sheetRows[ws], wbm.ITEM_COL, itemName, ws))  # SPECIAL CASE

def listAllOwner(longFileAbsolute:str, _:str, itemName:str, ws):
    wbm.incrementRowAndFileCount(ws)
    row = wbm.sheetRows[ws]
    return (2, 
            ExcelWritePackage(row, wbm.ITEM_COL, itemName, ws), 
            ExcelWritePackage(row, wbm.AUXILIARY_COL, getOwnerCatch(longFileAbsolute), ws))  # SPECIAL CASE


def spaceFileFind(_1:str, _2:str, itemName:str, ws):
    if " " in itemName: 
        wbm.incrementRowAndFileCount(ws)
        return (True, ExcelWritePackage(wbm.sheetRows[ws], wbm.ITEM_COL, itemName, ws, wbm.errorFormat))
    return (False,)


def spaceFolderFind(dirAbsolute:str, dirFolders, dirFiles, ws):
    folderName = getDirectoryBaseName(dirAbsolute)

    if " " in folderName:
        wbm.incrementRowAndFileCount(ws)
        return (True, ExcelWritePackage(wbm.sheetRows[ws], wbm.ITEM_COL, folderName, ws, wbm.errorFormat))
    return (False,)


def overCharLimitFind(_:str, dirAbsolute:str, itemName:str, ws) -> bool:
    absoluteItemLength = len(dirAbsolute + "/" + itemName)

    if (absoluteItemLength > CHARACTER_LIMIT):
        wbm.incrementRowAndFileCount(ws)
        row = wbm.sheetRows[ws]
        return (True,
                ExcelWritePackage(row, wbm.ITEM_COL, itemName, ws, wbm.errorFormat),
                ExcelWritePackage(row, wbm.OUTCOME_COL, absoluteItemLength, ws))
    return (False,)


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
    rootName, _ = getRootNameAndExtension(itemName)
    badChars = badCharHelper(rootName)

    # if any bad characters were found
    if (badChars):
        wbm.incrementRowAndFileCount(ws)
        row = wbm.sheetRows[ws]
        return (True,
                ExcelWritePackage(row, wbm.ITEM_COL, itemName, ws, wbm.errorFormat), 
                ExcelWritePackage(row, wbm.OUTCOME_COL, "".join(badChars), ws))
    return (False,)

def badCharFolderFind(dirAbsolute:str, dirFolders, dirFiles, ws):
    folderName = getDirectoryBaseName(dirAbsolute)
    badChars = badCharHelper(folderName)

    if (badChars):
        wbm.incrementRowAndFileCount(ws)
        row = wbm.sheetRows[ws]
        return (True,
                ExcelWritePackage(row, wbm.ITEM_COL, folderName, ws, wbm.errorFormat),
                ExcelWritePackage(row, wbm.OUTCOME_COL, "".join(badChars), ws))
    return (False,)


def oldFileFindStart(arg, ws):
    writeOwnerHeader(arg, ws)

    global DAYS_LOWER_BOUND
    DAYS_LOWER_BOUND = arg[0]
    # # Double-checking that this value is usable. Dire consequences if not.
    # if (DAYS_LOWER_BOUND <= 0):
    #     raise Exception("OldFile's lower bound argument cannot be less than 1.")

    global DAYS_UPPER_BOUND
    if len(arg) == 2:
        DAYS_UPPER_BOUND = arg[1]
    else:
        DAYS_UPPER_BOUND = MAXSIZE


def oldFileFindHelper(longFileAbsolute):
    # NOTE: A file that is 23 hours and 59 minutes old is still considered 0 days old.
    
    # Get date of file. This *can* error virtue of the library functions, hence try/except.
    try: fileDate = datetime.fromtimestamp(os.path.getatime(longFileAbsolute))
    except: return -1
    
    fileDaysAgoLastAccessed = (TODAY - fileDate).days

    if (fileDaysAgoLastAccessed >= DAYS_LOWER_BOUND): return fileDaysAgoLastAccessed
    else: return 0

def oldFileFind(longFileAbsolute:str, _:str, itemName:str, ws):
    fileDaysAgoLastAccessed = oldFileFindHelper(longFileAbsolute)

    # Either it's actually 0 days old or the fileDate is not within the cutOffDate range. Either way, don't flag.
    # If it's greater than the upperbound, exit
    if fileDaysAgoLastAccessed == 0 or fileDaysAgoLastAccessed >= DAYS_UPPER_BOUND:
        return (False,)
    
    wbm.incrementRow(ws)
    row = wbm.sheetRows[ws]
    itemEwp = ExcelWritePackage(row, wbm.ITEM_COL, itemName, ws)
    
    if fileDaysAgoLastAccessed != -1:
        wbm.incrementFileCount(ws)
        return (True,
                itemEwp,
                ExcelWritePackage(row, wbm.AUXILIARY_COL, getOwnerCatch(longFileAbsolute), ws),
                ExcelWritePackage(row, wbm.OUTCOME_COL, fileDaysAgoLastAccessed, ws, wbm.errorFormat))
    
    # This executes if fromtimestamp() threw an exception
    else:
        return (2,
                itemEwp,
                ExcelWritePackage(row, wbm.OUTCOME_COL, f"UNABLE TO READ DATE.", ws, wbm.errorFormat))    


def oldFileFindRecommend(longFileAbsolute:str, _:str, itemName:str, ws):
    fileDaysAgoLastAccessed = oldFileFindHelper(longFileAbsolute)

    if (fileDaysAgoLastAccessed == 0): return (False,)

    wbm.incrementRow(ws)
    row = wbm.sheetRows[ws]
    itemEwp = ExcelWritePackage(row, wbm.ITEM_COL, itemName, ws)

    if fileDaysAgoLastAccessed != -1:
        # Unlike the other variants, this will still log (and flag) items above the UPPER_BOUND threshold
        if fileDaysAgoLastAccessed >= DAYS_UPPER_BOUND:
            dynamicFormat = wbm.warningStrongFormat
        else:
            dynamicFormat = wbm.warningWeakFormat

        wbm.incrementFileCount(ws)

        return (True,
                itemEwp,
                ExcelWritePackage(row, wbm.AUXILIARY_COL, getOwnerCatch(longFileAbsolute), ws),
                ExcelWritePackage(row, wbm.OUTCOME_COL, fileDaysAgoLastAccessed, ws, dynamicFormat))
    else:
        return (2,
                itemEwp,
                ExcelWritePackage(row, wbm.OUTCOME_COL, "UNABLE TO READ DATE.", ws, wbm.errorFormat))

###

def emptyDirectoryStart(arg, ws):
    global TOO_FEW_AMOUNT
    TOO_FEW_AMOUNT = arg[0]

def emptyDirectory(dirAbsolute:str, dirFolders, dirFiles, ws):
    folderName = getDirectoryBaseName(dirAbsolute)
    numFilesContained = len(dirFiles)

    if len(dirFolders) == 0 and numFilesContained <= TOO_FEW_AMOUNT:
        wbm.incrementRowAndFileCount(ws)
        return (True,
                ExcelWritePackage(wbm.sheetRows[ws], wbm.ITEM_COL, folderName, ws, wbm.errorFormat),
                ExcelWritePackage(wbm.sheetRows[ws], wbm.OUTCOME_COL, numFilesContained, ws, wbm.errorFormat))
    return (False,)

def emptyDirectoryRecommend(dirAbsolute:str, dirFolders, dirFiles, ws):
    if len(dirFolders) != 0: return (False,)

    folderName = getDirectoryBaseName(dirAbsolute)

    fileAmount = len(dirFiles)
    if fileAmount <= TOO_FEW_AMOUNT:
        
        ## HARD CODED RECOMMENDATION VALUES
        # Dynamic Format
        if fileAmount <= 1:
            dynamicFormat = wbm.warningStrongFormat
        elif fileAmount == 2:
            dynamicFormat = wbm.warningWeakFormat
        else:
            dynamicFormat = wbm.errorFormat

        # wbm.writeOutcomeAndIncrement(ws, fileAmount, dynamicFormat)
        wbm.incrementRowAndFileCount(ws)
        return (True,
                ExcelWritePackage(wbm.sheetRows[ws], wbm.ITEM_COL, folderName, ws),
                ExcelWritePackage(wbm.sheetRows[ws], wbm.OUTCOME_COL, fileAmount, ws, dynamicFormat))
    return (False,)
###


def fileExtensionStart(arg, ws):
    global EXTENSION_COUNT
    global EXTENSION_TOTAL_SIZE
    global TOO_LARGE_SIZE_MB
    global LOCK_FILE_EXTENSION
    
    EXTENSION_COUNT = defaultdict(int)
    EXTENSION_TOTAL_SIZE = defaultdict(int)
    TOO_LARGE_SIZE_MB = 100
    LOCK_FILE_EXTENSION = Lock()

def fileExtensionConcurrent(longFileAbsolute:str, dirAbsolute:str, itemName:str, _):
    try: fileSize = os.path.getsize(longFileAbsolute) / 1000_000  # Bytes / 1000_000 = MBs
    except: return (False,)

    _, extension = getRootNameAndExtension(itemName)

    with LOCK_FILE_EXTENSION:
        EXTENSION_COUNT[extension] += 1
        EXTENSION_TOTAL_SIZE[extension] += fileSize
    return (False,)

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
            wbm.incrementFileCount(ws, EXTENSION_COUNT[extension])
        else:
            ws.write_string(row, 0, extension)

        ws.write_number(row, 1, EXTENSION_COUNT[extension])
        ws.write_number(row, 2, averageSize)
        ws.write_number(row, 3, round(EXTENSION_TOTAL_SIZE[extension], 1))
        row += 1

    EXTENSION_COUNT.clear()
    EXTENSION_TOTAL_SIZE.clear()
    ws.freeze_panes(1, 0)


def duplicateContentStart(arg, ws):
    global HASH_AND_FILES
    global EMPTY_INPUT_HASH_CODE
    global LOCK_DUPLICATE_CONTENT
    global EXTENSIONS_TO_IGNORE

    HASH_AND_FILES = {}
    hashFunc = hashlib.new("sha256")
    hashFunc.update("".encode())
    EMPTY_INPUT_HASH_CODE = hashFunc.hexdigest()
    LOCK_DUPLICATE_CONTENT = Lock()

    EXTENSIONS_TO_IGNORE = {
        ".shx", ".shp", ".gbd", ".sbd", ".sbx", 
        ".dbf", ".qpj", ".atx", ".gdbtablx", ".gdbtable",
        ".freelist", ".horizon", ".gdbindexes", ".cpg", ".prj" 
    }


def duplicateContentHelper(longFileAbsolute:str):    
    hashFunc = hashlib.new("sha256")
    
    with open(longFileAbsolute, "rb") as file:
        while chunk := file.read(8192):
            hashFunc.update(chunk)

    return hashFunc.hexdigest()

def duplicateContentConcurrent(longFileAbsolute:str, dirAbsolute:str, itemName:str, ws):
    _, ext = getRootNameAndExtension(longFileAbsolute)
    if ext in EXTENSIONS_TO_IGNORE:
        return (False,)

    try:
        hashCode = duplicateContentHelper(longFileAbsolute)
    except Exception:  # FileNotFoundError, PermissionError, OSError, UnicodeDecodeError
        # Unlike other procedures, this won't print out the error; it'll just assume it's not a duplicate.
        return (False,)
    
    if (not hashCode or hashCode == EMPTY_INPUT_HASH_CODE):
        return (False,)
    
    with LOCK_DUPLICATE_CONTENT:
        if hashCode in HASH_AND_FILES:
            HASH_AND_FILES[hashCode][0].append(itemName)
            HASH_AND_FILES[hashCode][1].append(dirAbsolute)
            wbm.incrementFileCount(ws)
            return (3,)
        else:
            HASH_AND_FILES[hashCode] = ([itemName], [dirAbsolute])
            return (False,)

def duplicateContentPost(ws):
    ws.write(0, 0, "Separator", wbm.headerFormat)
    ws.write(0, 1, "Files", wbm.headerFormat)
    ws.write(0, 2, "Directories", wbm.headerFormat)
    ws.write(0, 3, "Owner", wbm.headerFormat)

    row = 1
    for hashCode in HASH_AND_FILES.keys():
        if (numOfFiles := len(HASH_AND_FILES[hashCode][0])) > 1:
            ws.write(row, 0, "------------", wbm.logFormat)

            for i in range(numOfFiles):
                ws.write(row, 1, HASH_AND_FILES[hashCode][0][i], wbm.errorFormat)
                ws.write(row, 2, HASH_AND_FILES[hashCode][1][i], wbm.dirFormat)
                ws.write(row, 3, getOwnerCatch(
                    joinDirToFileName(HASH_AND_FILES[hashCode][1][i], HASH_AND_FILES[hashCode][0][i])))
                row += 1
                
    HASH_AND_FILES.clear()
    ws.freeze_panes(1, 0)

def duplicateContentPostRecommend(ws):
    ws.write(0, 0, "Separator", wbm.headerFormat)
    ws.write(0, 1, "Files", wbm.headerFormat)
    ws.write(0, 2, "Directories", wbm.headerFormat)
    ws.write(0, 3, "Owner", wbm.headerFormat)

    row = 1
    folderAndItem = defaultdict(list)

    for hashCode in HASH_AND_FILES.keys():
        if (numOfFiles := len(HASH_AND_FILES[hashCode][0])) > 1:
            ws.write(row, 0, "------------", wbm.logFormat)

            # If there are 3 or more duplicates, highlight them all yellow at least. Otherwise, just flag as normal.
            if (numOfFiles >= 3):
                defaultItemFormat = wbm.warningWeakFormat
            else:
                defaultItemFormat = wbm.errorFormat

            # Sort this group of identical files with dirAbsolute as the key, and itemName as the values
            for i in range(numOfFiles):
                folderAndItem[HASH_AND_FILES[hashCode][1][i]].append(HASH_AND_FILES[hashCode][0][i])

            for dirAbsoluteKey in folderAndItem.keys():
                # If 2 or more files are identical AND reside in the same folder
                if (dirAbsoluteNumOfFiles := len(folderAndItem[dirAbsoluteKey])) > 1:

                    # Sort the list of items in descending order, ordered by number of characters
                    folderAndItem[dirAbsoluteKey].sort(key=len, reverse=True)

                    # Write the first one normally
                    ws.write(row, 1, folderAndItem[dirAbsoluteKey][0], defaultItemFormat)
                    ws.write(row, 2, dirAbsoluteKey, wbm.dirFormat)
                    ws.write(row, 3, getOwnerCatch(
                        joinDirToFileName(dirAbsoluteKey, folderAndItem[dirAbsoluteKey][0])))
                    row += 1

                    # Write the rest in strong warning format
                    for i in range(1, dirAbsoluteNumOfFiles):
                        ws.write(row, 1, folderAndItem[dirAbsoluteKey][i], wbm.warningStrongFormat)
                        ws.write(row, 2, dirAbsoluteKey, wbm.dirFormat)
                        ws.write(row, 3, getOwnerCatch(
                            joinDirToFileName(dirAbsoluteKey, folderAndItem[dirAbsoluteKey][i])))
                        row += 1
                
                # If this file is only duplicated once in this directory, just write it normally
                else:
                    ws.write(row, 1, folderAndItem[dirAbsoluteKey][0], defaultItemFormat)
                    ws.write(row, 2, dirAbsoluteKey, wbm.dirFormat)
                    ws.write(row, 3, getOwnerCatch(
                        joinDirToFileName(dirAbsoluteKey, folderAndItem[dirAbsoluteKey][0])))
                    row += 1

            folderAndItem.clear()
                
    HASH_AND_FILES.clear()
    ws.freeze_panes(1, 0)







def emptyFileFind(longFileAbsolute:str, dirAbsolute:str, itemName:str, ws):
    try:
        fileSize = os.path.getsize(longFileAbsolute)
    except PermissionError:
        wbm.incrementRow(ws)
        row = wbm.sheetRows[ws]
        return (2,
                ExcelWritePackage(row, wbm.ITEM_COL, itemName, ws),
                ExcelWritePackage(row, wbm.OUTCOME_COL, f"UNABLE TO READ. PERMISSION ERROR.", ws, wbm.errorFormat))
    except Exception as e:
        wbm.incrementRow(ws)
        row = wbm.sheetRows[ws]
        return (2,
                ExcelWritePackage(row, wbm.ITEM_COL, itemName, ws),
                ExcelWritePackage(row, wbm.OUTCOME_COL, f"UNABLE TO READ. {e}", ws, wbm.errorFormat))
    
    if fileSize == 0:
        wbm.incrementRowAndFileCount(ws)
        row = wbm.sheetRows[ws]
        return (True,
                ExcelWritePackage(row, wbm.AUXILIARY_COL, getOwnerCatch(longFileAbsolute), ws),
                ExcelWritePackage(row, wbm.ITEM_COL, itemName, ws, wbm.errorFormat))
    
    return (False,)

def emptyFileFindRecommend(longFileAbsolute:str, dirAbsolute:str, itemName:str, ws):
    try:
        fileSize = os.path.getsize(longFileAbsolute)
    except PermissionError:
        wbm.incrementRow(ws)
        row = wbm.sheetRows[ws]
        return (2,
                ExcelWritePackage(row, wbm.ITEM_COL, itemName, ws),
                ExcelWritePackage(row, wbm.OUTCOME_COL, f"UNABLE TO READ. PERMISSION ERROR.", ws, wbm.errorFormat))
    except Exception as e:
        wbm.incrementRow(ws)
        row = wbm.sheetRows[ws]
        return (2,
                ExcelWritePackage(row, wbm.ITEM_COL, itemName, ws),
                ExcelWritePackage(row, wbm.OUTCOME_COL, f"UNABLE TO READ. {e}", ws, wbm.errorFormat))
    
    if fileSize == 0:
        wbm.incrementRowAndFileCount(ws)
        row = wbm.sheetRows[ws]
        return (True,
                ExcelWritePackage(row, wbm.AUXILIARY_COL, getOwnerCatch(longFileAbsolute), ws),
                ExcelWritePackage(row, wbm.ITEM_COL, itemName, ws, wbm.warningStrongFormat),
                ExcelWritePackage(row, wbm.OUTCOME_COL, "", ws, wbm.warningStrongFormat))
    
    return (False,)