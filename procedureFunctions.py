from .procedureHelpers import *
import hashlib
from .getFileOwner import getOwnerCatch
from collections import defaultdict
from .ExcelWritePackage import ExcelWritePackage
from threading import Lock
from sys import maxsize as MAXSIZE
from datetime import datetime
import string
import networkx as nx
from difflib import SequenceMatcher
from networkx.algorithms.clique import find_cliques
from itertools import combinations


def setWorkbookManager(newManager: WorkbookManager):
    # Globally declare the WorkbookManager object
    global wbm
    wbm = newManager


def listAllBase(_1, _2, _3, itemName:str, ws):
    wbm.incrementRowAndFileCount(ws)
    return (2, ExcelWritePackage(wbm.sheetRows[ws], wbm.ITEM_COL, itemName, ws))  # SPECIAL CASE

def listAllOwnerBase(longFileAbsolute:str, _1, _2, itemName:str, ws):
    wbm.incrementRowAndFileCount(ws)
    row = wbm.sheetRows[ws]

    return (2, 
            ExcelWritePackage(row, wbm.ITEM_COL, itemName, ws), 
            ExcelWritePackage(row, wbm.AUXILIARY_COL, getOwnerCatch(longFileAbsolute), ws),
            ExcelWritePackage(row, wbm.AUXILIARY_COL+1, getLastModifiedDate(longFileAbsolute), ws)
            )  # SPECIAL CASE


def spaceFileFindBase(_1, _2, _3, itemName:str, ws):
    if " " in itemName: 
        wbm.incrementRowAndFileCount(ws)
        return (True, ExcelWritePackage(wbm.sheetRows[ws], wbm.ITEM_COL, itemName, ws, wbm.errorFormat))
    return (False,)


def spaceFolderFindBase(dirAbsolute:str, dirBasename, dirFolders, dirFiles, ws):
    if " " in dirBasename:
        wbm.incrementRowAndFileCount(ws)
        return (True, ExcelWritePackage(wbm.sheetRows[ws], wbm.ITEM_COL, dirBasename, ws, wbm.errorFormat))
    return (False,)


def exceedCharacterLimitStart(_, ws):
    writeDefaultHeaders(_, ws)

    global LAST_DIR
    LAST_DIR = None
    global LAST_DIR_INDEX
    LAST_DIR_INDEX = None

    global FILENAME_LENGTH_COL
    FILENAME_LENGTH_COL = wbm.OUTCOME_COL
    ws.write(0, FILENAME_LENGTH_COL, "Filename length", wbm.headerFormat)

    global FILEPATH_LENGTH_COL
    FILEPATH_LENGTH_COL = wbm.OUTCOME_COL+1
    ws.write(0, FILEPATH_LENGTH_COL, "Filepath length", wbm.headerFormat)

    #global NEW_FILENAME_COL
    #NEW_FILENAME_COL = wbm.OUTCOME_COL+2
    #ws.write(0, NEW_FILENAME_COL, "New filename", wbm.headerFormat)
    ws.write(0, wbm.OUTCOME_COL+2, "New filename", wbm.headerFormat)

    global NEW_FILENAME_LENGTH_COL
    NEW_FILENAME_LENGTH_COL = wbm.OUTCOME_COL+3
    ws.write(0, NEW_FILENAME_LENGTH_COL, "New filename length", wbm.headerFormat)

    global NEW_FILEPATH_LENGTH_COL
    NEW_FILEPATH_LENGTH_COL = wbm.OUTCOME_COL+4
    ws.write(0, NEW_FILEPATH_LENGTH_COL, "New filepath length", wbm.headerFormat)


def exceedCharacterLimitBase(_1, _2, dirAbsolute:str, itemName:str, ws) -> bool:
    global LAST_DIR
    global LAST_DIR_INDEX

    # This operation is kind of slow to do every time, but I can't think of anything else that would work
    if dirAbsolute != LAST_DIR:
        LAST_DIR = dirAbsolute
        LAST_DIR_INDEX = wbm.sheetRows[ws] +2

    # The slash separating dirAbsolute and itemName in the path name needs to be accounted for, hence +1
    # HARD CODED at 200
    if (len(dirAbsolute + itemName) +1) > 200:
        wbm.incrementRowAndFileCount(ws)
        row = wbm.sheetRows[ws]
        oneIndexedRow = row +1

        # HARD CODED LETTER COLUMNS. Could use Openpyxl's utils if want it chosen programmatically, but to be quite honest, this shouldn't ever change.
        return (True,
                ExcelWritePackage(row, wbm.ITEM_COL, itemName, ws, wbm.errorFormat),
                ExcelWritePackage(row, FILENAME_LENGTH_COL, f"=SUM(LEN(B{oneIndexedRow}))", ws),
                ExcelWritePackage(row, FILEPATH_LENGTH_COL, f"=SUM(LEN(A{LAST_DIR_INDEX}),LEN(B{oneIndexedRow}),1)", ws),
                ExcelWritePackage(row, NEW_FILENAME_LENGTH_COL, f"=SUM(LEN(E{oneIndexedRow}))" , ws),
                ExcelWritePackage(row, NEW_FILEPATH_LENGTH_COL, f"=SUM(LEN(A{LAST_DIR_INDEX}),LEN(E{oneIndexedRow}),1)", ws),
                )
    return (False,)


def badCharacterStart(_, ws):
    writeDefaultHeaders(_, ws)
    ws.write(0, wbm.OUTCOME_COL, "Characters", wbm.headerFormat)
    global PERMISSIBLE_CHARACTERS
    PERMISSIBLE_CHARACTERS = set(string.ascii_letters + string.digits + "- ")

def badCharacterHelper(s:str) -> set:
    badChars = set()

    for i in range(len(s)):
        if s[i] not in PERMISSIBLE_CHARACTERS:
            badChars.add(s[i])
        
        # double-dash error
        elif s[i:i+2] == "--":
            badChars.add(s[i])

    return badChars

def badCharacterFileFind(_1, _2, _3, itemName:str, ws) -> bool:
    rootName, _ = getRootNameAndExtension(itemName)
    badChars = badCharacterHelper(rootName)

    # if any bad characters were found
    if (badChars):
        wbm.incrementRowAndFileCount(ws)
        row = wbm.sheetRows[ws]
        return (True,
                ExcelWritePackage(row, wbm.ITEM_COL, itemName, ws, wbm.errorFormat), 
                ExcelWritePackage(row, wbm.OUTCOME_COL, "".join(badChars), ws))
    return (False,)

def badCharacterFolderFind(dirAbsolute:str, dirBasename, dirFolders, dirFiles, ws):
    badChars = badCharacterHelper(dirBasename)

    if (badChars):
        wbm.incrementRowAndFileCount(ws)
        row = wbm.sheetRows[ws]
        return (True,
                ExcelWritePackage(row, wbm.ITEM_COL, dirBasename, ws, wbm.errorFormat),
                ExcelWritePackage(row, wbm.OUTCOME_COL, "".join(badChars), ws))
    return (False,)


def oldFileStart(arg, ws):
    writeDefaultAndOwnerAndLastModifiedHeaders(arg, ws)
    ws.write(0, wbm.OUTCOME_COL, "# Days Last Accessed", wbm.headerFormat)

    global TODAY
    TODAY = datetime.now()

    global DAYS_LOWER_BOUND
    DAYS_LOWER_BOUND = arg[0]

    global DAYS_UPPER_BOUND
    if len(arg) == 2:
        DAYS_UPPER_BOUND = arg[1]
    else:
        DAYS_UPPER_BOUND = MAXSIZE

def oldFileHelper(longFileAbsolute):
    # NOTE: A file that is 23 hours and 59 minutes old is still considered 0 days old.
    
    # Get date of file. This *can* error virtue of the library functions, hence try/except.
    try: fileDate = datetime.fromtimestamp(os.path.getatime(longFileAbsolute))
    except: return -1
    
    fileDaysAgoLastAccessed = (TODAY - fileDate).days

    if (fileDaysAgoLastAccessed >= DAYS_LOWER_BOUND): return fileDaysAgoLastAccessed
    else: return 0

def oldFileBase(longFileAbsolute:str, _1, _2, itemName:str, ws):
    fileDaysAgoLastAccessed = oldFileHelper(longFileAbsolute)

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
                ExcelWritePackage(row, wbm.AUXILIARY_COL+1, getLastModifiedDate(longFileAbsolute), ws),
                ExcelWritePackage(row, wbm.OUTCOME_COL, fileDaysAgoLastAccessed, ws, wbm.errorFormat))
    
    # This executes if fromtimestamp() threw an exception
    else:
        return (2,
                itemEwp,
                ExcelWritePackage(row, wbm.OUTCOME_COL, f"UNABLE TO READ DATE.", ws, wbm.errorFormat))    

def oldFileRecommend(longFileAbsolute:str, _1, _2, itemName:str, ws):
    fileDaysAgoLastAccessed = oldFileHelper(longFileAbsolute)

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
                ExcelWritePackage(row, wbm.AUXILIARY_COL+1, getLastModifiedDate(longFileAbsolute), ws),
                ExcelWritePackage(row, wbm.OUTCOME_COL, fileDaysAgoLastAccessed, ws, dynamicFormat))
    else:
        return (2,
                itemEwp,
                ExcelWritePackage(row, wbm.OUTCOME_COL, "UNABLE TO READ DATE.", ws, wbm.errorFormat))


def emptyDirectoryStart(arg, ws):
    writeDefaultAndOwnerAndLastModifiedHeaders(arg, ws)
    ws.write(0, wbm.OUTCOME_COL, "# Files Contained", wbm.headerFormat)

    global TOO_FEW_AMOUNT
    TOO_FEW_AMOUNT = arg[0]

def emptyDirectoryBase(dirAbsolute:str, dirBasename, dirFolders, dirFiles, ws):
    dirBasename = getDirectoryBaseName(dirAbsolute)
    numFilesContained = len(dirFiles)

    if len(dirFolders) == 0 and numFilesContained <= TOO_FEW_AMOUNT:
        wbm.incrementRowAndFileCount(ws)
        return (True,
                ExcelWritePackage(wbm.sheetRows[ws], wbm.ITEM_COL, dirBasename, ws, wbm.errorFormat),
                ExcelWritePackage(wbm.sheetRows[ws], wbm.OUTCOME_COL, numFilesContained, ws, wbm.errorFormat))
    return (False,)

def emptyDirectoryRecommend(dirAbsolute:str, dirBasename, dirFolders, dirFiles, ws):
    if len(dirFolders) != 0: return (False,)

    dirBasename = getDirectoryBaseName(dirAbsolute)

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
                ExcelWritePackage(wbm.sheetRows[ws], wbm.ITEM_COL, dirBasename, ws),
                ExcelWritePackage(wbm.sheetRows[ws], wbm.OUTCOME_COL, fileAmount, ws, dynamicFormat))
    return (False,)


def fileExtensionSummaryStart(arg, _):
    global EXTENSION_COUNT
    global EXTENSION_TOTAL_SIZE
    global FILES_BY_EXTENSION
    global FILE_SIZE_LIMIT
    global LOCK_FILE_EXTENSION
    
    EXTENSION_COUNT = defaultdict(int)
    EXTENSION_TOTAL_SIZE = defaultdict(int)
    FILES_BY_EXTENSION = defaultdict(list)
    FILE_SIZE_LIMIT = arg[0]
    LOCK_FILE_EXTENSION = Lock()

def fileExtensionSummaryBase(longFileAbsolute:str, _, dirAbsolute, itemName:str, ws):
    try: fileSize = round(os.path.getsize(longFileAbsolute) / 1000_000, 1)  # Bytes / 1000_000 = MBs
    except: return (False,)

    _, extension = getRootNameAndExtension(itemName)
    if not extension: extension = "."

    with LOCK_FILE_EXTENSION:
        EXTENSION_COUNT[extension] += 1
        EXTENSION_TOTAL_SIZE[extension] += fileSize
        FILES_BY_EXTENSION[extension].append((dirAbsolute, itemName, fileSize))
    
    if fileSize >= FILE_SIZE_LIMIT:
        wbm.incrementFileCount(ws)
        return (3,)
    else:
        return (False,)

def fileExtensionSummaryPost(ws):
    ws.write(0, 0, "Extensions", wbm.headerFormat)
    ws.write(0, 1, "Count", wbm.headerFormat)
    ws.write(0, 2, "Avg Size (MB)", wbm.headerFormat)
    ws.write(0, 3, "Total Size (MB)", wbm.headerFormat)
    ws.set_column('A:D', 20)

    row = 1
    for extension in sorted(EXTENSION_COUNT):
        averageSize = round(EXTENSION_TOTAL_SIZE[extension] / EXTENSION_COUNT[extension], 1)

        ws.write_string(row, 0, extension)
        ws.write_number(row, 1, EXTENSION_COUNT[extension])
        ws.write_number(row, 2, averageSize)
        ws.write_number(row, 3, round(EXTENSION_TOTAL_SIZE[extension], 1))
        row += 1

    # NOTE: Don't do this stuff if not recommending.
    row += 1
    ws.write(row, 0, "Extension / Directory", wbm.headerFormat)
    ws.write(row, 1, "Item", wbm.headerFormat)
    ws.write(row, 2, "Size (MB)", wbm.headerFormat)
    for extension in sorted(FILES_BY_EXTENSION):
        row += 1
        ws.write(row, 0, extension, wbm.boldFormat)

        for pathPair in FILES_BY_EXTENSION[extension]:
            row += 1
            ws.write(row, 0, pathPair[0], wbm.dirFormat)
            ws.write(row, 1, pathPair[1])
            ws.write(row, 2, pathPair[2])
        row += 1

    EXTENSION_COUNT.clear()
    EXTENSION_TOTAL_SIZE.clear()
    FILES_BY_EXTENSION.clear()

def fileExtensionSummaryPostRecommend(ws):
    ws.write(0, 0, "Extensions", wbm.headerFormat)
    ws.write(0, 1, "Count", wbm.headerFormat)
    ws.write(0, 2, "Avg Size (MB)", wbm.headerFormat)
    ws.write(0, 3, "Total Size (MB)", wbm.headerFormat)
    ws.set_column('A:D', 20)

    errorExtensions = set()
    nonErrorExtensionsWithAtLeastOneErrorFile = set()

    row = 1
    for extension in sorted(EXTENSION_COUNT):
        averageSize = round(EXTENSION_TOTAL_SIZE[extension] / EXTENSION_COUNT[extension], 1)

        # only count as an error if the average size is above a threshold
        # NOTE: Writes all extensions indiscriminately in the summary
        if (averageSize >= FILE_SIZE_LIMIT):
            errorExtensions.add(extension)
            ws.write_string(row, 0, extension, wbm.warningWeakFormat)
            ws.write_number(row, 2, averageSize, wbm.warningWeakFormat)
        else:
            ws.write_string(row, 0, extension)
            ws.write_number(row, 2, averageSize)

            # Check if this extension has at least one error file.
            # TODO: Do this another way because this is pretty rudimentary/naive. Could use a data structure instead.
            for pathPair in FILES_BY_EXTENSION[extension]:
                if pathPair[2] >= FILE_SIZE_LIMIT:
                    nonErrorExtensionsWithAtLeastOneErrorFile.add(extension)
                    break

        ws.write_number(row, 1, EXTENSION_COUNT[extension])
        ws.write_number(row, 3, round(EXTENSION_TOTAL_SIZE[extension], 1))
        row += 1

    row += 1
    ws.write(row, 0, "Extension / Directory", wbm.headerFormat)
    ws.write(row, 1, "Item", wbm.headerFormat)
    ws.write(row, 2, "Size (MB)", wbm.headerFormat)
    
    # Do the errorExtensions first
    for extension in sorted(errorExtensions):
        row += 1
        ws.write(row, 0, extension, wbm.warningWeakFormat)

        for pathPair in FILES_BY_EXTENSION[extension]:
            row += 1
            ws.write(row, 0, pathPair[0], wbm.dirFormat)
            ws.write(row, 1, pathPair[1])
            if (fileSize := pathPair[2]) >= FILE_SIZE_LIMIT:
                ws.write(row, 2, fileSize, wbm.warningWeakFormat)
            else:
                ws.write(row, 2, fileSize)
        row += 1
    
    # Then do the rest of the applicable extensions
    row += 1
    for extension in sorted(nonErrorExtensionsWithAtLeastOneErrorFile): 
        ws.write(row, 0, extension)
        row += 1

        for pathPair in FILES_BY_EXTENSION[extension]:
            if (fileSize := pathPair[2]) >= FILE_SIZE_LIMIT:
                ws.write(row, 0, pathPair[0], wbm.dirFormat)
                ws.write(row, 1, pathPair[1])
                ws.write(row, 2, fileSize, wbm.warningWeakFormat)
                row += 1 
        row += 1

    EXTENSION_COUNT.clear()
    EXTENSION_TOTAL_SIZE.clear()
    FILES_BY_EXTENSION.clear()


def identicalFileStart(arg, ws):
    global HASH_AND_FILES
    global EMPTY_INPUT_HASH_CODE
    global LOCK_DUPLICATE_CONTENT
    global EXTENSIONS_TO_OMIT

    HASH_AND_FILES = {}
    hashFunc = hashlib.new("sha256")
    hashFunc.update("".encode())
    EMPTY_INPUT_HASH_CODE = hashFunc.hexdigest()
    LOCK_DUPLICATE_CONTENT = Lock()

def identicalFileHelper(longFileAbsolute:str):    
    hashFunc = hashlib.new("sha256")
    
    with open(longFileAbsolute, "rb") as file:
        while chunk := file.read(8192):
            hashFunc.update(chunk)

    return hashFunc.hexdigest()

def identicalFileBase(longFileAbsolute:str, _, dirAbsolute:str, itemName:str, ws):
    try:
        hashCode = identicalFileHelper(longFileAbsolute)
    except Exception:  # FileNotFoundError, PermissionError, OSError, UnicodeDecodeError
        # Unlike other procedures, this won't print out the error; it'll just assume it's not a duplicate.
        return (False,)
    
    if (not hashCode or hashCode == EMPTY_INPUT_HASH_CODE):
        return (False,)
    
    # 0 -> itemName list
    # 1 -> dirAbsolute list
    # 2 -> longFileAbsolute list
    with LOCK_DUPLICATE_CONTENT:
        if hashCode in HASH_AND_FILES:
            HASH_AND_FILES[hashCode][0].append(itemName)
            HASH_AND_FILES[hashCode][1].append(dirAbsolute)
            HASH_AND_FILES[hashCode][2].append(longFileAbsolute)
            wbm.incrementFileCount(ws)
            return (3,)
        else:
            HASH_AND_FILES[hashCode] = ([itemName], [dirAbsolute], [longFileAbsolute])
            return (False,)

def identicalFilePost(ws):
    ws.write(0, 0, "Separator", wbm.headerFormat)
    ws.write(0, 1, "File", wbm.headerFormat)
    ws.write(0, 2, "Directory", wbm.headerFormat)
    ws.write(0, 3, "Owner", wbm.headerFormat)
    ws.write(0, 4, "Last Modified", wbm.headerFormat)

    row = 1
    for hashCode in HASH_AND_FILES.keys():
        if (numOfFiles := len(HASH_AND_FILES[hashCode][0])) > 1:
            ws.write(row, 0, GROUP_SEPARATOR, wbm.separatorFormat)

            for i in range(numOfFiles):
                longFileAbsolute = HASH_AND_FILES[hashCode][2][i]
                ws.write(row, 1, HASH_AND_FILES[hashCode][0][i], wbm.errorFormat)
                ws.write(row, 2, HASH_AND_FILES[hashCode][1][i], wbm.dirFormat)
                ws.write(row, 3, getOwnerCatch(longFileAbsolute))
                ws.write(row, 4, getLastModifiedDate(longFileAbsolute))
                row += 1
                
    HASH_AND_FILES.clear()
    ws.freeze_panes(1, 0)
    ws.autofit()

def identicalFilePostRecommend(ws):
    ws.write(0, 0, "Separator", wbm.headerFormat)
    ws.write(0, 1, "Staged for Deletion", wbm.headerFormat)
    ws.write(0, 2, "File", wbm.headerFormat)
    ws.write(0, 3, "Directory", wbm.headerFormat)
    ws.write(0, 4, "Owner", wbm.headerFormat)
    ws.write(0, 5, "Last Modified", wbm.headerFormat)

    # hide "Delete" column
    ws.set_column("B:B", None, None, {"hidden": True})

    ### purple highlights
    dirToHashGraph = nx.Graph() # bipartite graph
    flaggedHashes = set()
    duplicatedHashes = set()
    hashOccurrences = defaultdict(int)
    # dirAbsoluteNodeList = []
    for hashCode in HASH_AND_FILES:
        if (len(HASH_AND_FILES[hashCode][0])) <= 1:
            continue

        duplicatedHashes.add(hashCode)
        
        for dirAbsolute in HASH_AND_FILES[hashCode][1]:
            dirToHashGraph.add_edge(hashCode, dirAbsolute)
            # dirAbsoluteNodeList.append(dirAbsolute)

    for hashCodeNode in duplicatedHashes:
        # If this hashCode has already been flagged, don't look into it again
        if hashCodeNode in flaggedHashes: continue

        hashCodeNodeNeighborDirs = list(dirToHashGraph.neighbors(hashCodeNode))

        # If this hashCode is only found within 2 directories, skip it
        if len(hashCodeNodeNeighborDirs) < 3: continue

        # Loop through the nC3 combinations
        for setOfDirs in combinations(hashCodeNodeNeighborDirs, 3):
            hashOccurrences.clear()

            # Count how many times each hash appears across the set of 3 directories
            for dir in setOfDirs:
                for dirNodeNeighbourHashCode in dirToHashGraph.neighbors(dir):
                    hashOccurrences[dirNodeNeighbourHashCode] += 1
            
            # Record which hashes appear 3 times
            for ho in hashOccurrences:
                if (hashOccurrences[ho]) >= 3: # should be ==, probably
                    flaggedHashes.add(ho)
            
        #for dirAbsoluteNode in list(dirToHashGraph.neighbors(hashCodeNode)):
            # Loop through the files of the 3 or more neighbors
    # for hashCodeNode in HASH_AND_FILES:
    #     print(hashCodeNode, len(list(dirToHashGraph.neighbors(hashCodeNode))))
    #     print()
    ###

    row = 1
    folderAndItem = defaultdict(list)

    for hashCode in duplicatedHashes:
        numOfFiles = len(HASH_AND_FILES[hashCode][0])
        ws.write(row, 0, GROUP_SEPARATOR, wbm.separatorFormat)

        # Prioritize the seriesOfIdenticalFiles flag. Then the numFiles>=3 flag. Then regular flag.
        if (hashCode in flaggedHashes):
            defaultItemFormat = wbm.warningMiddlingFormat
        elif (numOfFiles >= 3):
            defaultItemFormat = wbm.warningWeakFormat
        else:
            defaultItemFormat = wbm.errorFormat

        # Sort this group of identical files with longFileAbsolute as the key, and itemName as the values
        for i in range(numOfFiles):
            folderAndItem[HASH_AND_FILES[hashCode][1][i]].append(
                (HASH_AND_FILES[hashCode][0][i], HASH_AND_FILES[hashCode][2][i])
                )

        for dirAbsoluteKey in folderAndItem.keys():
            # If 2 or more files are identical AND reside in the same folder
            if (dirAbsoluteNumOfFiles := len(folderAndItem[dirAbsoluteKey])) > 1:

                # Sort the list of items in descending order, ordered by number of characters
                folderAndItem[dirAbsoluteKey].sort(key=len, reverse=True)

                # Write the first one normally
                ws.write(row, 2, folderAndItem[dirAbsoluteKey][0][0], defaultItemFormat)
                ws.write(row, 3, dirAbsoluteKey, wbm.dirFormat)
                ws.write(row, 4, getOwnerCatch(folderAndItem[dirAbsoluteKey][0][1]))
                ws.write(row, 5, getLastModifiedDate(folderAndItem[dirAbsoluteKey][0][1]))
                row += 1

                # Write the rest in strong warning format
                for i in range(1, dirAbsoluteNumOfFiles):
                    ws.write(row, 2, folderAndItem[dirAbsoluteKey][i][0], wbm.warningStrongFormat)
                    ws.write(row, 3, dirAbsoluteKey, wbm.dirFormat)
                    ws.write(row, 4, getOwnerCatch(folderAndItem[dirAbsoluteKey][i][1]))
                    ws.write(row, 5, getLastModifiedDate(folderAndItem[dirAbsoluteKey][i][1]))
                    row += 1
            
            # If this file is only duplicated once in this directory, just write it normally
            else:
                ws.write(row, 2, folderAndItem[dirAbsoluteKey][0][0], defaultItemFormat)
                ws.write(row, 3, dirAbsoluteKey, wbm.dirFormat)
                ws.write(row, 4, getOwnerCatch(folderAndItem[dirAbsoluteKey][0][1]))
                ws.write(row, 5, getLastModifiedDate(folderAndItem[dirAbsoluteKey][0][1]))
                row += 1

        folderAndItem.clear()
                
    HASH_AND_FILES.clear()
    ws.freeze_panes(1, 0)
    ws.autofit()


def emptyFileFindBase(longFileAbsolute:str, _1, _2, itemName:str, ws):
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
                ExcelWritePackage(row, wbm.AUXILIARY_COL+1, getLastModifiedDate(longFileAbsolute), ws),
                ExcelWritePackage(row, wbm.ITEM_COL, itemName, ws, wbm.errorFormat))
    
    return (False,)

def emptyFileFindRecommend(longFileAbsolute:str, _1, _2, itemName:str, ws):
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
                ExcelWritePackage(row, wbm.AUXILIARY_COL+1, getLastModifiedDate(longFileAbsolute), ws),
                ExcelWritePackage(row, wbm.ITEM_COL, itemName, ws, wbm.warningStrongFormat),
                ExcelWritePackage(row, wbm.OUTCOME_COL, "", ws, wbm.warningStrongFormat))
    
    return (False,)




## FIX FUNCTIONS#########################################################################################################

# Used by spaceFolderFixModify or searchAndReplaceFolderModify (not logs)
# Only run modify function can be run at a time, so this is okay
def fixFolderModifyStart(_1, _2):
    global FOLDER_RENAMES
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


def spaceFolderFixStart(_, ws):
    writeDefaultHeaders(_, ws)
    fixFolderModifyStart(_, ws)

def spaceFolderFixHelper(oldFolderName) -> str:
    if (" " not in oldFolderName) and ("--" not in oldFolderName):
        return
    return "-".join(oldFolderName.replace("-", " ").split())

def spaceFolderFixBase(dirAbsolute, oldDirBasename, dirFolders, dirFiles, ws):
    newDirBasename = spaceFolderFixHelper(oldDirBasename)

    if (not newDirBasename): return (False,)
    
    wbm.incrementRowAndFileCount(ws)
    row = wbm.sheetRows[ws]
    
    return (True,
            ExcelWritePackage(row, wbm.DIR_COL, dirAbsolute, ws, wbm.dirFormat),
            ExcelWritePackage(row, wbm.ITEM_COL, oldDirBasename, ws, wbm.errorFormat),
            ExcelWritePackage(row, wbm.OUTCOME_COL, newDirBasename, ws, wbm.logFormat))

def spaceFolderFixModify(dirAbsolute, oldDirBasename, dirFolders, dirFiles, ws):
    newDirBasename = spaceFolderFixHelper(oldDirBasename)

    if (not newDirBasename): return (False,)

    directoryOfFolder = getDirectoryDirName(dirAbsolute)
    FOLDER_RENAMES.append([directoryOfFolder, oldDirBasename, newDirBasename])
    return (3,)


def spaceFileFixHelper(oldItemName) -> str:
    # Also fixes double dashes, even if no space is present
    if (" " not in oldItemName) and ("--" not in oldItemName):
        return
    
    rootName, extension = getRootNameAndExtension(oldItemName)
    newItemNameSansExt = "-".join(rootName.replace("-", " ").split())
    return newItemNameSansExt + extension

def spaceFileFixBase(_1:str, _2:str, _3:str, oldItemName:str, ws):
    newItemName = spaceFileFixHelper(oldItemName)
    if (not newItemName): return (False,)

    wbm.incrementRowAndFileCount(ws)
    row = wbm.sheetRows[ws]
    return (True,
            ExcelWritePackage(row, wbm.ITEM_COL, oldItemName, ws),
            ExcelWritePackage(row, wbm.OUTCOME_COL, newItemName, ws, wbm.logFormat)) 

def spaceFileFixModify(longFileAbsolute:str, longDirAbsolute:str, dirAbsolute:str, oldItemName:str, ws):
    newItemName = spaceFileFixHelper(oldItemName)
    if (not newItemName): return (False,)

    wbm.incrementRowAndFileCount(ws)
    row = wbm.sheetRows[ws]
    itemEwp = ExcelWritePackage(row, wbm.ITEM_COL, oldItemName, ws)

    # Log newItemName and rename file
    try:
        os.rename(longFileAbsolute, joinDirToFileName(longDirAbsolute, newItemName))
        outcomeEwp = ExcelWritePackage(row, wbm.OUTCOME_COL, newItemName, ws, wbm.modifyFormat)
    except PermissionError:
        outcomeEwp = ExcelWritePackage(row, wbm.OUTCOME_COL, "MODIFICATION FAILED. PERMISSION ERROR.", ws, wbm.errorFormat)
    except Exception as e:
        outcomeEwp = ExcelWritePackage(row, wbm.OUTCOME_COL, f"MODIFICATION FAILED. {e}", ws, wbm.errorFormat)
    return (True,
            itemEwp,
            outcomeEwp)


def searchAndReplaceFolderStart(arg, ws):
    writeDefaultHeaders(arg, ws)
    fixFolderModifyStart(arg, ws)
    global PAIRS_OF_FOLDER_REPLACEMENTS
    PAIRS_OF_FOLDER_REPLACEMENTS = arg

def searchAndReplaceFolderHelper(oldFolderName:str):
    newFolderName = oldFolderName
    for argPair in PAIRS_OF_FOLDER_REPLACEMENTS:
        toBeReplaced, replacer = argPair
        newFolderName = newFolderName.replace(toBeReplaced, replacer)

    if (oldFolderName == newFolderName): return
    return newFolderName

def searchAndReplaceFolderBase(dirAbsolute, oldDirBasename, dirFolders, dirFiles, ws):
    if not (newFolderName := searchAndReplaceFolderHelper(oldDirBasename)):
        return (False,)
    
    wbm.incrementRowAndFileCount(ws)
    row = wbm.sheetRows[ws]

    return (True,
            ExcelWritePackage(row, wbm.DIR_COL, dirAbsolute, ws, wbm.dirFormat),
            ExcelWritePackage(row, wbm.ITEM_COL, oldDirBasename, ws, wbm.errorFormat),
            ExcelWritePackage(row, wbm.OUTCOME_COL, newFolderName, ws, wbm.logFormat)
            )

def searchAndReplaceFolderModify(dirAbsolute, oldDirBasename, dirFolders, dirFiles, ws):
    if not (newFolderName := searchAndReplaceFolderHelper(oldDirBasename)):
        return (False,)
    
    directoryOfFolder = getDirectoryDirName(dirAbsolute)
    FOLDER_RENAMES.append([directoryOfFolder, oldDirBasename, newFolderName])
    return (3,)


def searchAndReplaceFileStart(arg, ws):
    writeDefaultHeaders(arg, ws)

    global PAIRS_OF_FILE_REPLACEMENTS
    PAIRS_OF_FILE_REPLACEMENTS = arg

def searchAndReplaceFileHelper(oldItemName:str):
    oldItemNameSansExt, extension = getRootNameAndExtension(oldItemName)
    
    # Order of argument pairs given matters.
    newItemNameSansExt = oldItemNameSansExt
    for toBeReplaced, replacer in PAIRS_OF_FILE_REPLACEMENTS:
        newItemNameSansExt = newItemNameSansExt.replace(toBeReplaced, replacer)

    if (oldItemNameSansExt == newItemNameSansExt): return
    return newItemNameSansExt + extension

def searchAndReplaceFileBase(longFileAbsolute:str, longDirAbsolute:str, _:str, oldItemName:str, ws):
    if not (newItemName := searchAndReplaceFileHelper(oldItemName)): return (False,)

    wbm.incrementRowAndFileCount(ws)
    row = wbm.sheetRows[ws]
    return (True,
            ExcelWritePackage(row, wbm.ITEM_COL, oldItemName, ws, wbm.errorFormat),
            ExcelWritePackage(row, wbm.OUTCOME_COL, newItemName, ws, wbm.logFormat))

def searchAndReplaceFileModify(longFileAbsolute:str, longDirAbsolute:str, dirAbsolute:str, oldItemName:str, ws):
    if not (newItemName := searchAndReplaceFileHelper(oldItemName)): return (False,)

    wbm.incrementRowAndFileCount(ws)
    row = wbm.sheetRows[ws]
    itemEwp = ExcelWritePackage(row, wbm.ITEM_COL, oldItemName, ws, wbm.errorFormat)

    try:
        os.rename(longFileAbsolute, joinDirToFileName(longDirAbsolute, newItemName))
        outcomeEwp = ExcelWritePackage(row, wbm.OUTCOME_COL, newItemName, ws, wbm.modifyFormat)
    except PermissionError:
        outcomeEwp = ExcelWritePackage(row, wbm.OUTCOME_COL, "MODIFICATION FAILED. PERMISSION ERROR.", ws, wbm.errorFormat)
    except Exception as e:
        outcomeEwp = ExcelWritePackage(row, wbm.OUTCOME_COL, f"MODIFICATION FAILED. {e}", ws, wbm.errorFormat)
    return (True,
            itemEwp,
            outcomeEwp)


def deleteEmptyFileStart(_, ws):
    writeDefaultAndOwnerAndLastModifiedHeaders(_, ws)
    ws.write(0, wbm.OUTCOME_COL, "Staged for Deletion", wbm.headerFormat)

def deleteEmptyFileBase(longFileAbsolute:str, longDirAbsolute:str, dirAbsolute:str, itemName:str, ws):
    try:
        fileSize = os.path.getsize(longFileAbsolute) # Bytes
    except PermissionError:
        wbm.incrementRow(ws)
        row = wbm.sheetRows[ws]
        return (2,
                ExcelWritePackage(row, wbm.ITEM_COL, itemName, ws),
                ExcelWritePackage(row, wbm.OUTCOME_COL, "UNABLE TO READ. PERMISSION ERROR.", ws, wbm.errorFormat))
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
                ExcelWritePackage(row, wbm.ITEM_COL, itemName, ws),
                ExcelWritePackage(row, wbm.AUXILIARY_COL, getOwnerCatch(longFileAbsolute), ws),
                ExcelWritePackage(row, wbm.AUXILIARY_COL+1, getLastModifiedDate(longFileAbsolute), ws),
                ExcelWritePackage(row, wbm.OUTCOME_COL, "", ws, wbm.logFormat))
    return (False,)

def deleteEmptyFileRecommend(longFileAbsolute:str, longDirAbsolute:str, dirAbsolute:str, itemName:str, ws):
    try:
        fileSize = os.path.getsize(longFileAbsolute)
    except PermissionError:
        wbm.incrementRow(ws)
        row = wbm.sheetRows[ws]
        return (2,
                ExcelWritePackage(row, wbm.ITEM_COL, itemName, ws),
                ExcelWritePackage(row, wbm.OUTCOME_COL, "UNABLE TO READ. PERMISSION ERROR.", ws, wbm.errorFormat))
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
                ExcelWritePackage(row, wbm.ITEM_COL, itemName, ws, wbm.warningStrongFormat),
                ExcelWritePackage(row, wbm.AUXILIARY_COL, getOwnerCatch(longFileAbsolute), ws),
                ExcelWritePackage(row, wbm.AUXILIARY_COL+1, getLastModifiedDate(longFileAbsolute), ws),
                ExcelWritePackage(row, wbm.OUTCOME_COL, "", ws, wbm.warningStrongFormat))
    return (False,)


def deleteEmptyFileModify(longFileAbsolute:str, longDirAbsolute:str, dirAbsolute:str, itemName:str, ws):
    """Glitch exists in that the current excel file will be considered empty.
    However, despite claiming so, the program does not actually delete it.'"""

    try:
        fileSize = os.path.getsize(longFileAbsolute)
    except PermissionError:
        wbm.incrementRow(ws)
        row = wbm.sheetRows[ws]
        return (2,
                ExcelWritePackage(row, wbm.ITEM_COL, itemName, ws),
                ExcelWritePackage(row, wbm.OUTCOME_COL, "UNABLE TO READ. PERMISSION ERROR.", ws, wbm.errorFormat))
    
    except Exception as e:
        wbm.incrementRow(ws)
        row = wbm.sheetRows[ws]
        return (2,
                ExcelWritePackage(row, wbm.ITEM_COL, itemName, ws),
                ExcelWritePackage(row, wbm.OUTCOME_COL, f"UNABLE TO READ. {e}", ws, wbm.errorFormat))


    # Stage for deletion
    if fileSize == 0:
        wbm.incrementRowAndFileCount(ws)
        row = wbm.sheetRows[ws]
        itemEwp = ExcelWritePackage(row, wbm.ITEM_COL, itemName, ws)
        ownerEwp = ExcelWritePackage(row, wbm.AUXILIARY_COL, getOwnerCatch(longFileAbsolute), ws)
        lastModifiedEwp = ExcelWritePackage(row, wbm.AUXILIARY_COL+1, getLastModifiedDate(longFileAbsolute), ws)

        try:
            os.remove(longFileAbsolute)
            outcomeEwp = ExcelWritePackage(row, wbm.OUTCOME_COL, "", ws, wbm.modifyFormat)
        except PermissionError:
            outcomeEwp = ExcelWritePackage(row, wbm.OUTCOME_COL, "FAILED TO DELETE. PERMISSION ERROR.", ws, wbm.errorFormat)
        except Exception as e:
            outcomeEwp = ExcelWritePackage(row, wbm.OUTCOME_COL, f"FAILED TO DELETE. {e}", ws, wbm.errorFormat)
        return (True,
                itemEwp,
                ownerEwp,
                lastModifiedEwp,
                outcomeEwp)   
    return (False,)


def multipleVersionStart(arg, _):
    global VG
    global NODE_TUPLES
    global FILTERED_NAMES
    global NUM_NODES
    global SIMILARITY_THRESHOLD
    VG = nx.Graph()
    NODE_TUPLES = []
    FILTERED_NAMES = []
    NUM_NODES = 0
    SIMILARITY_THRESHOLD = arg[0] / 100
    
def multipleVersionSimilarityMeasureHelper(s, t):  # similarity measurer
    return SequenceMatcher(None, s, t).ratio()

def multipleVersionFilterHelper(s):
    # Filters:
    # remove extension
    # enable case-insensitivty via lower casing ubiquitously
    return getRootNameAndExtension(s)[0].lower()

def multipleVersionBase(longFileAbsolute:str, longDirAbsolute:str, dirAbsolute:str, itemName:str, ws):
    global NUM_NODES
    newNodeTuple = (itemName, dirAbsolute, longFileAbsolute)
    filteredItemName = multipleVersionFilterHelper(itemName)

    for i in range(NUM_NODES):
        if multipleVersionSimilarityMeasureHelper(FILTERED_NAMES[i], filteredItemName) >= SIMILARITY_THRESHOLD:  # our threshold
            VG.add_edge(NODE_TUPLES[i], newNodeTuple)
    
    FILTERED_NAMES.append(filteredItemName)
    NODE_TUPLES.append(newNodeTuple)
    VG.add_node(newNodeTuple)  # only executes if no edges were added
    NUM_NODES += 1
    return (False,)
    
def multipleVersionPost(ws):
    global NODE_TUPLES
    ws.write(0, 0, "Separator", wbm.headerFormat)
    ws.write(0, 1, "Item", wbm.headerFormat)
    ws.write(0, 2, "Directory", wbm.headerFormat)
    ws.write(0, 3, "Owner", wbm.headerFormat)
    ws.write(0, 4, "Last Modified", wbm.headerFormat)
    
    row = 1
    cliques = find_cliques(VG)
    cliques = sorted(list(cliques), key=lambda clique: clique[0][1].lower())
    
    for clique in cliques:
        if len(clique) <= 1: continue

        wbm.incrementFileCount(ws)
        ws.write(row, 0, GROUP_SEPARATOR, wbm.separatorFormat)

        for nodeTuple in clique: #clique[1:]
            ws.write(row, 1, nodeTuple[0])
            ws.write(row, 2, nodeTuple[1], wbm.dirFormat)
            ws.write(row, 3, getOwnerCatch(nodeTuple[2]))
            ws.write(row, 4, getLastModifiedDate(nodeTuple[2]))
            row += 1

    NODE_TUPLES.clear()
    FILTERED_NAMES.clear()
    ws.freeze_panes(1, 0)
    ws.autofit()

def multipleVersionPostRecommend(ws):
    global NODE_TUPLES
    ws.write(0, 0, "Separator", wbm.headerFormat)
    ws.write(0, 1, "Item", wbm.headerFormat)
    ws.write(0, 2, "Directory", wbm.headerFormat)
    ws.write(0, 3, "Owner", wbm.headerFormat)
    ws.write(0, 4, "Last Modified", wbm.headerFormat)
    
    row = 1
    cliques = find_cliques(VG)
    cliques = sorted(list(cliques), key=lambda clique: clique[0][1].lower())
    
    for clique in cliques:
        if (cliqueLength := len(clique)) <= 1: continue

        wbm.incrementFileCount(ws)
        ws.write(row, 0, GROUP_SEPARATOR, wbm.separatorFormat)

        mostRecentLMDate = getLastModifiedDate(clique[0][2])
        lastModifiedDates = [mostRecentLMDate,]
        mostRecentDateIndexes = {0} # Since more than one lmDate can be the same date and the most recent simultaneously
        for i in range(1, cliqueLength):
            lmDate = getLastModifiedDate(clique[i][2])
            lastModifiedDates.append(lmDate)
            
            # Compare ISO date strings lexicographically, which works because they're in the format YYYY-MM-DD
            if lmDate > mostRecentLMDate:
                mostRecentLMDate = lmDate
                mostRecentDateIndexes = {i} # Start the set anew
            elif lmDate == mostRecentLMDate:
                mostRecentDateIndexes.add(i) # Add to the current set

        for i in range(cliqueLength): #clique[1:]
            ws.write(row, 1, clique[i][0], wbm.warningWeakFormat if i not in mostRecentDateIndexes else wbm.errorFormat)
            ws.write(row, 2, clique[i][1], wbm.dirFormat)
            ws.write(row, 3, getOwnerCatch(clique[i][2]))
            ws.write(row, 4, getLastModifiedDate(clique[i][2]))
            row += 1

    NODE_TUPLES.clear()
    FILTERED_NAMES.clear()
    ws.freeze_panes(1, 0)
    ws.autofit()
