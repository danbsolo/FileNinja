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
    oldFolderName = getDirectoryBaseName(dirAbsolute)
    newFolderName = spaceFolderFixHelper(oldFolderName)

    if (not newFolderName): return (False,)
    
    wbm.incrementRowAndFileCount(ws)
    row = wbm.sheetRows[ws]
    #wbm.writeDir(ws, dirAbsolute, wbm.dirFormat)
    #wbm.writeItem(ws, oldFolderName, wbm.errorFormat)
    #wbm.writeOutcomeAndIncrement(ws, newFolderName, wbm.logFormat)
    
    return (True,
            ExcelWritePackage(row, wbm.DIR_COL, dirAbsolute, ws, wbm.dirFormat),
            ExcelWritePackage(row, wbm.ITEM_COL, oldFolderName, ws, wbm.errorFormat),
            ExcelWritePackage(row, wbm.OUTCOME_COL, newFolderName, ws, wbm.logFormat))


def spaceFolderFixModify(dirAbsolute, dirFolders, dirFiles, ws, arg):
    oldFolderName = getDirectoryBaseName(dirAbsolute)
    newFolderName = spaceFolderFixHelper(oldFolderName)

    if (not newFolderName): return False

    directoryOfFolder = getDirectoryDirName(dirAbsolute)
    FOLDER_RENAMES.append([directoryOfFolder, oldFolderName, newFolderName])
    return (3,)


def spaceFileFixHelper(oldItemName) -> str:
    # Also fixes double dashes, even if no space is present
    if (" " not in oldItemName) and ("--" not in oldItemName):
        return
    
    rootName, extension = getRootNameAndExtension(oldItemName)
    newItemNameSansExt = "-".join(rootName.replace("-", " ").split())
    return newItemNameSansExt + extension


def spaceFileFixLog(_1:str, _2:str, _3:str, oldItemName:str, ws, _4):
    newItemName = spaceFileFixHelper(oldItemName)
    if (not newItemName): return (False,)

    #wbm.writeItem(ws, oldItemName)
    #wbm.writeOutcomeAndIncrement(ws, newItemName, wbm.logFormat)
    wbm.incrementRowAndFileCount(ws)
    row = wbm.sheetRows[ws]
    return (True,
            ExcelWritePackage(row, wbm.ITEM_COL, oldItemName, ws),
            ExcelWritePackage(row, wbm.OUTCOME_COL, newItemName, ws, wbm.logFormat)) 

def spaceFileFixModify(longFileAbsolute:str, longDirAbsolute:str, dirAbsolute:str, oldItemName:str, ws, _):
    newItemName = spaceFileFixHelper(oldItemName)
    if (not newItemName): return (False,)

    # wbm.writeItem(ws, oldItemName)
    wbm.incrementRowAndFileCount(ws)
    row = wbm.sheetRows[ws]
    itemEwp = ExcelWritePackage(row, wbm.ITEM_COL, oldItemName, ws)

    # Log newItemName and rename file
    try:
        os.rename(longFileAbsolute, joinDirToFileName(longDirAbsolute, newItemName))
        # wbm.writeOutcomeAndIncrement(ws, newItemName, wbm.modifyFormat)
        outcomeEwp = ExcelWritePackage(row, wbm.OUTCOME_COL, newItemName, ws, wbm.modifyFormat)
    except PermissionError:
        # wbm.writeOutcomeAndIncrement(ws, "MODIFICATION FAILED. PERMISSION ERROR.", wbm.errorFormat)
        outcomeEwp = ExcelWritePackage(row, wbm.OUTCOME_COL, "MODIFICATION FAILED. PERMISSION ERROR.", ws, wbm.errorFormat)
    except Exception as e:
        # wbm.writeOutcomeAndIncrement(ws, f"MODIFICATION FAILED. {e}", wbm.errorFormat)
        outcomeEwp = ExcelWritePackage(row, wbm.OUTCOME_COL, f"MODIFICATION FAILED. {e}", ws, wbm.errorFormat)
    return (True,
            itemEwp,
            outcomeEwp)


# def deleteOldFilesStart(arg, ws):
#     writeOwnerHeader(arg, ws)
    
#     global DAYS_LOWER_BOUND
#     DAYS_LOWER_BOUND = arg[0]
#     # Double-checking that this value is usable. Dire consequences if not.
#     if (DAYS_LOWER_BOUND <= 0):
#         raise Exception("DeleteOldFile's lower bound argument cannot be less than 1.")

#     global DAYS_UPPER_BOUND
#     if len(arg) == 2:
#         DAYS_UPPER_BOUND = arg[1]
#     else:
#         DAYS_UPPER_BOUND = MAXSIZE


# def deleteOldFilesLowerboundHelper(longFileAbsolute: str, arg) -> int:
#     """Note that a file that is 23 hours and 59 minutes old is still considered 0 days old."""

#     # Get date of file. This *can* error virtue of the library functions, hence try/except
#     try: fileDate = datetime.fromtimestamp(os.path.getatime(longFileAbsolute))
#     except: return -1

#     fileDaysAgo = (TODAY - fileDate).days

#     if (DAYS_LOWER_BOUND <= fileDaysAgo): return fileDaysAgo
#     else: return 0


# def deleteOldFilesLog(longFileAbsolute:str, longDirAbsolute:str, dirAbsolute:str, itemName:str, ws, arg):
#     daysOld = deleteOldFilesLowerboundHelper(longFileAbsolute, arg)
#     # Either it's actually 0 days old or the fileDate is not within the cutOffDate range. Either way, don't flag.
#     # If it's greater than the upperbound, exit
#     if (daysOld == 0 or daysOld >= DAYS_UPPER_BOUND):
#         return (False,)

#     # wbm.writeItem(ws, itemName)
#     wbm.incrementRow(ws)
#     row = wbm.sheetRows[ws]
#     itemEwp = ExcelWritePackage(row, wbm.ITEM_COL, itemName, ws)

#     if (daysOld == -1):
#         #wbm.writeOutcome(ws, "UNABLE TO READ DATE.", wbm.errorFormat)
#         #wbm.incrementRow(ws)
#         return (2,
#                 itemEwp,
#                 ExcelWritePackage(row, wbm.OUTCOME_COL, "UNABLE TO READ DATE.", ws, wbm.errorFormat))
#     else:
#         #wbm.writeAuxiliary(ws, getOwnerCatch(longFileAbsolute))
#         #wbm.writeOutcomeAndIncrement(ws, daysOld, wbm.logFormat)
#         wbm.incrementFileCount(ws)
#         return (True,
#                 itemEwp,
#                 ExcelWritePackage(row, wbm.AUXILIARY_COL, getOwnerCatch(longFileAbsolute), ws),
#                 ExcelWritePackage(row, wbm.OUTCOME_COL, daysOld, ws, wbm.logFormat))


# ###
# def deleteOldFilesRecommendLog(longFileAbsolute:str, longDirAbsolute:str, dirAbsolute:str, itemName:str, ws, arg):
#     daysOld = deleteOldFilesLowerboundHelper(longFileAbsolute, arg)

#     if (daysOld == 0): return (False,)

#     wbm.incrementRow(ws)
#     row = wbm.sheetRows[ws]
#     itemEwp = ExcelWritePackage(row, wbm.ITEM_COL, itemName, ws)

#     if (daysOld == -1):
#         #wbm.writeItem(ws, itemName)
#         #wbm.writeOutcome(ws, "UNABLE TO READ DATE.", wbm.errorFormat)
#         #wbm.incrementRow(ws)
#         return (2,
#                 itemEwp,
#                 ExcelWritePackage(row, wbm.OUTCOME_COL, "UNABLE TO READ DATE.", ws, wbm.errorFormat))
#     else:
#         # Unlike the other variants, this will still log (and flag) items above the UPPER_BOUND threshold
#         if daysOld >= DAYS_UPPER_BOUND:
#             dynamicFormat = wbm.warningStrongFormat
#         else:
#             dynamicFormat = wbm.warningWeakFormat

#         wbm.incrementFileCount(ws)

#         #wbm.writeItem(ws, itemName, wbm.errorFormat)
#         #wbm.writeAuxiliary(ws, getOwnerCatch(longFileAbsolute))
#         #wbm.writeOutcomeAndIncrement(ws, daysOld, dynamicFormat)
#     return (True,
#             itemEwp,
#             ExcelWritePackage(row, wbm.ITEM_COL, itemName, ws, wbm.errorFormat),
#             ExcelWritePackage(row, wbm.AUXILIARY_COL, getOwnerCatch(longFileAbsolute), ws),
#             ExcelWritePackage(row, wbm.OUTCOME_COL, daysOld, ws, dynamicFormat))
# ###


# def deleteOldFilesModify(longFileAbsolute:str, longDirAbsolute:str, dirAbsolute:str, itemName:str, ws, arg):
#     daysOld = deleteOldFilesLowerboundHelper(longFileAbsolute, arg)

#     # Either it's actually 0 days old or the fileDate is not within the cutOffDate range. Either way, don't flag.
#     # If it's greater than the upperbound, exit
#     if (daysOld == 0 or daysOld >= DAYS_UPPER_BOUND):
#         return (False,)

#     # wbm.writeItem(ws, itemName)
#     wbm.incrementRow(ws)
#     row = wbm.sheetRows[ws]
#     itemEwp = ExcelWritePackage(row, wbm.ITEM_COL, itemName, ws)

#     if (daysOld == -1):
#         # wbm.writeOutcome(ws, "UNABLE TO READ DATE.", wbm.errorFormat) 
#         # wbm.incrementRow(ws)
        
#         return (2,
#                 itemEwp,
#                 ExcelWritePackage(row, wbm.OUTCOME_COL, "UNABLE TO READ DATE.", ws, wbm.errorFormat))
#     else:
#         wbm.incrementFileCount(ws)

#         try:
#             # wbm.writeAuxiliary(ws, getOwnerCatch(longFileAbsolute))
#             os.remove(longFileAbsolute)
#             #wbm.writeOutcomeAndIncrement(ws, daysOld, wbm.modifyFormat)
#         except PermissionError:
#             #wbm.writeOutcomeAndIncrement(ws, "FAILED TO DELETE. PERMISSION ERROR.", wbm.errorFormat)
#             outcomeEwp = ExcelWritePackage(row, wbm.OUTCOME_COL, "FAILED TO DELETE. PERMISSION ERROR.", ws, wbm.errorFormat)
#         except Exception as e:
#             #wbm.writeOutcomeAndIncrement(ws, f"FAILED TO DELETE. {e}", wbm.errorFormat)
#             outcomeEwp = ExcelWritePackage(row, wbm.OUTCOME_COL, f"FAILED TO DELETE. {e}", ws, wbm.errorFormat)
#     return (True,
#             itemEwp,
#             ExcelWritePackage(row, wbm.AUXILIARY_COL, getOwnerCatch(longFileAbsolute), ws),
#             outcomeEwp)


# def deleteEmptyDirectoriesLog(_, dirFolders, dirFiles, ws, arg):
#     tooFewAmount = arg[0]

#     # If even 1 folder exists, this isn't empty
#     if len(dirFolders) != 0: return (False,)

#     # If equal to tooFewAmount or less, then this folder needs to be at least flagged
#     fileAmount = len(dirFiles)
#     if fileAmount <= tooFewAmount:
#         #wbm.writeOutcomeAndIncrement(ws, fileAmount, wbm.logFormat)
#         wbm.incrementRowAndFileCount(ws)
#         return (True,
#                 ExcelWritePackage(wbm.sheetRows[ws], wbm.OUTCOME_COL, fileAmount, ws, wbm.logFormat))
#     return (False,)


# ###
# def deleteEmptyDirectoriesRecommendLog(dirAbsolute, dirFolders, dirFiles, ws, arg):
#     tooFewAmount = arg[0]

#     if len(dirFolders) != 0: return (False,)

#     fileAmount = len(dirFiles)
#     if fileAmount <= tooFewAmount:
        
#         ## HARD CODED RECOMMENDATION VALUES
#         # Dynamic Format
#         if fileAmount <= 1:
#             dynamicFormat = wbm.warningStrongFormat
#         elif fileAmount == 2:
#             dynamicFormat = wbm.warningWeakFormat
#         else:
#             dynamicFormat = wbm.logFormat

#         # wbm.writeOutcomeAndIncrement(ws, fileAmount, dynamicFormat)
#         wbm.incrementRowAndFileCount(ws)
#         return (True,
#                 ExcelWritePackage(wbm.sheetRows[ws], wbm.OUTCOME_COL, fileAmount, ws, dynamicFormat))
#     return (False,)
# ###
    
    
# def deleteEmptyDirectoriesModify(dirAbsolute, dirFolders, dirFiles, ws, arg):
#     tooFewAmount = arg[0]

#     if len(dirFolders) != 0: return (False,)

#     fileAmount = len(dirFiles)
#     if fileAmount <= tooFewAmount:
#         wbm.incrementRowAndFileCount(ws)

#         # If it specifically has 0 files, delete the folder
#         if (fileAmount == 0):
#             try: 
#                 os.rmdir(addLongPathPrefix(dirAbsolute))
#                 # wbm.writeOutcomeAndIncrement(ws, fileAmount, wbm.modifyFormat)
#                 outcomeEwp = ExcelWritePackage(wbm.sheetRows[ws], wbm.OUTCOME_COL, fileAmount, ws, wbm.modifyFormat)
#             except Exception as e:
#                 # wbm.writeOutcomeAndIncrement(ws, f"0 FILES. UNABLE TO DELETE. {e}", wbm.errorFormat)
#                 outcomeEwp = ExcelWritePackage(wbm.sheetRows[ws], wbm.OUTCOME_COL, f"0 FILES. UNABLE TO DELETE. {e}", ws, wbm.errorFormat)

#             return (True,
#                     outcomeEwp
#                     )
        
#         # Otherwise, just flag as usual
#         else:
#             # wbm.writeOutcomeAndIncrement(ws, fileAmount, wbm.logFormat)

#             return (True,
#                     ExcelWritePackage(wbm.sheetRows[ws], wbm.OUTCOME_COL, fileAmount, ws, wbm.logFormat))
#     return (False,)


def searchAndReplaceFolderHelper(oldFolderName:str, arg):
    newFolderName = oldFolderName
    for argPair in arg:
        toBeReplaced, replacer = argPair
        newFolderName = newFolderName.replace(toBeReplaced, replacer)

    if (oldFolderName == newFolderName): return
    return newFolderName


def searchAndReplaceFolderLog(dirAbsolute, dirFolders, dirFiles, ws, arg):
    oldFolderName = getDirectoryBaseName(dirAbsolute)

    if not (newFolderName := searchAndReplaceFolderHelper(oldFolderName, arg)):
        return (False,)
    
    # wbm.writeDir(ws, dirAbsolute, wbm.dirFormat)
    # wbm.writeItem(ws, oldFolderName, wbm.errorFormat)
    # wbm.writeOutcomeAndIncrement(ws, newFolderName, wbm.logFormat)
    wbm.incrementRowAndFileCount(ws)
    row = wbm.sheetRows[ws]

    return (True,
            ExcelWritePackage(row, wbm.DIR_COL, dirAbsolute, ws, wbm.dirFormat),
            ExcelWritePackage(row, wbm.ITEM_COL, oldFolderName, ws, wbm.errorFormat),
            ExcelWritePackage(row, wbm.OUTCOME_COL, newFolderName, ws, wbm.logFormat)
            )


def searchAndReplaceFolderModify(dirAbsolute, dirFolders, dirFiles, ws, arg):
    oldFolderName = getDirectoryBaseName(dirAbsolute)

    if not (newFolderName := searchAndReplaceFolderHelper(oldFolderName, arg)):
        return (False,)
    
    directoryOfFolder = getDirectoryDirName(dirAbsolute)
    FOLDER_RENAMES.append([directoryOfFolder, oldFolderName, newFolderName])
    return (3,)


def searchAndReplaceFileHelper(oldItemName:str, arg):
    oldItemNameSansExt, extension = getRootNameAndExtension(oldItemName)
    
    # Order of argument pairs given matters.
    newItemNameSansExt = oldItemNameSansExt
    for argPair in arg:
        toBeReplaced, replacer = argPair
        newItemNameSansExt = newItemNameSansExt.replace(toBeReplaced, replacer)

    if (oldItemNameSansExt == newItemNameSansExt): return
    return newItemNameSansExt + extension


def searchAndReplaceFileLog(longFileAbsolute:str, longDirAbsolute:str, _:str, oldItemName:str, ws, arg):
    if not (newItemName := searchAndReplaceFileHelper(oldItemName, arg)): return (False,)

    #wbm.writeItem(ws, oldItemName, wbm.errorFormat)
    #wbm.writeOutcomeAndIncrement(ws, newItemName, wbm.logFormat)
    wbm.incrementRowAndFileCount(ws)
    row = wbm.sheetRows[ws]
    return (True,
            ExcelWritePackage(row, wbm.ITEM_COL, oldItemName, ws, wbm.errorFormat),
            ExcelWritePackage(row, wbm.OUTCOME_COL, newItemName, ws, wbm.logFormat))


def searchAndReplaceFileModify(longFileAbsolute:str, longDirAbsolute:str, dirAbsolute:str, oldItemName:str, ws, arg):
    if not (newItemName := searchAndReplaceFileHelper(oldItemName, arg)): return (False,)

    # wbm.writeItem(ws, oldItemName, wbm.errorFormat)
    wbm.incrementRowAndFileCount(ws)
    row = wbm.sheetRows[ws]
    itemEwp = ExcelWritePackage(row, wbm.ITEM_COL, oldItemName, ws, wbm.errorFormat)

    try:
        os.rename(longFileAbsolute, joinDirToFileName(longDirAbsolute, newItemName))
        # wbm.writeOutcomeAndIncrement(ws, newItemName, wbm.modifyFormat)
        outcomeEwp = ExcelWritePackage(row, wbm.OUTCOME_COL, newItemName, ws, wbm.modifyFormat)
    except PermissionError:
        # wbm.writeOutcomeAndIncrement(ws, "MODIFICATION FAILED. PERMISSION ERROR.", wbm.errorFormat)
        outcomeEwp = ExcelWritePackage(row, wbm.OUTCOME_COL, "MODIFICATION FAILED. PERMISSION ERROR.", ws, wbm.errorFormat)
    except Exception as e:
        # wbm.writeOutcomeAndIncrement(ws, f"MODIFICATION FAILED. {e}", wbm.errorFormat)
        outcomeEwp = ExcelWritePackage(row, wbm.OUTCOME_COL, f"MODIFICATION FAILED. {e}", ws, wbm.errorFormat)
    return (True,
            itemEwp,
            outcomeEwp)


def deleteEmptyFilesStart(_, ws):
    writeOwnerHeader(_, ws)


def deleteEmptyFilesLog(longFileAbsolute:str, longDirAbsolute:str, dirAbsolute:str, itemName:str, ws, _):
    try:
        fileSize = os.path.getsize(longFileAbsolute) # Bytes
    except PermissionError:
        #wbm.writeItem(ws, itemName)
        #wbm.writeOutcome(ws, "UNABLE TO READ. PERMISSION ERROR.", wbm.errorFormat)
        #wbm.incrementRow(ws)
        wbm.incrementRow(ws)
        row = wbm.sheetRows[ws]
        return (2,
                ExcelWritePackage(row, wbm.ITEM_COL, itemName, ws),
                ExcelWritePackage(row, wbm.OUTCOME_COL, "UNABLE TO READ. PERMISSION ERROR.", ws, wbm.errorFormat))
    except Exception as e:
        #wbm.writeItem(ws, itemName)
        #wbm.writeOutcome(ws, f"UNABLE TO READ. {e}", wbm.errorFormat)
        #wbm.incrementRow(ws)
        wbm.incrementRow(ws)
        row = wbm.sheetRows[ws]
        return (2,
                ExcelWritePackage(row, wbm.ITEM_COL, itemName, ws),
                ExcelWritePackage(row, wbm.OUTCOME_COL, f"UNABLE TO READ. {e}", ws, wbm.errorFormat))

    if fileSize == 0:
        #wbm.writeItem(ws, itemName)
        #wbm.writeAuxiliary(ws, getOwnerCatch(longFileAbsolute))
        #wbm.writeOutcomeAndIncrement(ws, "", wbm.logFormat)
        wbm.incrementRowAndFileCount(ws)
        row = wbm.sheetRows[ws]
        return (True,
                ExcelWritePackage(row, wbm.ITEM_COL, itemName, ws),
                ExcelWritePackage(row, wbm.AUXILIARY_COL, getOwnerCatch(longFileAbsolute), ws),
                ExcelWritePackage(row, wbm.OUTCOME_COL, "", ws, wbm.logFormat))
    return (False,)


###
def deleteEmptyFilesRecommendLog(longFileAbsolute:str, longDirAbsolute:str, dirAbsolute:str, itemName:str, ws, _):
    try:
        fileSize = os.path.getsize(longFileAbsolute)
    except PermissionError:
        #wbm.writeItem(ws, itemName)
        #wbm.writeOutcome(ws, "UNABLE TO READ. PERMISSION ERROR.", wbm.errorFormat)
        #wbm.incrementRow(ws)
        wbm.incrementRow(ws)
        row = wbm.sheetRows[ws]
        return (2,
                ExcelWritePackage(row, wbm.ITEM_COL, itemName, ws),
                ExcelWritePackage(row, wbm.OUTCOME_COL, "UNABLE TO READ. PERMISSION ERROR.", ws, wbm.errorFormat))
    except Exception as e:
        #wbm.writeItem(ws, itemName)
        #wbm.writeOutcome(ws, f"UNABLE TO READ. {e}", wbm.errorFormat)
        #wbm.incrementRow(ws)
        wbm.incrementRow(ws)
        row = wbm.sheetRows[ws]
        return (2,
                ExcelWritePackage(row, wbm.ITEM_COL, itemName, ws),
                ExcelWritePackage(row, wbm.OUTCOME_COL, f"UNABLE TO READ. {e}", ws, wbm.errorFormat))


    if fileSize == 0:
        #wbm.writeItem(ws, itemName, wbm.warningStrongFormat)
        #wbm.writeAuxiliary(ws, getOwnerCatch(longFileAbsolute))
        #wbm.writeOutcomeAndIncrement(ws, "", wbm.warningStrongFormat)
        wbm.incrementRowAndFileCount(ws)
        row = wbm.sheetRows[ws]
        return (True,
                ExcelWritePackage(row, wbm.ITEM_COL, itemName, ws, wbm.warningStrongFormat),
                ExcelWritePackage(row, wbm.AUXILIARY_COL, getOwnerCatch(longFileAbsolute), ws),
                ExcelWritePackage(row, wbm.OUTCOME_COL, "", ws, wbm.warningStrongFormat))
    return (False,)
###


def deleteEmptyFilesModify(longFileAbsolute:str, longDirAbsolute:str, dirAbsolute:str, itemName:str, ws, _):
    """Glitch exists in that the current excel file will be considered empty.
    However, despite claiming so, the program does not actually delete it.'"""

    try:
        fileSize = os.path.getsize(longFileAbsolute)
    except PermissionError:
        #wbm.writeItem(ws, itemName)
        #wbm.writeOutcome(ws, "UNABLE TO READ. PERMISSION ERROR.", wbm.errorFormat)
        #wbm.incrementRow(ws)
        wbm.incrementRow(ws)
        row = wbm.sheetRows[ws]
        return (2,
                ExcelWritePackage(row, wbm.ITEM_COL, itemName, ws),
                ExcelWritePackage(row, wbm.OUTCOME_COL, "UNABLE TO READ. PERMISSION ERROR.", ws, wbm.errorFormat))
    
    except Exception as e:
        #wbm.writeItem(ws, itemName)
        #wbm.writeOutcome(ws, f"UNABLE TO READ. {e}", wbm.errorFormat)
        #wbm.incrementRow(ws)
        wbm.incrementRow(ws)
        row = wbm.sheetRows[ws]
        return (2,
                ExcelWritePackage(row, wbm.ITEM_COL, itemName, ws),
                ExcelWritePackage(row, wbm.OUTCOME_COL, f"UNABLE TO READ. {e}", ws, wbm.errorFormat))


    # Stage for deletion
    if fileSize == 0:
        # wbm.writeItem(ws, itemName)
        wbm.incrementRowAndFileCount(ws)
        row = wbm.sheetRows[ws]
        itemEwp = ExcelWritePackage(row, wbm.ITEM_COL, itemName, ws)
        auxiliaryEwp = ExcelWritePackage(row, wbm.AUXILIARY_COL, getOwnerCatch(longFileAbsolute), ws)

        try:
            # wbm.writeAuxiliary(ws, getOwnerCatch(longFileAbsolute))
            os.remove(longFileAbsolute)
            # wbm.writeOutcomeAndIncrement(ws, "", wbm.modifyFormat)
            outcomeEwp = ExcelWritePackage(row, wbm.OUTCOME_COL, "", ws, wbm.modifyFormat)
        except PermissionError:
            # wbm.writeOutcomeAndIncrement(ws, "FAILED TO DELETE. PERMISSION ERROR.", wbm.errorFormat)
            outcomeEwp = ExcelWritePackage(row, wbm.OUTCOME_COL, "FAILED TO DELETE. PERMISSION ERROR.", ws, wbm.errorFormat)
        except Exception as e:
            # wbm.writeOutcomeAndIncrement(ws, f"FAILED TO DELETE. {e}", wbm.errorFormat)
            outcomeEwp = ExcelWritePackage(row, wbm.OUTCOME_COL, f"FAILED TO DELETE. {e}", ws, wbm.errorFormat)
        return (True,
                itemEwp,
                auxiliaryEwp,
                outcomeEwp)    
    return (False,)



# def deleteIdenticalFilesStart(_, ws):
#     global HASH_AND_FILES
#     global EMPTY_INPUT_HASH_CODE
#     global LOCK_DELETE_IDENTICAL_FILES

#     HASH_AND_FILES = {}
#     hashFunc = hashlib.new("sha256")
#     hashFunc.update("".encode())
#     EMPTY_INPUT_HASH_CODE = hashFunc.hexdigest()
#     LOCK_DELETE_IDENTICAL_FILES = Lock()

# def deleteIdenticalFilesHelper(longFileAbsolute:str):    
#     hashFunc = hashlib.new("sha256")
    
#     with open(longFileAbsolute, "rb") as file:
#         while chunk := file.read(8192):
#             hashFunc.update(chunk)

#     return hashFunc.hexdigest()

# def deleteIdenticalFilesLogConcurrent(longFileAbsolute:str, longDirAbsolute:str, dirAbsolute:str, itemName:str, ws, _):
#     try:
#         hashCode = deleteIdenticalFilesHelper(longFileAbsolute)
#     except Exception:  # FileNotFoundError, PermissionError, OSError, UnicodeDecodeError
#         # Unlike other procedures, this won't print out the error; it'll just assume it's not a duplicate.
#         return (False,)
    
#     if (not hashCode or hashCode == EMPTY_INPUT_HASH_CODE):
#         return (False,)
    
#     with LOCK_DELETE_IDENTICAL_FILES:
#         if hashCode in HASH_AND_FILES:
#             HASH_AND_FILES[hashCode][0].append(itemName)
#             HASH_AND_FILES[hashCode][1].append(dirAbsolute)
#             wbm.incrementFileCount(ws)
#             return (3,)
#         else:
#             HASH_AND_FILES[hashCode] = ([itemName], [dirAbsolute])
#             return (False,)


# def deleteIdenticalFilesRecommendPost(ws):
#     ws.write(0, 0, "Separator", wbm.headerFormat)
#     ws.write(0, 1, "Files", wbm.headerFormat)
#     ws.write(0, 2, "Directories", wbm.headerFormat)
#     ws.write(0, 3, "Owner", wbm.headerFormat)

#     row = 1
#     folderAndItem = defaultdict(list)

#     for hashCode in HASH_AND_FILES.keys():
#         if (numOfFiles := len(HASH_AND_FILES[hashCode][0])) > 1:
#             ws.write(row, 0, "------------", wbm.logFormat)

#             # If there are 3 or more duplicates, highlight them all yellow at least. Otherwise, just flag as normal.
#             if (numOfFiles >= 3):
#                 defaultItemFormat = wbm.warningWeakFormat
#             else:
#                 defaultItemFormat = wbm.errorFormat

#             # Sort this group of identical files with dirAbsolute as the key, and itemName as the values
#             for i in range(numOfFiles):
#                 folderAndItem[HASH_AND_FILES[hashCode][1][i]].append(HASH_AND_FILES[hashCode][0][i])

#             for dirAbsoluteKey in folderAndItem.keys():
#                 # If 2 or more files are identical AND reside in the same folder
#                 if (dirAbsoluteNumOfFiles := len(folderAndItem[dirAbsoluteKey])) > 1:

#                     # Sort the list of items in descending order, ordered by number of characters
#                     folderAndItem[dirAbsoluteKey].sort(key=len, reverse=True)

#                     # Write the first one normally
#                     ws.write(row, 1, folderAndItem[dirAbsoluteKey][0], defaultItemFormat)
#                     ws.write(row, 2, dirAbsoluteKey, wbm.dirFormat)
#                     ws.write(row, 3, getOwnerCatch(
#                         joinDirToFileName(dirAbsoluteKey, folderAndItem[dirAbsoluteKey][0])))
#                     row += 1

#                     # Write the rest in strong warning format
#                     for i in range(1, dirAbsoluteNumOfFiles):
#                         ws.write(row, 1, folderAndItem[dirAbsoluteKey][i], wbm.warningStrongFormat)
#                         ws.write(row, 2, dirAbsoluteKey, wbm.dirFormat)
#                         ws.write(row, 3, getOwnerCatch(
#                             joinDirToFileName(dirAbsoluteKey, folderAndItem[dirAbsoluteKey][i])))
#                         row += 1
                
#                 # If this file is only duplicated once in this directory, just write it normally
#                 else:
#                     ws.write(row, 1, folderAndItem[dirAbsoluteKey][0], defaultItemFormat)
#                     ws.write(row, 2, dirAbsoluteKey, wbm.dirFormat)
#                     ws.write(row, 3, getOwnerCatch(
#                         joinDirToFileName(dirAbsoluteKey, folderAndItem[dirAbsoluteKey][0])))
#                     row += 1

#             folderAndItem.clear()
                
#     HASH_AND_FILES.clear()
#     ws.freeze_panes(1, 0)


# def deleteIdenticalFilesPost(ws):
#     ws.write(0, 0, "Separator", wbm.headerFormat)
#     ws.write(0, 1, "Files", wbm.headerFormat)
#     ws.write(0, 2, "Directories", wbm.headerFormat)
#     ws.write(0, 3, "Owner", wbm.headerFormat)

#     row = 1
#     for hashCode in HASH_AND_FILES.keys():
#         if (numOfFiles := len(HASH_AND_FILES[hashCode][0])) > 1:
#             ws.write(row, 0, "------------", wbm.logFormat)

#             for i in range(numOfFiles):
#                 ws.write(row, 1, HASH_AND_FILES[hashCode][0][i], wbm.errorFormat)
#                 ws.write(row, 2, HASH_AND_FILES[hashCode][1][i], wbm.dirFormat)
#                 ws.write(row, 3, joinDirToFileName(HASH_AND_FILES[hashCode][1][i], HASH_AND_FILES[hashCode][0][i]))
#                 row += 1
                
#     HASH_AND_FILES.clear()
#     ws.freeze_panes(1, 0)