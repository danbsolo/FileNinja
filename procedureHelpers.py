import os
from workbookManager import WorkbookManager


def setWorkbookManager(newManager: WorkbookManager):
    # Globally declare the WorkbookManager object
    global wbm
    wbm = newManager


def addLongPathPrefix(dirAbsolute):
    if dirAbsolute.startswith('\\\\'):
        return '\\\\?\\UNC' + dirAbsolute[1:]
    else:
        return '\\\\?\\' + dirAbsolute


def writeOwnerHeader(_, ws):
    ws.write(0, wbm.AUXILIARY_COL, "Owner", wbm.headerFormat)


def writeDefaultHeaders(_, ws):
    ws.freeze_panes(1, 0)
    ws.write(0, wbm.DIR_COL, "Directory", wbm.headerFormat)
    ws.write(0, wbm.ITEM_COL, "Item", wbm.headerFormat)
    ws.write(0, wbm.OUTCOME_COL, "Outcome", wbm.headerFormat)


def writeDefaultAndOwnerHeaders(_, ws):
    writeDefaultHeaders(_, ws)
    writeOwnerHeader(_, ws)


def joinDirToFileName(dirAbsolute, fileName):
    return dirAbsolute + "\\" + fileName


def getRootNameAndExtension(itemName):
    rootName, extension = os.path.splitext(itemName)
    return (rootName, extension.lower())


def getDirectoryBaseName(dirAbsolute):
    return os.path.basename(dirAbsolute)


def getDirectoryDirName(dirAbsolute):
    return os.path.dirname(dirAbsolute)