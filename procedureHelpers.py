import os
from workbookManager import WorkbookManager
import string
from datetime import datetime
import hashlib
from getFileOwner import getOwnerCatch
from collections import defaultdict
from ExcelWritePackage import ExcelWritePackage
from threading import Lock
from sys import maxsize as MAXSIZE


# Used by badCharFind
# Includes ' ' (space) as there is a separate procedure for finding that error
PERMISSIBLE_CHARACTERS = set(string.ascii_letters + string.digits + "- ")
CHARACTER_LIMIT = 200

# Used by oldFileFind and deleteOldFiles
TODAY = datetime.now()



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

# def writeDefaultHeaders(_, ws):
#     ws.freeze_panes(1, 0)
#     ws.write(0, wbm.DIR_COL, "Directory", wbm.headerFormat)
#     ws.write(0, wbm.ITEM_COL, "Item", wbm.headerFormat)
#     ws.write(0, wbm.OUTCOME_COL, "Outcome", wbm.headerFormat)

# def writeDefaultAndOwnerHeaders(_, ws):
#     writeDefaultAndOwnerHeaders(_, ws)
#     writeOwnerHeader(_, ws)


def joinDirToFileName(dirAbsolute, fileName):
    return dirAbsolute + "\\" + fileName



def getRootNameAndExtension(itemName):
    return os.path.splitext(itemName)



def getDirectoryBaseName(dirAbsolute):
    return os.path.basename(dirAbsolute)


def getDirectoryDirName(dirAbsolute):
    return os.path.dirname(dirAbsolute)