import os
from workbookManager import WorkbookManager
from datetime import date
import threading

GROUP_SEPARATOR = "------------"

def setWorkbookManager(newManager: WorkbookManager):
    # Globally declare the WorkbookManager object
    global wbm
    wbm = newManager


def addLongPathPrefix(dirAbsolute):
    if dirAbsolute.startswith('\\\\'):
        return '\\\\?\\UNC' + dirAbsolute[1:]
    else:
        return '\\\\?\\' + dirAbsolute


def writeOwnerAndLastModifiedHeaders(_, ws):
    ws.write(0, wbm.AUXILIARY_COL, "Owner", wbm.headerFormat)
    ws.write(0, wbm.AUXILIARY_COL+1, "Last Modified", wbm.headerFormat)

def writeDefaultHeaders(_, ws):
    ws.freeze_panes(1, 0)
    ws.write(0, wbm.DIR_COL, "Directory", wbm.headerFormat)
    ws.write(0, wbm.ITEM_COL, "Item", wbm.headerFormat)
    ws.write(0, wbm.OUTCOME_COL, "Outcome", wbm.headerFormat)


def writeDefaultAndOwnerAndLastModifiedHeaders(_, ws):
    writeDefaultHeaders(_, ws)
    writeOwnerAndLastModifiedHeaders(_, ws)


def joinDirToFileName(dirAbsolute, fileName):
    return dirAbsolute + "\\" + fileName


def getRootNameAndExtension(itemName):
    rootName, extension = os.path.splitext(itemName)
    return (rootName, extension.lower())


def getDirectoryBaseName(dirAbsolute):
    return os.path.basename(dirAbsolute)


def getDirectoryDirName(dirAbsolute):
    return os.path.dirname(dirAbsolute)


LAST_MODIFIED_INFO_CACHE = {}
LM_CACHE_EVENTS = {}
LM_CACHE_LOCK = threading.Lock()
dummyData = "DUMMY"

def getLastModifiedDate(longFileAbsolute):
    lastModifiedInfo = LAST_MODIFIED_INFO_CACHE.get(longFileAbsolute)
    
    # fast track
    if lastModifiedInfo is not None and lastModifiedInfo != dummyData:
        return lastModifiedInfo
    
    # middle track
    with LM_CACHE_LOCK:
        lastModifiedInfo = LAST_MODIFIED_INFO_CACHE.get(longFileAbsolute)

        if lastModifiedInfo is None:
            LAST_MODIFIED_INFO_CACHE[longFileAbsolute] = dummyData
            LM_CACHE_EVENTS[longFileAbsolute] = threading.Event()
            computeLastModified = True
        elif lastModifiedInfo == dummyData:
            computeLastModified = False
        else:
            return lastModifiedInfo

    # slow track
    if computeLastModified:
        try:
            lastModifiedInfo = date.fromtimestamp(os.path.getmtime(longFileAbsolute)).isoformat()
        except Exception as e:
            lastModifiedInfo = f"GET LAST MODIFIED FAILED: {e}"
        
        LAST_MODIFIED_INFO_CACHE[longFileAbsolute] = lastModifiedInfo
        LM_CACHE_EVENTS[longFileAbsolute].set()
        return lastModifiedInfo
    
    LM_CACHE_EVENTS[longFileAbsolute].wait()
    return LAST_MODIFIED_INFO_CACHE[longFileAbsolute]

