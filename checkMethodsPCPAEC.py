from typing import Union, Set
import string
import os
from workbookManager import WorkbookManager


# Global variable so it's set in stone
# Used by badCharacters()
# This includes space as a permissibleCharacter as it is checked by a different method instead
permissibleCharacters = set(string.ascii_letters + string.digits + "-. ")

# Declare a global variable within a function
# ~ Usually a bad idea, but here, it makes sense
def setWorkBookManager(newManager: WorkbookManager):
    global wbm
    wbm = newManager


def hasSpace(dirAbsolute:str, itemName:str, ws) -> bool:
    if " " in itemName:
        wbm.writeInCell(ws, wbm.ITEM_COL, itemName, wbm.fileErrorFormat, rowIncrement=1, errorIncrement=1)
        # no need to write in the error column as this won't vary between errors found
        return True
    
    return False


def overCharacterLimit(dirAbsolute:str, itemName:str, ws) -> bool:
    absoluteItemLength = len(dirAbsolute + "/" + itemName)
    if (absoluteItemLength > 200):
        wbm.writeInCell(ws, wbm.ITEM_COL, itemName, wbm.fileErrorFormat)
        wbm.writeInCell(ws, wbm.ERROR_COL, "{} > 200".format(absoluteItemLength), rowIncrement=1, errorIncrement=1)
        return True
    
    return False


def badCharacters(dirAbsolute:str, itemName:str, ws) -> Set[str]:
    """Does not check for SPC characters nor extra periods."""
    
    badChars = set()
    itemNameLength = len(itemName)

    for i in range(itemNameLength):
        if itemName[i] not in permissibleCharacters:
            # Not necessary to set errorPresent=True because we're returning non-empty badChars instead
            badChars.add(itemName[i])
        
        # double dash error
        elif itemName[i:i+2] == "--":
            badChars.add(itemName[i])

    # write to own sheet here
    if (badChars):
        wbm.writeInCell(ws, wbm.ITEM_COL, itemName, wbm.fileErrorFormat)
        wbm.writeInCell(ws, wbm.ERROR_COL, "".join(badChars), rowIncrement=1, errorIncrement=1)

        
    return badChars


def fixSpacesShow(dirAbsolute:str, oldItemName:str, ws):
    # if no spaces, just leave
    if (" " not in oldItemName):
        return
    
    # Replace "-" characters with " " to make the string homogenous for the upcoming split()
    # split() automatically removes leading, trailing, and excess middle whitespace
    newItemName = oldItemName.replace("-", " ").split()
    newItemName = "-".join(newItemName)

    # For if there's a dash to the left of the file extension period
    lastPeriodIndex = newItemName.rfind(".")
    # If lastPeriodIndex isn't the very first character and there actually is a period
    if (lastPeriodIndex > 0 and newItemName[lastPeriodIndex -1] == "-"):
        newItemName = newItemName[0:lastPeriodIndex-1] + newItemName[lastPeriodIndex:]

    # write the new name in the appropriate field
    wbm.writeInCell(ws, wbm.ITEM_COL, oldItemName)
    wbm.writeInCell(ws, wbm.RENAME_COL, newItemName, wbm.showRenameFormat, 1, 1)

    
def fixSpacesDo(dirAbsolute:str, oldItemName:str, ws):
    if (" " not in oldItemName):
        return
    
    newItemName = oldItemName.replace("-", " ").split()
    newItemName = "-".join(newItemName)

    lastPeriodIndex = newItemName.rfind(".")
    if (lastPeriodIndex > 0 and newItemName[lastPeriodIndex -1] == "-"):
        newItemName = newItemName[0:lastPeriodIndex-1] + newItemName[lastPeriodIndex:]

    wbm.writeInCell(ws, wbm.ITEM_COL, oldItemName)

    # Log newItemName and rename file
    try:
        os.rename(dirAbsolute + "/" + oldItemName, dirAbsolute + "/" + newItemName)
        wbm.writeInCell(ws, wbm.RENAME_COL, newItemName, wbm.renameFormat, 1, 1)
    except PermissionError:
        wbm.writeInCell(wbm.mainSheet, wbm.RENAME_COL, "FILE LOCKED. RENAME FAILED.", wbm.fileErrorFormat, 1, 0)
    except OSError:
        wbm.writeInCell(wbm.mainSheet, wbm.RENAME_COL, "OS ERROR. RENAME FAILED.", wbm.fileErrorFormat, 1, 0)
