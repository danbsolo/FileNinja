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




def checkNamingConvention(dirAbsolute:str, itemName:str) -> Union[Set[str], bool]:
    # If errorChars is empty, returns errorPresent, which may be True or False.
    errorPresent = False
    variableErrorCol = wbm.ERROR_COL

    # If it's just a temporary file via Microsoft, end here
    if (itemName[0:2] == "~$"):
        return errorPresent
    
    absoluteItemLength = len(dirAbsolute + "/" + itemName)
    if (absoluteItemLength > 200):
        errorPresent = True
        wbm.writeInCell(wbm.mainSheet, variableErrorCol, "{} chars > 200. Terminating checks.".format(absoluteItemLength))
        # variableErrorCol += 1
        return errorPresent

    errorChars = set()
    periodCount = 0
    itemNameLength = len(itemName)

    for i in range(itemNameLength):
        if itemName[i] not in permissibleCharacters:
            # Not necessary to set errorPresent=True because we're returning non-empty errorChars instead
            errorChars.add(itemName[i])
        
        # double dash error
        elif itemName[i:i+2] == "--":
            errorChars.add(itemName[i])
        
        if itemName[i] == ".":
            periodCount += 1

    # only checks if 2 or more periods are present
    # TODO: Error if no periods are present (since folder names aren't looked at)
    if (periodCount >= 2):
        errorChars.add(".")

    # if not empty. AKA, some error character has been detected
    if (errorChars):
        # Surrounding the set with "||" makes space chars visible
        wbm.writeInCell(wbm.mainSheet, variableErrorCol, "Bad chars: |{}|".format("".join(errorChars)))

        # Only returns a set if it's populated
        return errorChars

    return errorPresent



def showRename(dirAbsolute:str , itemName:str) -> bool:
    result = checkNamingConvention(dirAbsolute, itemName)

    if (isinstance(result, bool)):
        return result
    
    if (not {" ", "-"}.isdisjoint(result)):
        newItemName = produceNewName(itemName)

        # Log newItemName, but don't follow-through with renaming
        wbm.writeInCell(wbm.mainSheet, wbm.RENAME_COL, newItemName, wbm.showRenameFormat)

    return True



def renameItem(dirAbsolute:str , itemName:str) -> bool:
    result = checkNamingConvention(dirAbsolute, itemName)

    if (isinstance(result, bool)):
        return result
    
    ### Change spaces and double dashes into dashes
    if (not {" ", "-"}.isdisjoint(result)):
        newItemName = produceNewName(itemName)

        # Log newItemName and rename file
        try:
            os.rename(dirAbsolute + "/" + itemName, dirAbsolute + "/" + newItemName)
            wbm.writeInCell(wbm.mainSheet, wbm.RENAME_COL, newItemName, wbm.renameFormat)
        except PermissionError:
            wbm.writeInCell(wbm.mainSheet, wbm.RENAME_COL, "FILE LOCKED. RENAME FAILED.", wbm.fileErrorFormat)
        except OSError:
            wbm.writeInCell(wbm.mainSheet, wbm.RENAME_COL, "OS ERROR. RENAME FAILED.", wbm.fileErrorFormat)

    return True



def produceNewName(oldItemName:str) -> str:
    # Replace "-" characters with " " to make the string homogenous for the upcoming split()
    # split() automatically removes leading, trailing, and excess middle whitespace
    newItemName = oldItemName.replace("-", " ").split()
    newItemName = "-".join(newItemName)

    # For if there's a dash to the left of the file extension period
    lastPeriodIndex = newItemName.rfind(".")
    # If lastPeriodIndex isn't the very first character and there actually is a period
    if (lastPeriodIndex > 0 and newItemName[lastPeriodIndex -1] == "-"):
        newItemName = newItemName[0:lastPeriodIndex-1] + newItemName[lastPeriodIndex:]
    
    return newItemName