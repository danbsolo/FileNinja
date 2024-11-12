from typing import Union, Set
import string
import os
from workBookManager import WorkbookManager


# Global variable so it's set in stone
permissibleCharacters = set(string.ascii_letters + string.digits + "-.")

# Declare a global variable within a function
# ~ Usually a bad idea, but here, it makes sense
def setWorkBookManager(newManager: WorkbookManager):
    global wbm
    wbm = newManager


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

    # if this program didn't include folders, then a periodCount of 0 would be bad; no file extension
    # but alas, the only true error is if two or more periods are present
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