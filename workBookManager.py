import xlsxwriter
import os
from typing import List, Tuple, Callable, Union, Set
import string

class WorkbookManager:

    def __init__(self, workbookName: str):
        self.workbookName = workbookName
        self.wb = xlsxwriter.Workbook(workbookName)
        
        # TODO: add more worksheets here if/when necessary
        self.mainSheetRow = 1
        self.mainSheet = self.wb.add_worksheet("MainSheet")
        

        ### Extra data members
        # Used by checkMethods
        self.permissibleCharacters = set(string.ascii_letters + string.digits + "-.")
        
        # An empty function by default. Called within fileCrawl()
        self.checkMethod = lambda: None

        ## Set formatting 
        # Column indexes
        self.DIR_COL = 0
        self.ITEM_COL = 1
        self.RENAME_COL = 2
        self.ERROR_COL = 3

        # Column width
        self.mainSheet.set_column(self.DIR_COL, self.DIR_COL, 50)
        self.mainSheet.set_column(self.ITEM_COL, self.RENAME_COL, 30)
        self.mainSheet.set_column(self.ERROR_COL, self.ERROR_COL+5, 20)  # +5 for adding more than one error

        # Default cell styles
        self.dirColFormat = self.wb.add_format({
            "bg_color": "#99CCFF", # blueish
            "bold": True
        })

        self.fileErrorFormat = self.wb.add_format({
            "bg_color": "#FF4444", # reddish
            "bold": True
        })

        self.renameFormat = self.wb.add_format({
            "bg_color": "#9999FF", # purplish
            "bold": True
        })

        self.headerFormat = self.wb.add_format({
            "bg_color": "#C0C0C0", # grayish
            "bold": True
        })

        # Write headers
        self.mainSheet.write(0, self.DIR_COL, "Directories", self.headerFormat)
        self.mainSheet.write(0, self.ITEM_COL, "Items", self.headerFormat)
        self.mainSheet.write(0, self.RENAME_COL, "Rename", self.headerFormat)
        self.mainSheet.write(0, self.ERROR_COL, "Errors", self.headerFormat)
        


    def fileCrawl(self, dirAbsolute, dirItems: List[str]):
        # can be used to write down folders as well
        for itemName in dirItems:
            
            # If no error present
            if (self.checkMethod(dirAbsolute, itemName)):
                self.mainSheet.write(self.mainSheetRow, self.ITEM_COL, itemName, self.fileErrorFormat)
            else:
                self.mainSheet.write(self.mainSheetRow, self.ITEM_COL, itemName)

            self.mainSheetRow += 1
    


    def folderCrawl(self, dirTree: List[Tuple[str, list, list]]):
        for (dirAbsolute, dirFolders, dirFiles) in dirTree:
            self.mainSheet.write(self.mainSheetRow, self.DIR_COL, dirAbsolute, self.dirColFormat)
            self.fileCrawl(dirAbsolute, dirFolders + dirFiles)



    def setCheckMethod(self, functionSelection: Callable) -> bool:
        if (functionSelection == "PC-PAEC"):
            self.checkMethod = self.checkPCPAECNamingConvention
        elif (functionSelection == "PC-PAEC-Rename"):
            # TODO: make the corresponding function for renaming
            self.checkMethod = self.checkPCPAECNamingConventionRename

        # Nothing has been selected. Return False to indicate failure.
        else:
            return False
        
        return True



    def checkPCPAECNamingConvention(self, dirAbsolute: str, itemName: str) -> Union[Set[str], bool]:
        # If errorChars is empty, returns errorPresent, which may be True or False.
        errorPresent = False
        variableErrorCol = self.ERROR_COL

        # If it's just a temporary file via Microsoft, end here
        if (itemName[0:2] == "~$"):
            return errorPresent

        # TODO: figure out if 200 is inclusive or exclusive. Safe bet is inclusive
        absoluteItemLength = len(dirAbsolute + "/" + itemName)
        if (absoluteItemLength >= 200):
            errorPresent = True
            self.mainSheet.write(self.mainSheetRow, variableErrorCol, "{} chars >= 200. Terminating checks.".format(absoluteItemLength))
            # variableErrorCol += 1
            return errorPresent

        errorChars = set()
        periodCount = 0
        itemNameLength = len(itemName)

        for i in range(itemNameLength):
            if itemName[i] not in self.permissibleCharacters:
                # Not necessary to set errorPresent because we're returning errorChars instead
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
            # Surrounding the set with "||" makes space chars visible while not running a check each loop to change " " (SPC) to something human-readable
            self.mainSheet.write(self.mainSheetRow, variableErrorCol, 
            "Bad chars: |{}|".format("".join(errorChars)))

            # Only returns a set if it's populated
            return errorChars

        return errorPresent



    def checkPCPAECNamingConventionRename(self, dirAbsolute:str , itemName: str) -> bool:
        result = self.checkPCPAECNamingConvention(dirAbsolute, itemName)

        if (isinstance(result, bool)):
            return result
        
        ### Change spaces and double dashes into dashes
        if (not {" ", "-"}.isdisjoint(result)):
            # Replace "-" characters with " " to make the string homogenous for the upcoming split()
            # split() automatically removes leading, trailing, and excess middle whitespace
            newItemName = itemName.replace("-", " ").split()
            newItemName = "-".join(newItemName)

            # For if there's a dash to the left of the file extension period
            lastPeriodIndex = newItemName.rfind(".")
            # If lastPeriodIndex isn't the very first character and there actually is a period
            if (lastPeriodIndex > 0 and newItemName[lastPeriodIndex -1] == "-"):
                newItemName = newItemName[0:lastPeriodIndex-1] + newItemName[lastPeriodIndex:]

            # Log the new name
            try:
                os.rename(dirAbsolute + "/" + itemName, dirAbsolute + "/" + newItemName)
                self.mainSheet.write(self.mainSheetRow, self.RENAME_COL, newItemName, self.renameFormat)
            except PermissionError:
                self.mainSheet.write(self.mainSheetRow, self.RENAME_COL, "FILE LOCKED. RENAME FAILED.", self.fileErrorFormat)
            except OSError:
                self.mainSheet.write(self.mainSheetRow, self.RENAME_COL, "OS ERROR. RENAME FAILED.", self.fileErrorFormat)


        return True



    def close(self):
        self.wb.close()