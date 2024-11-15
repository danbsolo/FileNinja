import xlsxwriter
from typing import List, Tuple, Callable


class WorkbookManager:

    def __init__(self, workbookName: str):
        self.workbookName = workbookName
        self.wb = xlsxwriter.Workbook(workbookName)
        
        # TODO: add more worksheets here if/when necessary
        self.mainSheet = self.wb.add_worksheet("MainSheet")
        
        # Freeze header row
        self.mainSheet.freeze_panes(1, 0)
        
        # For if more sheets are used
        self.sheetRow = {}
        self.sheetRow[self.mainSheet] = 1

        # An empty function by default. Called within fileCrawl()
        self.checkMethod = lambda: None



    def setDefault(self):
        self.setColumnIndexes()
        self.setDefaultFormatting()



    def setColumnIndexes(self, dirCol:int=0, itemCol:int=1, renameCol:int=2, errorCol:int=3):
        self.DIR_COL = dirCol
        self.ITEM_COL = itemCol
        self.RENAME_COL = renameCol
        self.ERROR_COL = errorCol



    def setDefaultFormatting(self):
        # TODO: user should have the ability to set most (or all) of the variables hereafter
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
            "bg_color": "#00FF80", # greenish
            "bold": True
        })

        self.showRenameFormat = self.wb.add_format({
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
        # can be used to write down folders as well, hence "items" and not "files"
        for itemName in dirItems:
            
            # If error present
            if (self.checkMethod(dirAbsolute, itemName)):
                self.mainSheet.write(self.sheetRow[self.mainSheet], self.ITEM_COL, itemName, self.fileErrorFormat)
            else:
                self.mainSheet.write(self.sheetRow[self.mainSheet], self.ITEM_COL, itemName)

            self.sheetRow[self.mainSheet] += 1
    


    def folderCrawl(self, dirTree: List[Tuple[str, list, list]]):
        for (dirAbsolute, dirFolders, dirFiles) in dirTree:
            self.mainSheet.write(self.sheetRow[self.mainSheet], self.DIR_COL, dirAbsolute, self.dirColFormat)
            self.fileCrawl(dirAbsolute, dirFolders + dirFiles)



    def setCheckMethod(self, functionSelection: Callable[[str, str], bool]):
        self.checkMethod = functionSelection



    def writeInCell(self, sheet, col: str, text: str, format=False):
        if (format): sheet.write(self.sheetRow[sheet], col, text, format)
        else: sheet.write(self.sheetRow[sheet], col, text)



    def close(self):
        self.wb.close()
