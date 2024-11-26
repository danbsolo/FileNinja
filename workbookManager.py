import xlsxwriter
from typing import List, Tuple, Callable
from time import time

class WorkbookManager:

    def __init__(self, workbookPathName: str):
        self.workbookPathName = workbookPathName
        self.wb = xlsxwriter.Workbook(workbookPathName)
        
        self.mainSheet = self.wb.add_worksheet("Main")
        self.summarySheet = self.wb.add_worksheet("Summary")
        self.summarySheet.activate()  # view this worksheet on startup

        # Freeze header row
        self.mainSheet.freeze_panes(1, 0)

        # For if more sheets are used
        self.sheetRow = {}
        self.sheetRow[self.mainSheet] = 1

        # Summary sheet metrics
        self.filesScannedCount = 0
        self.executionTime = 0
        self.errorCount = 0

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

        # SummarySheet columns are set using autofit() *after* its cells have been populated

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


        self.summaryValueFormat = self.wb.add_format({
            "bold": True
        })


        # Write headers
        self.mainSheet.write(0, self.DIR_COL, "Directories", self.headerFormat)
        self.mainSheet.write(0, self.ITEM_COL, "Items", self.headerFormat)
        self.mainSheet.write(0, self.RENAME_COL, "Rename", self.headerFormat)
        self.mainSheet.write(0, self.ERROR_COL, "Errors", self.headerFormat)

        self.summarySheet.write(0, 0, "File count", self.headerFormat)
        self.summarySheet.write(1, 0, "Error count", self.headerFormat)
        self.summarySheet.write(2, 0, "Error percentage (%)", self.headerFormat)
        self.summarySheet.write(3, 0, "Execution time (s)", self.headerFormat)



    def fileCrawl(self, dirAbsolute, dirItems: List[str]):
        # can be used to write down folders as well, hence "items" and not "files"
        for itemName in dirItems:

            # If error present
            if (self.checkMethod(dirAbsolute, itemName)):
                self.mainSheet.write(self.sheetRow[self.mainSheet], self.ITEM_COL, itemName, self.fileErrorFormat)
                
                self.errorCount += 1
                self.sheetRow[self.mainSheet] += 1
            else:
                pass
                # self.mainSheet.write(self.sheetRow[self.mainSheet], self.ITEM_COL, itemName)


    


    def folderCrawl(self, dirTree: List[Tuple[str, list, list]]):
        start = time()
        
        for (dirAbsolute, dirFolders, dirFiles) in dirTree:
            self.mainSheet.write(self.sheetRow[self.mainSheet], self.DIR_COL, dirAbsolute, self.dirColFormat)
            self.fileCrawl(dirAbsolute, dirFiles)

            self.filesScannedCount += len(dirFiles)
        
        end = time()
        self.executionTime = end - start



    def setCheckMethod(self, functionSelection: Callable[[str, str], bool]):
        self.checkMethod = functionSelection



    def writeInCell(self, sheet, col: str, text: str, format=False):
        if (format): sheet.write(self.sheetRow[sheet], col, text, format)
        else: sheet.write(self.sheetRow[sheet], col, text)



    def populateSummarySheet(self):
        errorPercentage = round(self.errorCount / self.filesScannedCount * 100)

        self.summarySheet.write_number(0, 1, self.filesScannedCount, self.summaryValueFormat)
        self.summarySheet.write_number(1, 1, self.errorCount, self.summaryValueFormat)
        self.summarySheet.write_number(2, 1, round(errorPercentage, 4), self.summaryValueFormat)
        self.summarySheet.write_number(3, 1, round(self.executionTime, 4), self.summaryValueFormat)
        
        self.summarySheet.autofit()


    def close(self):
        self.populateSummarySheet()
        self.wb.close()
