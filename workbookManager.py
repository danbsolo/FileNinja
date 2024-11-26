import xlsxwriter
from typing import List, Tuple, Callable
from time import time

class WorkbookManager:

    def __init__(self, workbookPathName: str):
        self.workbookPathName = workbookPathName
        self.wb = xlsxwriter.Workbook(workbookPathName)

        self.spaceErrorSheet = self.wb.add_worksheet("SPC-Error")
        self.charLimitErrorSheet = self.wb.add_worksheet("CharLimit-Error")
        self.badCharErrorSheet = self.wb.add_worksheet("BadChar-Error")
        self.summarySheet = self.wb.add_worksheet("Summary")
        self.summarySheet.activate()  # view this worksheet on startup

        # Freeze header row
        self.spaceErrorSheet.freeze_panes(1, 0)
        self.charLimitErrorSheet.freeze_panes(1, 0)
        self.badCharErrorSheet.freeze_panes(1, 0)

        # For if more sheets are used
        self.sheetRow = {}
        self.sheetRow[self.spaceErrorSheet] = 1
        self.sheetRow[self.charLimitErrorSheet] = 1
        self.sheetRow[self.badCharErrorSheet] = 1

        # Summary sheet metrics
        self.filesScannedCount = 0
        self.executionTime = 0
        self.errorCount = 0

        # List of checkMethods to run on the fileName. Called within fileCrawl()
        self.checkMethods = []
        # lambda: None <-- empty method


    def setDefault(self):
        self.setColumnIndexes()
        self.setDefaultFormatting()



    def setColumnIndexes(self, dirCol:int=0, itemCol:int=1, renameCol:int=2, errorCol:int=3):
        self.DIR_COL = dirCol
        self.ITEM_COL = itemCol
        self.RENAME_COL = renameCol
        self.ERROR_COL = errorCol



    def setDefaultFormatting(self):
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


        for ws in [self.spaceErrorSheet, self.charLimitErrorSheet, self.badCharErrorSheet]:
            # Column width
            ws.set_column(self.DIR_COL, self.DIR_COL, 50)
            ws.set_column(self.ITEM_COL, self.RENAME_COL, 30)
            ws.set_column(self.ERROR_COL, self.ERROR_COL, 20)

            # SummarySheet columns are set using autofit() *after* its cells have been populated

            # Write headers
            ws.write(0, self.DIR_COL, "Directories", self.headerFormat)
            ws.write(0, self.ITEM_COL, "Items", self.headerFormat)
            ws.write(0, self.RENAME_COL, "Potential Rename / Renamed", self.headerFormat)
            ws.write(0, self.ERROR_COL, "Errors", self.headerFormat)


        self.summarySheet.write(0, 0, "File count", self.headerFormat)
        self.summarySheet.write(1, 0, "Error count", self.headerFormat)
        self.summarySheet.write(2, 0, "Error percentage (%)", self.headerFormat)
        self.summarySheet.write(3, 0, "Execution time (s)", self.headerFormat)



    def fileCrawl(self, dirAbsolute, dirItems: List[str]):
        # can be used to write down folders as well, hence "items" and not "files"
        for itemName in dirItems:
            
            # If it's just a temporary file via Microsoft, skip file
            if (itemName[0:2] == "~$"):
                continue

            # Run every selected checkMethod on itemName
            for cm in self.checkMethods:
                # If error present
                if (cm(dirAbsolute, itemName)):
                    self.errorCount += 1
                    


    def folderCrawl(self, dirTree: List[Tuple[str, list, list]]):
        start = time()
        
        for (dirAbsolute, dirFolders, dirFiles) in dirTree:
            for ws in [self.spaceErrorSheet, self.charLimitErrorSheet, self.badCharErrorSheet]:
                ws.write(self.sheetRow[ws], self.DIR_COL, dirAbsolute, self.dirColFormat)

            self.fileCrawl(dirAbsolute, dirFiles)

            self.filesScannedCount += len(dirFiles)
        
        end = time()
        self.executionTime = end - start



    def appendCheckMethod(self, functionSelection: Callable[[str, str], bool]):
        self.checkMethods.append(functionSelection)



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
