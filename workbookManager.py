import xlsxwriter
from typing import List, Tuple, Callable
from time import time

class WorkbookManager:

    def __init__(self, workbookPathName: str):
        self.workbookPathName = workbookPathName
        self.wb = xlsxwriter.Workbook(workbookPathName)

        self.summarySheet = self.wb.add_worksheet("Summary")
        self.summarySheet.activate()  # view this worksheet on startup

        # Doesn't add summarySheet
        self.worksheetsList = []

        # For summarySheet first 3 rows are used for other stats. Skip a line, then write variable errors.
        self.sheetRow = {self.summarySheet: 4}
        self.checkSheetErrorCount = {}

        # Summary sheet metrics
        self.filesScannedCount = 0
        self.executionTime = 0
        self.errorCount = 0

        # List of checkMethods to run on the fileName. Called within fileCrawl()
        self.checkMethods = []

        self.DIR_COL = 0
        self.ITEM_COL = 1
        self.RENAME_COL = 2
        self.ERROR_COL = 3

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

        # Nothing so far
        self.summaryValueFormat = self.wb.add_format({
        })



    def addCheckSheet(self, wsName, functionSelection: Callable[[str, str], bool]):
        """Adds a worksheet and corresponding checkMethod"""
        tmpWsVar = self.wb.add_worksheet(wsName)
        self.worksheetsList.append(tmpWsVar)
        self.sheetRow[tmpWsVar] = 1
        self.checkSheetErrorCount[tmpWsVar] = 0

        self.checkMethods.append(functionSelection)

        self.summarySheet.write(self.sheetRow[self.summarySheet], 0, wsName + " count", self.headerFormat)
        self.sheetRow[self.summarySheet] += 1



    def setDefaultFormatting(self):
        for ws in self.worksheetsList:
            # Column width
            ws.set_column(self.DIR_COL, self.DIR_COL, 50)
            ws.set_column(self.ITEM_COL, self.RENAME_COL, 30)
            ws.set_column(self.ERROR_COL, self.ERROR_COL, 20)
            ws.freeze_panes(1, 0)

            # Write headers
            ws.write(0, self.DIR_COL, "Directories", self.headerFormat)
            ws.write(0, self.ITEM_COL, "Items", self.headerFormat)

            # TODO: Change this title dynamically based on whether user chose to rename or not rename
            ws.write(0, self.RENAME_COL, "Potential Rename / Renamed", self.headerFormat)
            ws.write(0, self.ERROR_COL, "Error", self.headerFormat)


        # SummarySheet columns are set using autofit() *after* its cells have been populated. Check populateSummarySheet()
        
        self.summarySheet.write(0, 0, "File count", self.headerFormat)
        self.summarySheet.write(1, 0, "Error count / %", self.headerFormat)
        self.summarySheet.write(2, 0, "Execution time (s)", self.headerFormat)



    def fileCrawl(self, dirAbsolute, dirItems: List[str]):
        # For ensuring a file already counted isn't counted more than once, even if it has multiple errors
        alreadyCounted = False

        # can be used to write down folders as well, hence "items" and not "files"
        for itemName in dirItems:
            
            # If it's just a temporary file via Microsoft, skip file
            if (itemName[0:2] == "~$"):
                continue

            # Run every selected checkMethod on itemName
            for i in range(len(self.checkMethods)):
                # If error present
                if (self.checkMethods[i](dirAbsolute, itemName, self.worksheetsList[i])):
                    if (not alreadyCounted):
                        self.errorCount += 1
                        alreadyCounted = True
            
            alreadyCounted = False
                    


    def folderCrawl(self, dirTree: List[Tuple[str, list, list]]):
        start = time()
        
        for (dirAbsolute, dirFolders, dirFiles) in dirTree:
            for ws in self.worksheetsList:
                ws.write(self.sheetRow[ws], self.DIR_COL, dirAbsolute, self.dirColFormat)

            self.fileCrawl(dirAbsolute, dirFiles)

            self.filesScannedCount += len(dirFiles)
        
        end = time()
        self.executionTime = end - start



    def writeInCell(self, sheet, col: str, text: str, format=False):
        # write_string so no equations are accidentally written
        if (format): sheet.write_string(self.sheetRow[sheet], col, text, format)
        else: sheet.write_string(self.sheetRow[sheet], col, text)



    def populateSummarySheet(self):
        errorPercentage = round(self.errorCount / self.filesScannedCount * 100, 2)

        self.summarySheet.write_number(0, 1, self.filesScannedCount, self.summaryValueFormat)
        self.summarySheet.write_number(1, 1, self.errorCount, self.summaryValueFormat)
        self.summarySheet.write(1, 2, "{}%".format(errorPercentage), self.summaryValueFormat)
        self.summarySheet.write_number(2, 1, round(self.executionTime, 4), self.summaryValueFormat)
        
        i = 0
        for ws in self.worksheetsList:
            self.summarySheet.write(4+i, 1, self.checkSheetErrorCount[ws], self.summaryValueFormat)
            i += 1

        self.summarySheet.autofit()



    def close(self):
        self.populateSummarySheet()
        self.wb.close()
