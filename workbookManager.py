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
        self.checkSheetsList = []
        self.fixSheetsList = []

        # For summarySheet first 3 rows are used for other stats. Skip a line, then write variable errors.
        self.sheetRow = {self.summarySheet: 7}
        self.checkSheetErrorCount = {}

        # Summary sheet metrics
        self.filesScannedCount = 0
        self.executionTime = 0
        self.errorCount = 0

        # List of checkMethods to run on the fileName. Called within fileCrawl()
        self.checkMethods = []
        self.fixMethods = []

        self.DIR_COL = 0
        self.ITEM_COL = 1
        self.RENAME_COL = 2
        self.ERROR_COL = 2

        # Default cell styles
        self.dirColFormat = self.wb.add_format({
            "bg_color": "#99CCFF", # blueish
            "bold": True
        })

        # "bg_color": "#FF4444", # reddish
        self.fileErrorFormat = self.wb.add_format({
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
        self.summarySheet.write(self.sheetRow[self.summarySheet] +len(self.checkSheetsList), 0, wsName + " count", self.headerFormat)

        tmpWsVar = self.wb.add_worksheet(wsName)
        self.checkSheetsList.append(tmpWsVar)
        self.sheetRow[tmpWsVar] = 1
        self.checkSheetErrorCount[tmpWsVar] = 0

        self.checkMethods.append(functionSelection)


    def addFixSheet(self, wsName, functionSelection: Callable[[str, str], bool]):
        self.summarySheet.write(self.sheetRow[self.summarySheet] +len(self.checkSheetsList), 0, wsName + " count", self.headerFormat)

        tmpWsVar = self.wb.add_worksheet(wsName)
        self.fixSheetsList.append(tmpWsVar)
        self.sheetRow[tmpWsVar] = 1
        self.checkSheetErrorCount[tmpWsVar] = 0

        self.fixMethods.append(functionSelection)
        


    def setDefaultFormatting(self, dirAbsolute, includeSubFolders, renamingFiles):
        for wsc in self.checkSheetsList:
            # Column width
            # wsc.set_column(self.DIR_COL, self.DIR_COL, 50)
            # wsc.set_column(self.ITEM_COL, self.ITEM_COL, 30)
            # wsc.set_column(self.ERROR_COL, self.ERROR_COL, 20)

            wsc.freeze_panes(1, 0)

            # Write headers
            wsc.write(0, self.DIR_COL, "Directories", self.headerFormat)
            wsc.write(0, self.ITEM_COL, "Items", self.headerFormat)
            wsc.write(0, self.ERROR_COL, "Error", self.headerFormat)


        if (renamingFiles):
            renameColName = "Renamed"
        else:
            renameColName = "Potential Rename"

        
        for wsf in self.fixSheetsList:
            # Column width
            # wsf.set_column(self.DIR_COL, self.DIR_COL, 50)
            # wsf.set_column(self.ITEM_COL, self.RENAME_COL, 30)

            wsf.freeze_panes(1, 0)
            
            # Write headers
            wsf.write(0, self.DIR_COL, "Directories", self.headerFormat)
            wsf.write(0, self.ITEM_COL, "Items", self.headerFormat)
            wsf.write(0, self.ERROR_COL, renameColName, self.headerFormat)
            


        self.summarySheet.set_column(0, 0, 20)
        self.summarySheet.set_column(1, 1, 15)
        
        self.summarySheet.write(0, 0, "FilePath", self.headerFormat)
        self.summarySheet.write(1, 0, "SubFolder inclusion", self.headerFormat)
        self.summarySheet.write(2, 0, "Renaming", self.headerFormat)
        self.summarySheet.write(3, 0, "File count", self.headerFormat)
        self.summarySheet.write(4, 0, "Error count / %", self.headerFormat)
        self.summarySheet.write(5, 0, "Execution time (s)", self.headerFormat)

        self.summarySheet.write(0, 1, dirAbsolute, self.summaryValueFormat)
        self.summarySheet.write(1, 1, str(includeSubFolders), self.summaryValueFormat)
        self.summarySheet.write(2, 1, str(renamingFiles), self.summaryValueFormat)
        


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
                if (self.checkMethods[i](dirAbsolute, itemName, self.checkSheetsList[i])):
                    if (not alreadyCounted):
                        self.errorCount += 1
                        alreadyCounted = True
            
            alreadyCounted = False

            for i in range(len(self.fixMethods)):
                self.fixMethods[i](dirAbsolute, itemName, self.fixSheetsList[i]);
                
                    

    def folderCrawl(self, dirTree: List[Tuple[str, list, list]]):
        start = time()
        
        for (dirAbsolute, dirFolders, dirFiles) in dirTree:
            for ws in self.checkSheetsList + self.fixSheetsList:
                ws.write(self.sheetRow[ws], self.DIR_COL, dirAbsolute, self.dirColFormat)

            self.fileCrawl(dirAbsolute, dirFiles)

            self.filesScannedCount += len(dirFiles)

        # If no erred files are under the last directory, it still gets printed
        # A fix for that would have to be here, if necessary


        end = time()
        self.executionTime = end - start



    def writeInCell(self, ws, col: str, text: str, format=False, rowIncrement=0, errorIncrement=0):
        # write_string so no equations are accidentally written
        if (format): 
            ws.write_string(self.sheetRow[ws], col, text, format)
        else: 
            ws.write_string(self.sheetRow[ws], col, text)

        self.sheetRow[ws] += rowIncrement
        self.checkSheetErrorCount[ws] += errorIncrement



    def populateSummarySheet(self):
        errorPercentage = round(self.errorCount / self.filesScannedCount * 100, 2)

        self.summarySheet.write_number(3, 1, self.filesScannedCount, self.summaryValueFormat)
        self.summarySheet.write_number(4, 1, self.errorCount, self.summaryValueFormat)
        self.summarySheet.write(4, 2, "{}%".format(errorPercentage), self.summaryValueFormat)
        self.summarySheet.write_number(5, 1, round(self.executionTime, 4), self.summaryValueFormat)
        
        i = 0
        for ws in self.checkSheetsList:
            self.summarySheet.write(self.sheetRow[self.summarySheet] + i, 1, self.checkSheetErrorCount[ws], self.summaryValueFormat)
            ws.autofit()
            i += 1
        
        i = 0
        for ws in self.fixSheetsList:
            self.summarySheet.write(self.sheetRow[self.summarySheet] + i + len(self.checkSheetsList), 1, self.checkSheetErrorCount[ws], self.summaryValueFormat)
            ws.autofit()
            i += 1


    def close(self):
        self.populateSummarySheet()
        self.wb.close()
