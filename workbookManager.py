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
        self.concurrentCheckSheetsList = []
        self.postCheckSheetsList = []
        self.fixSheet = None # initializing for clarity

        self.concurrentCheckMethods = []  # List of checkMethods to run on the fileName. Called within fileCrawl()
        self.miscCheckMethods = []
        self.postCheckMethods = []
        self.fixMethod = lambda _1, _2, _3: None  # default dummy function
        self.folderFixMethod = lambda _1, _2, _3, _4: None
        self.fixArg = None  # initialize fixArg to a dummy value

        # For summarySheet first 7 rows are used for other stats. Skip a line, then write variable number of errors.
        self.sheetRow = {self.summarySheet: 8}
        self.checkSheetFileCount = {}

        # Summary sheet metrics
        self.filesScannedCount = 0
        self.executionTime = 0
        self.errorCount = 0

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


    def addConcurrentCheckMethod(self, ccmName:str, functionSelection: Callable[[str, str], bool]):
        """Adds a worksheet and corresponding checkMethod"""

        self.concurrentCheckMethods.append(functionSelection)
        tmpWsVar = self.wb.add_worksheet(ccmName)

        self.summarySheet.write(self.sheetRow[self.summarySheet] +len(self.concurrentCheckSheetsList) +len(self.postCheckSheetsList), 0, ccmName + " count", self.headerFormat)
        self.concurrentCheckSheetsList.append(tmpWsVar)
        self.sheetRow[tmpWsVar] = 1
        self.checkSheetFileCount[tmpWsVar] = 0


    def addMiscCheckMethod(self, mcmName:str, functionSelection: Callable):
        self.miscCheckMethods.append(functionSelection)


    def addPostCheckMethod(self, pcmName:str, functionSelection: Callable):
        self.postCheckMethods.append(functionSelection)

        self.summarySheet.write(self.sheetRow[self.summarySheet] +len(self.concurrentCheckSheetsList) +len(self.postCheckSheetsList), 0, pcmName + " count", self.headerFormat)
        tmpWsVar = self.wb.add_worksheet(pcmName)
        self.postCheckSheetsList.append(tmpWsVar)
        self.checkSheetFileCount[tmpWsVar] = 0
        self.sheetRow[tmpWsVar] = 1



    # TODO: Create a generic function for the following 2 fix method setting functions
    def setFixMethod(self, wsName, functionSelection: Callable[[str, str], bool]):
        self.summarySheet.write(self.sheetRow[self.summarySheet] +len(self.concurrentCheckSheetsList) +len(self.postCheckSheetsList), 0, wsName + " count", self.headerFormat)

        tmpWsVar = self.wb.add_worksheet(wsName)
        self.fixSheet = tmpWsVar
        self.sheetRow[tmpWsVar] = 1
        self.checkSheetFileCount[tmpWsVar] = 0

        self.fixMethod = functionSelection


    def setFolderFixMethod(self, wsName, functionSelection):
        self.summarySheet.write(self.sheetRow[self.summarySheet] +len(self.concurrentCheckSheetsList) +len(self.postCheckSheetsList), 0, wsName + " count", self.headerFormat)

        tmpWsVar = self.wb.add_worksheet(wsName)
        self.fixSheet = tmpWsVar
        self.sheetRow[tmpWsVar] = 1
        self.checkSheetFileCount[tmpWsVar] = 0

        self.folderFixMethod = functionSelection

        

    def setFixArg(self, arg: int):
        self.fixArg = arg


    # Lambdas are automatically considered bad functions
    def isFixSheetSet(self) -> bool:
        return bool(self.fixSheet)  # if self.fixSheet is not set to its default value "None"

        # The following line is specifically for checking fixMethod being set to a valid function, and *not* a lambda
        # return callable(self.fixMethod) and self.fixMethod.__name__ != "<lambda>"


    def setDefaultFormatting(self, dirAbsolute, includeSubFolders, renamingFiles):
        for wsc in self.concurrentCheckSheetsList:
            # Column width
            # wsc.set_column(self.DIR_COL, self.DIR_COL, 50)
            # wsc.set_column(self.ITEM_COL, self.ITEM_COL, 30)
            # wsc.set_column(self.ERROR_COL, self.ERROR_COL, 20)
            wsc.freeze_panes(1, 0)

            wsc.write(0, self.DIR_COL, "Directories", self.headerFormat)
            wsc.write(0, self.ITEM_COL, "Items", self.headerFormat)
            wsc.write(0, self.ERROR_COL, "Error", self.headerFormat)


        if (self.isFixSheetSet()):
            self.fixSheet.freeze_panes(1, 0)
            self.fixSheet.write(0, self.DIR_COL, "Directories", self.headerFormat)
            self.fixSheet.write(0, self.ITEM_COL, "Items", self.headerFormat)
            self.fixSheet.write(0, self.ERROR_COL, "Modification" if renamingFiles else "Potential Modification", self.headerFormat)


        self.summarySheet.set_column(0, 0, 20)
        self.summarySheet.set_column(1, 1, 15)
        
        self.summarySheet.write(0, 0, "FilePath", self.headerFormat)
        self.summarySheet.write(1, 0, "SubFolder inclusion", self.headerFormat)
        self.summarySheet.write(2, 0, "Renaming", self.headerFormat)
        self.summarySheet.write(3, 0, "Argument", self.headerFormat)
        self.summarySheet.write(4, 0, "File count", self.headerFormat)
        self.summarySheet.write(5, 0, "Error count / %", self.headerFormat)
        self.summarySheet.write(6, 0, "Execution time (s)", self.headerFormat)

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

            # Run every selected checkMethod
            # Concurrent check methods
            for i in range(len(self.concurrentCheckMethods)):
                # If error present
                if (self.concurrentCheckMethods[i](dirAbsolute, itemName, self.concurrentCheckSheetsList[i])):
                    if (not alreadyCounted):
                        self.errorCount += 1
                        alreadyCounted = True

            alreadyCounted = False

            # Misc check methods
            for mcm in self.miscCheckMethods:
                mcm(dirAbsolute, itemName)
            
            # Fix method
            self.filesScannedCount += 1
            self.fixMethod(dirAbsolute, itemName, self.fixSheet)
                
                    

    def folderCrawl(self, dirTree: List[Tuple[str, list, list]]):
        start = time()

        allSheets = self.concurrentCheckSheetsList[:]
        if (self.isFixSheetSet()):
            allSheets.append(self.fixSheet)   
        
        for (dirAbsolute, dirFolders, dirFiles) in dirTree:
            for ws in allSheets:
                ws.write(self.sheetRow[ws], self.DIR_COL, dirAbsolute, self.dirColFormat)

            self.fileCrawl(dirAbsolute, dirFiles)

            # Folder fix method
            self.folderFixMethod(dirAbsolute, dirFolders, dirFiles, self.fixSheet)


        for i in range(len(self.postCheckMethods)):
            self.postCheckMethods[i](self.postCheckSheetsList[i])

        end = time()
        self.executionTime = end - start


    def writeInCell(self, ws, col: str, text: str, format=None, rowIncrement=0, fileIncrement=0):
        # write_string() usage so no equations are accidentally written
        if (format): ws.write_string(self.sheetRow[ws], col, text, format)
        else: ws.write_string(self.sheetRow[ws], col, text)

        self.sheetRow[ws] += rowIncrement
        self.checkSheetFileCount[ws] += fileIncrement



    def incrementRow(self, ws, amount:int=1):
        self.sheetRow[ws] += amount

    def incrementFileCount(self, ws, amount:int=1):
        self.checkSheetFileCount[ws] += amount



    def populateSummarySheet(self):
        if self.filesScannedCount == 0: errorPercentage = 0
        else: errorPercentage = round(self.errorCount / self.filesScannedCount * 100, 2)

        if (self.fixArg != None): self.summarySheet.write_number(3, 1, self.fixArg, self.summaryValueFormat)
        
        self.summarySheet.write_number(4, 1, self.filesScannedCount, self.summaryValueFormat)
        self.summarySheet.write_number(5, 1, self.errorCount, self.summaryValueFormat)
        self.summarySheet.write(5, 2, "{}%".format(errorPercentage), self.summaryValueFormat)
        self.summarySheet.write_number(6, 1, round(self.executionTime, 4), self.summaryValueFormat)
        
        i = 0
        for ws in (self.concurrentCheckSheetsList + self.postCheckSheetsList):
            self.summarySheet.write(self.sheetRow[self.summarySheet] + i, 1, self.checkSheetFileCount[ws], self.summaryValueFormat)
            i += 1

        if (self.isFixSheetSet()):
            self.summarySheet.write(self.sheetRow[self.summarySheet] + i, 1, self.checkSheetFileCount[self.fixSheet], self.summaryValueFormat)

        

    def createHelpMeSheet(self):
        # attempt to open HELPME.txt
        try: helpMeFile = open("HELPME.txt", "r")
        except FileNotFoundError: return
        
        termDict = {}
        lines = helpMeFile.readlines()
        linesLength = len(lines)

        i = 0
        while i < linesLength:
            if lines[i][0] == "-":
                term = lines[i][1:].strip()
                definition = lines[i+1].strip()
                termDict[term] = definition
                i += 1
            i += 1

        helpMeSheet = self.wb.add_worksheet("HelpMe")
        helpMeSheet.write(0, 0, "Term", self.headerFormat)
        helpMeSheet.write(0, 1, "Definition", self.headerFormat)

        row = 1
        for termKey in termDict:
            helpMeSheet.write(row, 0, termKey)
            helpMeSheet.write(row, 1, termDict[termKey])
            row += 1

        helpMeSheet.autofit()



    def autofitSheets(self):
        # summarySheet is not autofit as the filepath is much much wider than everything else. It is set manually at the beginning.
        # Note that autofit only functions as intended if done after the data has been entered

        allSheets = self.concurrentCheckSheetsList[:]
        if (self.isFixSheetSet()):
            allSheets.append(self.fixSheet)  
        
        for ws in allSheets:
            ws.autofit()
            


    def close(self):
        self.populateSummarySheet()
        self.autofitSheets()
        self.createHelpMeSheet()
        self.wb.close()
