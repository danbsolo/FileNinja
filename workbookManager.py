import xlsxwriter
from typing import List, Tuple, Callable
from time import time

class WorkbookManager:

    def __init__(self, workbookPathName: str):
        self.workbookPathName = workbookPathName
        self.wb = xlsxwriter.Workbook(workbookPathName)

        self.summarySheet = self.wb.add_worksheet("Summary")
        self.summarySheet.activate()  # view this worksheet on startup

        self.fixSheet = None # initializing for clarity
        self.fileFixProcedure = lambda _1, _2, _3: None  # default dummy function
        self.folderFixProcedure = lambda _1, _2, _3, _4: None
        self.fixArg = None  # initialize fixArg to a dummy value

        self.findSheets = {} # procedureObject : worksheet
        # For summarySheet first 7 rows are used for mainstay metrics. Skip a line, then write variable number of errors.
        self.sheetRows = {self.summarySheet: 8} # worksheet : Integer
        self.summaryCounts = {}  # worksheet : Integer

        # Summary metrics
        self.filesScannedCount = 0
        self.executionTime = 0
        self.errorCount = 0

        self.DIR_COL = 0
        self.ITEM_COL = 1
        # used describe either an Error or a Modification, depending on the procedure type
        self.OUTCOME_COL = 2

        # Default cell styles
        self.dirColFormat = self.wb.add_format({
            "bg_color": "#99CCFF", # blueish
            "bold": True
        })

        # "bg_color": "#FF4444", # reddish
        self.errorFormat = self.wb.add_format({
            **self.defaultFormat.get_properties(),
            "bold": True
        })

        self.modifyFormat = self.wb.add_format({
            "bg_color": "#00FF80", # greenish
            "bold": True
        })

        self.logFormat = self.wb.add_format({
            "bg_color": "#9999FF", # purplish
            "bold": True
        })

        self.headerFormat = self.wb.add_format({
            "bg_color": "#C0C0C0", # grayish
            "bold": True
        })

        self.summaryValueFormat = self.wb.add_format({
        })


    def getAllProcedureSheets(self):
        sheets = list(self.findSheets.values())
        if self.fixSheet:
            sheets.append(self.fixSheet)
        return sheets


    def addFindProcedure(self, findProcedureObject):
        tmpWsVar = self.wb.add_worksheet(findProcedureObject.name)
        self.summarySheet.write(self.sheetRows[self.summarySheet] +len(self.findSheets.keys()), 0, findProcedureObject.name + " count", self.headerFormat)
        
        self.findSheets[findProcedureObject] = tmpWsVar
        self.sheetRows[tmpWsVar] = 1
        self.summaryCounts[tmpWsVar] = 0

        if findProcedureObject.isStateless:
            tmpWsVar.freeze_panes(1, 0)
            tmpWsVar.write(0, self.DIR_COL, "Directories", self.headerFormat)
            tmpWsVar.write(0, self.ITEM_COL, "Items", self.headerFormat)
            tmpWsVar.write(0, self.OUTCOME_COL, "Error", self.headerFormat)


    def setFixProcedure(self, fixProcedureObject, modify):
        self.summarySheet.write(self.sheetRows[self.summarySheet] +len(self.findSheets.keys()), 0, fixProcedureObject.name + " count", self.headerFormat)
        tmpWsVar = self.wb.add_worksheet(fixProcedureObject.name)
        self.fixSheet = tmpWsVar
        self.sheetRows[tmpWsVar] = 1
        self.summaryCounts[tmpWsVar] = 0

        if fixProcedureObject.isFileFix:
            if modify: self.fileFixProcedure = fixProcedureObject.modifyFunction
            else: self.fileFixProcedure = fixProcedureObject.logFunction
        else:
            if modify: self.folderFixProcedure = fixProcedureObject.modifyFunction
            else: self.folderFixProcedure = fixProcedureObject.logFunction

        self.fixSheet.freeze_panes(1, 0)
        self.fixSheet.write(0, self.DIR_COL, "Directories", self.headerFormat)
        self.fixSheet.write(0, self.ITEM_COL, "Items", self.headerFormat)
        self.fixSheet.write(0, self.OUTCOME_COL, fixProcedureObject.columnName, self.headerFormat)


    def setFixArg(self, fixProcedureObject, unprocessedArg) -> bool:
        if not fixProcedureObject.validatorFunction:
            return True 
        if ((arg := fixProcedureObject.validatorFunction(unprocessedArg, fixProcedureObject.argBoundary)) is None):
            return False

        self.fixArg = arg
        return True
        

    def fileCrawl(self, dirAbsolute, dirItems: List[str]):
        alreadyCounted = False

        for itemName in dirItems:            
            # If it's just a temporary file via Microsoft, skip file
            if (itemName[0:2] == "~$"):
                continue

            self.fileFixProcedure(dirAbsolute, itemName, self.fixSheet)

            for findProcedureObject in self.findSheets.keys():
                if (findProcedureObject.mainFunction(dirAbsolute, itemName, self.findSheets[findProcedureObject])):
                    if (not alreadyCounted):
                        self.errorCount += 1
                        alreadyCounted = True

            alreadyCounted = False
            self.filesScannedCount += 1                    


    def folderCrawl(self, dirTree: List[Tuple[str, list, list]]):
        start = time()

        allSheets = self.getAllProcedureSheets() 
        
        for (dirAbsolute, dirFolders, dirFiles) in dirTree:
            for ws in allSheets:
                ws.write(self.sheetRows[ws], self.DIR_COL, dirAbsolute, self.dirColFormat)

            self.fileCrawl(dirAbsolute, dirFiles)

            # Folder fix procedure
            self.folderFixProcedure(dirAbsolute, dirFolders, dirFiles, self.fixSheet)

        for findProcedureObject in self.findSheets.keys():
            if findProcedureObject.postFunction:
                findProcedureObject.postFunction(self.findSheets[findProcedureObject])

        end = time()
        self.executionTime = end - start


    def writeItem(self, ws, text, format=None):
        self.writeHelper(ws, self.ITEM_COL, text, format)

    def writeItemAndIncrement(self, ws, text, format=None):
        self.writeItem(ws, text, format)
        self.incrementRow(ws)
        self.incrementFileCount(ws)

    def writeOutcome(self, ws, text, format=None):
        self.writeHelper(ws, self.OUTCOME_COL, text, format)

    def writeOutcomeAndIncrement(self, ws, text, format=None):
        self.writeOutcome(ws, text, format)
        self.incrementRow(ws)
        self.incrementFileCount(ws)

    def writeHelper(self, ws, col: str, text: str, format=None):
        if (format): ws.write_string(self.sheetRows[ws], col, text, format)
        else: ws.write_string(self.sheetRows[ws], col, text)

    def incrementRow(self, ws, amount:int=1):
        self.sheetRows[ws] += amount

    def incrementFileCount(self, ws, amount:int=1):
        self.summaryCounts[ws] += amount


    def styleSummarySheet(self, dirAbsolute, includeSubFolders, modify):
        self.summarySheet.set_column(0, 0, 20)
        self.summarySheet.set_column(1, 1, 15)
        
        self.summarySheet.write(0, 0, "FilePath", self.headerFormat)
        self.summarySheet.write(1, 0, "Subfolders inclusion", self.headerFormat)
        self.summarySheet.write(2, 0, "Modify", self.headerFormat)
        self.summarySheet.write(3, 0, "Argument", self.headerFormat)
        self.summarySheet.write(4, 0, "File count", self.headerFormat)
        self.summarySheet.write(5, 0, "Error count / %", self.headerFormat)
        self.summarySheet.write(6, 0, "Execution time (s)", self.headerFormat)

        self.summarySheet.write(0, 1, dirAbsolute, self.summaryValueFormat)
        self.summarySheet.write(1, 1, str(includeSubFolders), self.summaryValueFormat)
        self.summarySheet.write(2, 1, str(modify), self.summaryValueFormat)


    def populateSummarySheet(self):
        if self.filesScannedCount == 0: errorPercentage = 0
        else: errorPercentage = round(self.errorCount / self.filesScannedCount * 100, 2)

        if (self.fixArg != None): self.summarySheet.write_string(3, 1, str(self.fixArg), self.summaryValueFormat)
        
        self.summarySheet.write_number(4, 1, self.filesScannedCount, self.summaryValueFormat)
        self.summarySheet.write_number(5, 1, self.errorCount, self.summaryValueFormat)
        self.summarySheet.write(5, 2, "{}%".format(errorPercentage), self.summaryValueFormat)
        self.summarySheet.write_number(6, 1, round(self.executionTime, 4), self.summaryValueFormat)

        i = 0
        for ws in self.getAllProcedureSheets():
            self.summarySheet.write(self.sheetRows[self.summarySheet] + i, 1, self.summaryCounts[ws], self.summaryValueFormat)
            i += 1

        
    def createHelpMeSheet(self):
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
        for ws in self.getAllProcedureSheets():
            ws.autofit()
            

    def close(self):
        self.populateSummarySheet()
        self.autofitSheets()
        self.createHelpMeSheet()
        self.wb.close()
