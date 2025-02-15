import xlsxwriter
from typing import List, Tuple
from time import time

class WorkbookManager:

    def __init__(self, workbookPathName: str):
        self.workbookPathName = workbookPathName
        self.wb = xlsxwriter.Workbook(workbookPathName)

        self.summarySheet = self.wb.add_worksheet("Summary")
        self.summarySheet.activate()  # view this worksheet on startup

        self.excludedDirs = []  # set within folderCrawl()

        self.fixSheet = None # initializing for clarity
        self.fileFixProcedure = lambda _1, _2, _3: None  # default dummy function
        self.folderFixProcedure = lambda _1, _2, _3, _4: None
        self.fixArg = None  # initialize fixArg to a dummy value

        self.findSheets = {} # procedureObject : worksheet
        # For summarySheet first 8 rows are used for mainstay metrics. Skip a line, then write variable number of errors.
        self.sheetRows = {self.summarySheet: 10} # worksheet : Integer
        self.summaryCounts = {}  # worksheet : Integer

        # Summary metrics
        self.filesScannedCount = 0
        self.foldersScannedCount = 0
        self.executionTime = 0
        self.errorCount = 0

        self.DIR_COL = 0
        self.ITEM_COL = 1
        # used describe either an Error or a Modification, depending on the procedure type
        self.OUTCOME_COL = 2

        # Default cell styles
        self.dirFormat = self.wb.add_format({
            "bg_color": "#99CCFF", # blueish
            "bold": True
        })

        # "bg_color": "#FF4444", # reddish
        self.errorFormat = self.wb.add_format({
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


        self.findProceduresConcurrentOnly = []

        self.fixSheetsDict = {}
        self.fixProcedureArgs = {}
        self.fixProcedureFunctions = {}


    def getAllProcedureSheets(self):
        sheets = list(self.findSheets.values()) + list(self.fixSheetsDict.values())
        return sheets

    def getAllProcedureSheetsSansNonConcurrentAndFolderFind(self):
        sheets = []

        for findProcedureObject in list(self.findSheets.keys()):
            if findProcedureObject.isConcurrentOnly and findProcedureObject.isFileFind:
                sheets.append(self.findSheets[findProcedureObject])

        # ADDS ALL FIX PROCEDURES INDISCRIMINATELY. Could possibly just use extend()?
        for fixProcedureObject in self.fixSheetsDict.keys():
            sheets.append(self.fixSheetsDict[fixProcedureObject])

        
        #if self.fixSheet:
        #    sheets.append(self.fixSheet)

        return sheets

    def addFindProcedure(self, findProcedureObject):
        tmpWsVar = self.wb.add_worksheet(findProcedureObject.name)
        self.summarySheet.write(self.sheetRows[self.summarySheet] +len(self.findSheets.keys()) +len(self.fixSheetsDict.keys()), 0, findProcedureObject.name + " count", self.headerFormat)
        
        self.findSheets[findProcedureObject] = tmpWsVar
        self.sheetRows[tmpWsVar] = 1
        self.summaryCounts[tmpWsVar] = 0

        if findProcedureObject.isConcurrentOnly:
            tmpWsVar.freeze_panes(1, 0)
            tmpWsVar.write(0, self.DIR_COL, "Directories", self.headerFormat)
            tmpWsVar.write(0, self.ITEM_COL, "Items", self.headerFormat)
            tmpWsVar.write(0, self.OUTCOME_COL, "Error", self.headerFormat)

            self.findProceduresConcurrentOnly.append(findProcedureObject)


    def addFixProcedure(self, fixProcedureObject, modify, arg) -> bool:
        tmpWsVar = self.wb.add_worksheet(fixProcedureObject.name)
        self.summarySheet.write(self.sheetRows[self.summarySheet] +len(self.findSheets.keys()) +len(self.fixSheetsDict.keys()), 0, fixProcedureObject.name + " count", self.headerFormat)

        self.fixSheetsDict[fixProcedureObject] = tmpWsVar
        self.sheetRows[tmpWsVar] = 1
        self.summaryCounts[tmpWsVar] = 0

        if modify:
            self.fixProcedureFunctions[fixProcedureObject] = fixProcedureObject.modifyFunction
        else:
            self.fixProcedureFunctions[fixProcedureObject] = fixProcedureObject.logFunction

        tmpWsVar.freeze_panes(1, 0)
        tmpWsVar.write(0, self.DIR_COL, "Directories", self.headerFormat)
        tmpWsVar.write(0, self.ITEM_COL, "Items", self.headerFormat)
        tmpWsVar.write(0, self.OUTCOME_COL, fixProcedureObject.columnName, self.headerFormat)

        return self.setFixArg(fixProcedureObject, arg)
        
                        

#    def setFixProcedure(self, fixProcedureObject, modify):
#        self.summarySheet.write(self.sheetRows[self.summarySheet] +len(self.findSheets.keys()), 0, fixProcedureObject.name + " count", self.headerFormat)
#        tmpWsVar = self.wb.add_worksheet(fixProcedureObject.name)
#        self.fixSheet = tmpWsVar
#        self.sheetRows[tmpWsVar] = 1
#        self.summaryCounts[tmpWsVar] = 0
#
#        if fixProcedureObject.isFileFix:
#            if modify: self.fileFixProcedure = fixProcedureObject.modifyFunction
#            else: self.fileFixProcedure = fixProcedureObject.logFunction
#        else:
#            if modify: self.folderFixProcedure = fixProcedureObject.modifyFunction
#            else: self.folderFixProcedure = fixProcedureObject.logFunction

#        self.fixSheet.freeze_panes(1, 0)
#        self.fixSheet.write(0, self.DIR_COL, "Directories", self.headerFormat)
#        self.fixSheet.write(0, self.ITEM_COL, "Items", self.headerFormat)
#        self.fixSheet.write(0, self.OUTCOME_COL, fixProcedureObject.columnName, self.headerFormat)


    def setFixArg(self, fixProcedureObject, arg) -> bool:
        if not fixProcedureObject.validatorFunction:
            pass
        elif ((arg := fixProcedureObject.validatorFunction(arg, fixProcedureObject.argBoundary)) is None):
            return False

        self.fixProcedureArgs[fixProcedureObject] = arg
        return True
        

    def fileCrawl(self, dirAbsolute, dirItems: List[str]):
        alreadyCounted = False

        for itemName in dirItems:            
            # Temporary Microsoft files begin with "~$". If it is so, skip file
            # Onenote files have a ".one[]" extension. The longest onenote extension is 8 characters long.
            # Technically, something called "fileName.one.txt" would get ignored, but the likelihood of that existing is very low 
            if itemName.startswith("~$") or ".one" in itemName[-8:]:
                continue

            for findProcedureObject in self.findSheets.keys():
                # BRUTE FORCE. Proof of concept.
                if not findProcedureObject.isFileFind:
                    continue
                #
                if (findProcedureObject.mainFunction(dirAbsolute, itemName, self.findSheets[findProcedureObject])):
                    if (not alreadyCounted):
                        self.errorCount += 1
                        alreadyCounted = True

            #
            for fixProcedureObject in self.fixSheetsDict.keys():
                if fixProcedureObject.isFileFix:
                    self.fixProcedureFunctions[fixProcedureObject](dirAbsolute, itemName, self.fixSheetsDict[fixProcedureObject], self.fixProcedureArgs[fixProcedureObject])
            #


            # self.fileFixProcedure(dirAbsolute, itemName, self.fixSheet)

            alreadyCounted = False
            self.filesScannedCount += 1                    


    def folderCrawl(self, dirTree: List[Tuple[str, list, list]], excludedDirs):
        start = time()

        self.excludedDirs = excludedDirs

        allSheets = self.getAllProcedureSheetsSansNonConcurrentAndFolderFind()
        
        #
        folderFindProcedures = []
        for findProcedureObject in self.findSheets.keys():
            if not (findProcedureObject.isFileFind):
                folderFindProcedures.append(findProcedureObject)

        folderFixProcedures = []
        for fixProcedureObject in self.fixSheetsDict.keys():
            if not (fixProcedureObject.isFileFix):
                folderFixProcedures.append(fixProcedureObject)
        #

        for (dirAbsolute, dirFolders, dirFiles) in dirTree:
            for ws in allSheets:
                ws.write(self.sheetRows[ws], self.DIR_COL, dirAbsolute, self.dirFormat)

            self.fileCrawl(dirAbsolute, dirFiles)

            # Folder fix procedure
            # self.folderFixProcedure(dirAbsolute, dirFolders, dirFiles, self.fixSheet)

            self.foldersScannedCount += 1

            #
            for findProcedureObject in folderFindProcedures:
                findProcedureObject.mainFunction(dirAbsolute, dirFolders, dirFiles, self.findSheets[findProcedureObject])

            for fixProcedureObject in folderFixProcedures:
                self.fixProcedureFunctions[fixProcedureObject](dirAbsolute, dirFolders, dirFiles, self.fixSheetsDict[fixProcedureObject], self.fixProcedureArgs[fixProcedureObject])
            #

        for findProcedureObject in self.findSheets.keys():
            if findProcedureObject.postFunction:
                findProcedureObject.postFunction(self.findSheets[findProcedureObject])

        end = time()
        self.executionTime = end - start

    def writeDir(self, ws, text, format=None):
        self.writeHelper(ws, self.DIR_COL, text, format)
    
    def writeDirAndIncrement(self, ws, text, format=None):
        self.writeDir(ws, text, format)
        self.incrementRow(ws)
        self.incrementFileCount(ws)

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
        self.summarySheet.set_column(0, 0, 34)
        self.summarySheet.set_column(1, 1, 15)
        
        self.summarySheet.write(0, 0, "File Path", self.headerFormat)
        self.summarySheet.write(1, 0, "Excluded Directories", self.headerFormat)
        self.summarySheet.write(2, 0, "Subfolders inclusion", self.headerFormat)
        self.summarySheet.write(3, 0, "Modify", self.headerFormat)
        self.summarySheet.write(4, 0, "Argument", self.headerFormat)
        self.summarySheet.write(5, 0, "Folder count", self.headerFormat)
        self.summarySheet.write(6, 0, "File count", self.headerFormat)
        self.summarySheet.write(7, 0, "Error count / %", self.headerFormat)
        self.summarySheet.write(8, 0, "Execution time (s)", self.headerFormat)

        self.summarySheet.write(0, 1, dirAbsolute, self.summaryValueFormat)
        self.summarySheet.write(2, 1, str(includeSubFolders), self.summaryValueFormat)
        self.summarySheet.write(3, 1, str(modify), self.summaryValueFormat)


    def populateSummarySheet(self):
        if self.filesScannedCount == 0: errorPercentage = 0
        else: errorPercentage = round(self.errorCount / self.filesScannedCount * 100, 2)

        # if (self.fixArg != None): self.summarySheet.write_string(4, 1, str(self.fixArg), self.summaryValueFormat)
        col = 1
        for fixProcedureObject in self.fixProcedureArgs.keys():
            self.summarySheet.write(4, col, "{} -> {}".format(self.fixProcedureArgs[fixProcedureObject], fixProcedureObject.name), self.summaryValueFormat)
            col += 1
        
        self.summarySheet.write_number(5, 1, self.foldersScannedCount, self.summaryValueFormat)
        self.summarySheet.write_number(6, 1, self.filesScannedCount, self.summaryValueFormat)
        self.summarySheet.write_number(7, 1, self.errorCount, self.summaryValueFormat)
        self.summarySheet.write(7, 2, "{}%".format(errorPercentage), self.summaryValueFormat)
        self.summarySheet.write_number(8, 1, round(self.executionTime, 4), self.summaryValueFormat)
        
        i = 1
        for exDir in self.excludedDirs:
            self.summarySheet.write_string(1, i, exDir, self.summaryValueFormat)
            i += 1

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
        for findProcedureObject in self.findProceduresConcurrentOnly:
            self.findSheets[findProcedureObject].autofit()

        for fixProcedureSheet in self.fixSheetsDict.values():
            fixProcedureSheet.autofit()
        
        #if self.fixSheet:
        #    self.fixSheet.autofit()

            

    def close(self):
        self.populateSummarySheet()
        self.autofitSheets()
        self.createHelpMeSheet()
        self.wb.close()
