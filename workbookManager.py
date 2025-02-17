import xlsxwriter
from typing import List
from time import time
import os
from copy import deepcopy


class WorkbookManager:

    def __init__(self, workbookPathName: str):
        self.workbookPathName = workbookPathName
        self.wb = xlsxwriter.Workbook(workbookPathName)

        self.summarySheet = self.wb.add_worksheet("Summary")
        self.summarySheet.activate()  # view this worksheet on startup

        self.excludedDirs = []  # set within initiateCrawl()

        self.findSheets = {} # procedureObject : worksheet
        self.fixSheets = {}
        self.fixProcedureArgs = {}
        self.fixProcedureFunctions = {}
        # For summarySheet, first # rows are used for mainstay metrics. Skip a line, then write variable number of procedure metrics.
        self.sheetRows = {self.summarySheet: 11} # worksheet : Integer
        self.summaryCounts = {}  # worksheet : Integer

        # Lists to avoid many redundant if statements
        self.findProceduresConcurrentOnly = []
        self.fileFindProcedures = []
        self.fileFixProcedures = []
        self.folderFindProcedures = []
        self.folderFixProcedures = []

        # Summary metrics
        self.filesScannedCount = 0
        self.foldersScannedCount = 0
        self.executionTime = 0
        self.errorCount = 0

        # Constant columns
        self.DIR_COL = 0
        self.ITEM_COL = 1
        self.OUTCOME_COL = 2  # describes either an Error or a Modification, depending on the procedure type

        # Default cell styles
        self.dirFormat = self.wb.add_format({"bg_color": "#99CCFF", "bold": True})  # blueish
        self.errorFormat = self.wb.add_format({"bold": True})  # "bg_color": "#FF4444", # reddish
        self.modifyFormat = self.wb.add_format({"bg_color": "#00FF80", "bold": True})  # greenish
        self.logFormat = self.wb.add_format({"bg_color": "#9999FF", "bold": True})  # purplish
        self.headerFormat = self.wb.add_format({"bg_color": "#C0C0C0", "bold": True})  # grayish
        self.summaryValueFormat = self.wb.add_format({})


    def getAllProcedureSheets(self):
        return (list(self.findSheets.values()) + list(self.fixSheets.values()))


    def addFindProcedure(self, findProcedureObject):
        tmpWsVar = self.wb.add_worksheet(findProcedureObject.name)
        self.summarySheet.write(self.sheetRows[self.summarySheet] +len(self.getAllProcedureSheets()), 0, findProcedureObject.name + " count", self.headerFormat)
        
        self.findSheets[findProcedureObject] = tmpWsVar
        self.sheetRows[tmpWsVar] = 1
        self.summaryCounts[tmpWsVar] = 0

        if findProcedureObject.isConcurrentOnly:
            tmpWsVar.freeze_panes(1, 0)
            tmpWsVar.write(0, self.DIR_COL, "Directories", self.headerFormat)
            tmpWsVar.write(0, self.ITEM_COL, "Items", self.headerFormat)
            tmpWsVar.write(0, self.OUTCOME_COL, "Error", self.headerFormat)

            self.findProceduresConcurrentOnly.append(findProcedureObject)
        
        if findProcedureObject.isFileFind:
            self.fileFindProcedures.append(findProcedureObject)
        else:
            self.folderFindProcedures.append(findProcedureObject)
            

    def addFixProcedure(self, fixProcedureObject, allowModify, arg) -> bool:
        tmpWsVar = self.wb.add_worksheet(fixProcedureObject.name)
        self.summarySheet.write(self.sheetRows[self.summarySheet] +len(self.getAllProcedureSheets()), 0, fixProcedureObject.name + " count", self.headerFormat)

        self.fixSheets[fixProcedureObject] = tmpWsVar
        self.sheetRows[tmpWsVar] = 1
        self.summaryCounts[tmpWsVar] = 0

        if allowModify:
            self.fixProcedureFunctions[fixProcedureObject] = fixProcedureObject.modifyFunction
        else:
            self.fixProcedureFunctions[fixProcedureObject] = fixProcedureObject.logFunction
        
        if fixProcedureObject.isFileFix:
            self.fileFixProcedures.append(fixProcedureObject)
        else:
            self.folderFixProcedures.append(fixProcedureObject)

        tmpWsVar.freeze_panes(1, 0)
        tmpWsVar.write(0, self.DIR_COL, "Directories", self.headerFormat)
        tmpWsVar.write(0, self.ITEM_COL, "Items", self.headerFormat)
        tmpWsVar.write(0, self.OUTCOME_COL, fixProcedureObject.columnName, self.headerFormat)

        return self.setFixArg(fixProcedureObject, arg)


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

            for findProcedureObject in self.fileFindProcedures:
                if (findProcedureObject.mainFunction(dirAbsolute, itemName, self.findSheets[findProcedureObject])):
                    if (not alreadyCounted):
                        self.errorCount += 1
                        alreadyCounted = True

            for fixProcedureObject in self.fileFixProcedures:
                self.fixProcedureFunctions[fixProcedureObject](dirAbsolute, itemName, self.fixSheets[fixProcedureObject], self.fixProcedureArgs[fixProcedureObject])

            alreadyCounted = False
            self.filesScannedCount += 1                    


    def folderCrawl(self, dirAbsolute, dirFolders, dirFiles):
        for findProcedureObject in self.folderFindProcedures:
            findProcedureObject.mainFunction(dirAbsolute, dirFolders, dirFiles, self.findSheets[findProcedureObject])

        for fixProcedureObject in self.folderFixProcedures:
            self.fixProcedureFunctions[fixProcedureObject](dirAbsolute, dirFolders, dirFiles, self.fixSheets[fixProcedureObject], self.fixProcedureArgs[fixProcedureObject])

        self.foldersScannedCount += 1


    def initiateCrawl(self, baseDirAbsolute, includeSubfolders, allowModify, excludedDirs):
        start = time()

        self.styleSummarySheet(baseDirAbsolute, includeSubfolders, allowModify)
        self.excludedDirs = excludedDirs
        copyOfExcludedDirs = deepcopy(excludedDirs)

        #
        sheetsSansNonConcurrentAndFolderFind = []
        for findProcedureObject in list(self.findSheets.keys()):
            if findProcedureObject.isConcurrentOnly and findProcedureObject.isFileFind:
                sheetsSansNonConcurrentAndFolderFind.append(self.findSheets[findProcedureObject])
        sheetsSansNonConcurrentAndFolderFind.extend(list(self.fixSheets.values()))  # Adds all fix procedures indiscrimnately

        ##
        walkObject = []

        if (includeSubfolders):
            if (not excludedDirs):
                walkObject = os.walk(baseDirAbsolute)
            else:
                for dirAbsolute, dirFolders, dirFiles in os.walk(baseDirAbsolute):
                    isDirIncluded = True
                    
                    ### NOTE: Should this ignore hidden folders? i.e. Folders that begin with ".". Probably.

                    for exDir in copyOfExcludedDirs:
                        if dirAbsolute.startswith(exDir):  # exclude by subDirAbsolute to be precise
                            isDirIncluded = False
                            copyOfExcludedDirs.remove(exDir)
                            break

                    if isDirIncluded:
                        walkObject.append((dirAbsolute, dirFolders, dirFiles))
                    else:
                        dirFolders[:] = []  # stop traversal of this folder and its subfolders
        
        else:
            # mimic os.walk()'s output but only for the current directory
            dirFolders = []
            dirFiles = []
            
            for item in os.listdir(baseDirAbsolute):
                if os.path.isfile(os.path.join(baseDirAbsolute, item)):
                    dirFiles.append(item)
                else:
                    dirFolders.append(item)

            walkObject = [(baseDirAbsolute, dirFolders, dirFiles)]
        ##

        #
        for (dirAbsolute, dirFolders, dirFiles) in walkObject:
            for ws in sheetsSansNonConcurrentAndFolderFind:
                ws.write(self.sheetRows[ws], self.DIR_COL, dirAbsolute, self.dirFormat)

            self.folderCrawl(dirAbsolute, dirFolders, dirFiles)
            self.fileCrawl(dirAbsolute, dirFiles)

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


    def styleSummarySheet(self, dirAbsolute, includeSubFolders, allowModify):
        self.summarySheet.set_column(0, 0, 34)
        self.summarySheet.set_column(1, 1, 15)
        
        self.summarySheet.write(0, 0, "File Path", self.headerFormat)
        self.summarySheet.write(1, 0, "Excluded Directories", self.headerFormat)
        self.summarySheet.write(2, 0, "Include Subfolders", self.headerFormat)
        self.summarySheet.write(3, 0, "Allow Modify", self.headerFormat)
        self.summarySheet.write(4, 0, "Argument(s)", self.headerFormat)
        self.summarySheet.write(6, 0, "Folder count", self.headerFormat)
        self.summarySheet.write(7, 0, "File count", self.headerFormat)
        self.summarySheet.write(8, 0, "File error count / %", self.headerFormat)
        self.summarySheet.write(9, 0, "Execution time (s)", self.headerFormat)

        self.summarySheet.write(0, 1, dirAbsolute, self.summaryValueFormat)
        self.summarySheet.write(2, 1, str(includeSubFolders), self.summaryValueFormat)
        self.summarySheet.write(3, 1, str(allowModify), self.summaryValueFormat)


    def populateSummarySheet(self):
        if self.filesScannedCount == 0: errorPercentage = 0
        else: errorPercentage = round(self.errorCount / self.filesScannedCount * 100, 2)

        col = 1
        for fixProcedureObject in self.fixProcedureArgs.keys():
            arg = self.fixProcedureArgs[fixProcedureObject]
            if arg == None: continue
            self.summarySheet.write(4, col, f"{arg[0] if len(arg) <= 1 else arg} : {fixProcedureObject.name}", self.summaryValueFormat)
            col += 1
        
        self.summarySheet.write_number(6, 1, self.foldersScannedCount, self.summaryValueFormat)
        self.summarySheet.write_number(7, 1, self.filesScannedCount, self.summaryValueFormat)
        self.summarySheet.write_number(8, 1, self.errorCount, self.summaryValueFormat)
        self.summarySheet.write(8, 2, "{}%".format(errorPercentage), self.summaryValueFormat)
        self.summarySheet.write_number(9, 1, round(self.executionTime, 4), self.summaryValueFormat)
        
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

        for fixProcedureSheet in self.fixSheets.values():
            fixProcedureSheet.autofit()


    def close(self):
        self.populateSummarySheet()
        self.autofitSheets()
        self.createHelpMeSheet()
        self.wb.close()
