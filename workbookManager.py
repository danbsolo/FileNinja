import xlsxwriter
from typing import List
from time import time
import os
import stat
import filesScannedSharedVar
from ExcelWritePackage import ExcelWritePackage
from concurrent.futures import ThreadPoolExecutor, as_completed
from threading import Lock


class WorkbookManager:
    def __init__(self, workbookPathName: str):
        self.workbookPathName = workbookPathName
        self.wb = xlsxwriter.Workbook(workbookPathName)

        self.summarySheet = self.wb.add_worksheet("Summary")
        self.summarySheet.activate()  # view this worksheet on startup

        self.procedureObjectSheets = {} # procedureObject : worksheet
        self.procedureObjectFunctions = {}
        self.procedureObjectArgs = {}

        # Lists to avoid making objects with .keys() and .values() every time
        self.procedureObjects = []
        self.procedureSheets = []

        # Lists to avoid many redundant if statements
        self.fileProcedures = []
        self.folderProcedures = []

        # Differentiated so that Fix functions can access hidden files
        self.fileFindProcedures = []
        self.fileFixProcedures = []

        # For summarySheet, first # rows are used for mainstay metrics. Skip a line, then write a variable number of procedure metrics.
        self.sheetRows = {self.summarySheet: 14} # worksheet : Integer
        self.summaryCounts = {}  # worksheet : Integer

        # Summary metrics
        self.filesScannedCount = 0
        self.foldersScannedCount = 0
        self.fileErrorCount = 0
        self.folderErrorCount = 0
        self.executionTime = 0

        # Constant columns
        self.DIR_COL = 0
        self.ITEM_COL = 1
        self.OUTCOME_COL = 2  # describes either an Error or a Modification, depending on the procedure type
        self.AUXILIARY_COL = 3

        # Default cell styles
        self.dirFormat = self.wb.add_format({"bg_color": "#99CCFF", "bold": True})  # blueish
        self.errorFormat = self.wb.add_format({"bold": True})  # "bg_color": "#FF4444", # reddish
        self.modifyFormat = self.wb.add_format({"bg_color": "#00FF80", "bold": True})  # greenish
        self.logFormat = self.wb.add_format({"bg_color": "#9999FF", "bold": True})  # purplish
        self.headerFormat = self.wb.add_format({"bg_color": "#C0C0C0", "bold": True})  # grayish
        self.summaryValueFormat = self.wb.add_format({})

        self.warningWeakFormat = self.wb.add_format({"bg_color": "#FFEB9C", "bold": True})  # yellowish
        self.warningStrongFormat = self.wb.add_format({"bg_color": "#FFC7CE", "bold": True})  # reddish


    def getAllProcedureSheets(self):
        return self.procedureSheets


    def addProcedure(self, procedureObject, allowModify, addRecommendations, arg) -> bool:
        # Initialize procedure's corresponding worksheet
        tmpWsVar = self.wb.add_worksheet(procedureObject.name)
        self.summarySheet.write(self.sheetRows[self.summarySheet] +len(self.getAllProcedureSheets()), 0, procedureObject.name + " count", self.headerFormat)

        self.procedureObjects.append(procedureObject)
        self.procedureObjectSheets[procedureObject] = tmpWsVar
        self.procedureSheets.append(tmpWsVar)
        self.sheetRows[tmpWsVar] = 0
        self.summaryCounts[tmpWsVar] = 0

        self.procedureObjectFunctions[procedureObject] = procedureObject.getMainFunction(allowModify, addRecommendations)

        # Add to either fileProcedures or folderProcedures lists
        if procedureObject.getIsFileProcedure():
            self.fileProcedures.append(procedureObject)

            if procedureObject.isFixFunction():
                self.fileFixProcedures.append(procedureObject)
            else:
                self.fileFindProcedures.append(procedureObject)

        else:
            self.folderProcedures.append(procedureObject)

        # Set argument
        potentialArg = procedureObject.getValidArgument(arg)
        if potentialArg is None:
            return False

        self.procedureObjectArgs[procedureObject] = potentialArg
        return True
    
    
    def processFile(self, longDirAbsolute, dirAbsolute, fileName, needsFolderWritten, tpeIndex):
        # Onenote files have a ".one-----" extension. The longest onenote extension is 8 characters long. Ignore them.
        # > Technically, something called "fileName.one.txt" would get ignored, but the likelihood of that existing is very low
        # Hidden files should be ignored. This includes temporary Microsoft files (begins with "~$").
        if fileName.startswith("~$") or ".one" in fileName[-8:]:
            return
    
        longFileAbsolute = longDirAbsolute + "\\" + fileName

        hiddenStatus = self.hiddenFileCheck(longFileAbsolute)
        if hiddenStatus == 2:
            return
        
        countAsError = False

        ## THREADING STUFF STARTS
        futures = {
            self.fileProcedureThreadPoolExecutors[tpeIndex].submit(
                self.procedureObjectFunctions[procedureObject],
                longFileAbsolute,
                longDirAbsolute,
                dirAbsolute,
                fileName,
                self.procedureObjectSheets[procedureObject],
            ): self.procedureObjectSheets[procedureObject]
            for procedureObject in self.fileFindProcedures
        }

        if hiddenStatus == 0:
            for procedureObject in self.fileFixProcedures:
                futures[self.fileProcedureThreadPoolExecutors[tpeIndex].submit(
                        self.procedureObjectFunctions[procedureObject],
                        longFileAbsolute,
                        longDirAbsolute,
                        dirAbsolute,
                        fileName,
                        self.procedureObjectSheets[procedureObject])] = self.procedureObjectSheets[procedureObject]


        for fut in as_completed(futures):
            result = fut.result()
            status = result[0]
            fileSheet = futures[fut]

            if (status == True):
                with self.lockFileNeedsFolderWritten:
                    needsFolderWritten.add(fileSheet)
                countAsError = True
            elif (not status):  # returning None or False
                continue  # It's not super necessary to continue here, but might as well
            elif (status == 2):  # Special case (ex: Used by List All Files)
                with self.lockFileNeedsFolderWritten:
                    needsFolderWritten.add(fileSheet)
            elif (status == 3):  # Special case (ex: used by Identical File)
                countAsError = True

            with self.workbookLock:
                for ewp in result[1:]:
                    ewp.executeWrite()

        # So two files don't finish and try to increment these counters simultaneously
        with self.lockFileScan:
            self.filesScannedCount += 1
            if countAsError:
                self.fileErrorCount += 1


    def fileCrawl(self, longDirAbsolute, dirAbsolute, dirFiles: List[str]):
        needsFolderWritten = set()

        futures = []
        for fileName in dirFiles:
            with self.fileThreadsCounterLock:
                nextFileThreadIndex = self.fileThreadsCounter % self.numFileThreads
                self.fileThreadsCounter += 1
                
            futures.append(self.fileThreadPoolExecutor.submit(
                self.processFile,
                longDirAbsolute,
                dirAbsolute,
                fileName,
                needsFolderWritten,
                nextFileThreadIndex
            ))

        # If any thread raises an exception, this will ensure they are raised in this "main" WorkbookManager thread
        # , as opposed to just using concurrent.futures(wait)
        for fut in as_completed(futures):
            fut.result()

        return needsFolderWritten            


    def folderCrawl(self, dirAbsolute, dirFolders, dirFiles):
        needsFolderWritten = set()
        countAsError = False

        futures = {}

        for procedureObject in self.folderProcedures:
            futures[self.folderProcedureThreadPoolExecutor.submit(
                self.procedureObjectFunctions[procedureObject],
                dirAbsolute,
                os.path.basename(dirAbsolute),
                dirFolders,
                dirFiles,
                self.procedureObjectSheets[procedureObject]
            )] = self.procedureObjectSheets[procedureObject]

        for fut in as_completed(futures):
            result = fut.result()
            status = result[0]
            folderSheet = futures[fut]

            if status == True:
                with self.lockFolderNeedsFolderWritten:
                    needsFolderWritten.add(folderSheet)
                countAsError = True
            elif not status:  # returning None or False
                pass
            elif status == 2:
                with self.lockFolderNeedsFolderWritten:
                    needsFolderWritten.add(self.procedureObjectSheets[procedureObject])
            elif status == 3:
                countAsError = True

            with self.workbookLock:
                for ewp in result[1:]:
                    ewp.executeWrite()

        if countAsError:
            self.folderErrorCount += 1

        return needsFolderWritten


    def isHidden(self, longItemAbsolute):
        return bool(os.stat(longItemAbsolute).st_file_attributes & stat.FILE_ATTRIBUTE_HIDDEN)

    # When we want to INCLUDE all hidden files (for find procedures only)
    def includeHiddenFilesCheck(self, longFileAbsolute):
        # 0 == not hidden
        # 1 == hidden, but not required to skip for find procedures, just for fix procedures
        if self.isHidden(longFileAbsolute):
            return 1
        else:
            return 0
    
    # When we want to EXCLUDE all hidden files
    def excludeHiddenFilesCheck(self, longFileAbsolute):
        # 0 == not hidden
        # 2 == hidden, and we must skip in totality
        if self.isHidden(longFileAbsolute):
            return 2
        else:
            return 0
        

    def doFileProceduresExist(self):
        return len(self.fileProcedures) != 0

    def doFolderProceduresExist(self):
        return len(self.folderProcedures) != 0
    

    def initFileCrawlOnly(self, longDirAbsolute, dirAbsolute, dirFiles, _):
        return self.fileCrawl(longDirAbsolute, dirAbsolute, dirFiles)

    def initFolderCrawlOnly(self, _, dirAbsolute, dirFiles, dirFolders):
        return self.folderCrawl(dirAbsolute, dirFolders, dirFiles)

    def initFileFolderCrawl(self, longDirAbsolute, dirAbsolute, dirFiles, dirFolders):
        fileCrawlFuture = self.crawlThreadPoolExecutor.submit(
            self.fileCrawl,
            longDirAbsolute,
            dirAbsolute,
            dirFiles
        )

        folderCrawlFuture = self.crawlThreadPoolExecutor.submit(
            self.folderCrawl,
            dirAbsolute,
            dirFolders,
            dirFiles
        )
        
        # union operator usage, lol
        return folderCrawlFuture.result() | fileCrawlFuture.result()
    

    def createFileThreads(self):
        ## Create thread pool executors and necessary locks
        numFileProcedures = len(self.fileProcedures)

        self.fileProcedureThreadPoolExecutors = []

        # Dynamically choose the number of file threads based on a hard-coded max total.
        totalThreads = 150
        self.numFileThreads = totalThreads // numFileProcedures

        self.fileThreadsCounter = 0  # used by self.fileCrawl()
        self.fileThreadsCounterLock = Lock()

        for _ in range(self.numFileThreads):
            self.fileProcedureThreadPoolExecutors.append(
                ThreadPoolExecutor(max_workers = numFileProcedures)
            )

        self.fileThreadPoolExecutor = ThreadPoolExecutor(max_workers = self.numFileThreads)

        self.lockFileNeedsFolderWritten = Lock()
        self.lockFileScan = Lock()

    def createFolderThreads(self):
        numFolderProcedures = len(self.folderProcedures)
        self.folderProcedureThreadPoolExecutor = ThreadPoolExecutor(max_workers = numFolderProcedures)

        self.lockFolderNeedsFolderWritten = Lock()

    def createCrawlThreads(self):
        # One for file crawl and one for folder crawl
        self.crawlThreadPoolExecutor = ThreadPoolExecutor(max_workers = 2)

    def createSheetLocks(self):
        self.sheetLocks = {}
        for ws in self.getAllProcedureSheets():
            self.sheetLocks[ws] = Lock()


    def initiateCrawl(self, baseDirAbsolute, includeSubdirectories, allowModify, includeHiddenFiles, addRecommendations, excludedDirs):
        def addLongPathPrefix(dirAbsolute):
            if dirAbsolute.startswith('\\\\'):
                return '\\\\?\\UNC' + dirAbsolute[1:]
            else:
                return '\\\\?\\' + dirAbsolute
        
        start = time()

        if includeHiddenFiles:
            self.hiddenFileCheck = lambda longFileAbsolute: self.includeHiddenFilesCheck(longFileAbsolute)
        else:
            self.hiddenFileCheck = lambda longFileAbsolute: self.excludeHiddenFilesCheck(longFileAbsolute)
        
        self.styleSummarySheet(baseDirAbsolute, includeSubdirectories, allowModify, includeHiddenFiles, addRecommendations)
        

        if self.doFileProceduresExist():
            self.createFileThreads()

            if self.doFolderProceduresExist():
                self.createFolderThreads()
                self.createCrawlThreads()
                crawlFunction = self.initFileFolderCrawl
            else:
                crawlFunction = self.initFileCrawlOnly

        elif self.doFolderProceduresExist():
            self.createFolderThreads()
            crawlFunction = self.initFolderCrawlOnly
        
        self.createSheetLocks()

        # Only one thread can access the workbook at a time, hence a lock
        self.workbookLock = Lock()

        sheetsSansNonConcurrent = []
        for procedureObject in self.procedureObjects:
            if procedureObject.getIsConcurrentOnly():
                sheetsSansNonConcurrent.append(self.procedureObjectSheets[procedureObject])


        walkObject = []

        if (includeSubdirectories):
            walkObject = os.walk(baseDirAbsolute)
        
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


        for procedureObject in self.procedureObjects:
            potentialStartFunction = procedureObject.getStartFunction()
            if potentialStartFunction:
                potentialStartFunction(self.procedureObjectArgs[procedureObject], self.procedureObjectSheets[procedureObject])
        
        # TODO: Clean this up?
        excludedDirsSet = set(excludedDirs)

        initialRows = {}
        for (dirAbsolute, dirFolders, dirFiles) in walkObject:
            # Ignore specifically OneNote_RecycleBin folders. Assumes these NEVER have subdirectories.
            # Ignore excluded directories.
            # Get "long file path". If folder is hidden, ignore it. Anything within a hidden folder is inadvertently ignored.
            if (dirAbsolute.endswith("OneNote_RecycleBin")) or (dirAbsolute in excludedDirsSet) or self.isHidden(longDirAbsolute := addLongPathPrefix(dirAbsolute)):
                dirFolders[:] = []
                continue

            initialRows.clear()
            for ws in sheetsSansNonConcurrent:
                initialRows[ws] = self.sheetRows[ws] + 1

            needsFolderWritten = crawlFunction(longDirAbsolute, dirAbsolute, dirFiles, dirFolders)

            for ws in needsFolderWritten:
                ws.write(initialRows[ws], self.DIR_COL, dirAbsolute, self.dirFormat)
 
            self.foldersScannedCount += 1
            filesScannedSharedVar.FILES_SCANNED += len(dirFiles)  # just a rough estimate of the number of files scanned


        for procedureObject in self.procedureObjects:
            potentialPostFunction = procedureObject.getPostFunction(addRecommendations)
            if potentialPostFunction:
                potentialPostFunction(self.procedureObjectSheets[procedureObject])


        # shutdown threads
        if self.doFileProceduresExist():
            self.fileThreadPoolExecutor.shutdown(wait=True)
            for i in range(self.numFileThreads):
                self.fileProcedureThreadPoolExecutors[i].shutdown(wait=True)

            if self.doFolderProceduresExist():
                self.folderProcedureThreadPoolExecutor.shutdown(wait=True)
                self.crawlThreadPoolExecutor.shutdown(wait=True)

        elif self.doFolderProceduresExist():
            self.folderProcedureThreadPoolExecutor.shutdown(wait=True)
        #
        
        self.executionTime = time() - start
        self.populateSummarySheet(excludedDirs)
        self.autofitSheets()
        


    def incrementRow(self, ws):
        with self.sheetLocks[ws]:
            self.sheetRows[ws] += 1

    def incrementFileCount(self, ws, amount=1):
        with self.sheetLocks[ws]:
            self.summaryCounts[ws] += amount

    def incrementRowAndFileCount(self, ws, amount=1):
       with self.sheetLocks[ws]:
            self.sheetRows[ws] += amount
            self.summaryCounts[ws] += amount


    def styleSummarySheet(self, dirAbsolute, includeSubdirectories, allowModify, includeHiddenFiles, addRecommendations):
        self.summarySheet.set_column(0, 0, 34)
        self.summarySheet.set_column(1, 1, 15)
        
        self.summarySheet.write(0, 0, "Directory Path", self.headerFormat)
        self.summarySheet.write(1, 0, "Excluded Directories", self.headerFormat)
        self.summarySheet.write(2, 0, "Include Subdirectories", self.headerFormat)
        self.summarySheet.write(3, 0, "Allow Modify", self.headerFormat)
        self.summarySheet.write(4, 0, "Include Hidden Files", self.headerFormat)
        self.summarySheet.write(5, 0, "Add Recommendations", self.headerFormat)
        self.summarySheet.write(6, 0, "Argument(s)", self.headerFormat)
        self.summarySheet.write(8, 0, "Directory count", self.headerFormat)
        self.summarySheet.write(9, 0, "Directory error count / %", self.headerFormat)
        self.summarySheet.write(10, 0, "File count", self.headerFormat)
        self.summarySheet.write(11, 0, "File error count / %", self.headerFormat)
        self.summarySheet.write(12, 0, "Execution time (s)", self.headerFormat)

        self.summarySheet.write(0, 1, dirAbsolute, self.summaryValueFormat)
        self.summarySheet.write(2, 1, str(includeSubdirectories), self.summaryValueFormat)
        self.summarySheet.write(3, 1, str(allowModify), self.summaryValueFormat)
        self.summarySheet.write(4, 1, str(includeHiddenFiles), self.summaryValueFormat)
        self.summarySheet.write(5, 1, str(addRecommendations), self.summaryValueFormat)


    def populateSummarySheet(self, excludedDirs):
        col = 1

        for procedureObject in self.procedureObjects:
            arg = self.procedureObjectArgs[procedureObject]
            
            if arg == ():
                continue
            
            self.summarySheet.write(6, col, f"{arg[0] if len(arg) <= 1 else arg} : {procedureObject.name}", self.summaryValueFormat)
            col += 1

        if self.filesScannedCount == 0: fileErrorPercentage = 0
        else: fileErrorPercentage = round(self.fileErrorCount / self.filesScannedCount * 100, 2)
        if self.foldersScannedCount == 0: folderErrorPercentage = 0
        else: folderErrorPercentage = round(self.folderErrorCount / self.foldersScannedCount * 100, 2)

        self.summarySheet.write_number(8, 1, self.foldersScannedCount, self.summaryValueFormat)
        self.summarySheet.write_number(9, 1, self.folderErrorCount, self.summaryValueFormat)
        self.summarySheet.write(9, 2, "{}%".format(folderErrorPercentage), self.summaryValueFormat)
        self.summarySheet.write_number(10, 1, self.filesScannedCount, self.summaryValueFormat)
        self.summarySheet.write_number(11, 1, self.fileErrorCount, self.summaryValueFormat)
        self.summarySheet.write(11, 2, "{}%".format(fileErrorPercentage), self.summaryValueFormat)
        self.summarySheet.write_number(12, 1, round(self.executionTime, 4), self.summaryValueFormat)
        
        i = 1
        for exDir in excludedDirs:
            self.summarySheet.write_string(1, i, exDir, self.summaryValueFormat)
            i += 1

        i = 0
        for ws in self.getAllProcedureSheets():
            self.summarySheet.write(self.sheetRows[self.summarySheet] + i, 1, self.summaryCounts[ws], self.summaryValueFormat)
            i += 1


    def autofitSheets(self):
        for ws in self.procedureSheets:
            ws.autofit()


    def close(self):
        self.wb.close()
