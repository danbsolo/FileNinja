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

        self.sheetRows = {} # worksheet : Integer
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
        self.boldFormat = self.wb.add_format({"bold": True})
        self.modifyFormat = self.wb.add_format({"bg_color": "#00FF80", "bold": True})  # greenish
        self.logFormat = self.wb.add_format({"bg_color": "#8A7768", "bold": True})  # brownish
        self.headerFormat = self.wb.add_format({"bg_color": "#C0C0C0", "bold": True})  # grayish
        self.summaryValueFormat = self.wb.add_format({})
        self.separatorFormat = self.wb.add_format({"bg_color": "#202124", "font_color": "#FFFFFF"})

        self.warningMiddlingFormat = self.wb.add_format({"bg_color": "#CCC0DA", "bold": True})
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
        if self.functionMap["processFileExcludeExtensions"][self.processFileExcludeExtensionsOption](fileName):
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


    def initiateCrawl(self, baseDirAbsolute, includeSubdirectories, allowModify, includeHiddenFiles, addRecommendations, excludedDirs, excludedExtensions):
        def addLongPathPrefix(dirAbsolute):
            if dirAbsolute.startswith('\\\\'):
                return '\\\\?\\UNC' + dirAbsolute[1:]
            else:
                return '\\\\?\\' + dirAbsolute
        
        start = time()

        # TODO: Perhaps introduce a parallel "functionOption"
        self.functionMap = {}

        self.functionMap["processFileExcludeExtensions"] = {
            True: lambda fileName: fileName.startswith("~$") or os.path.splitext(fileName)[1].lower() in self.excludedExtensions,
            False: lambda fileName: fileName.startswith("~$")
        }
        if excludedExtensions:
            self.processFileExcludeExtensionsOption = True
            self.excludedExtensions = excludedExtensions
        else:
            self.processFileExcludeExtensionsOption = False
            self.excludedExtensions = None

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
        
        # autofit all concurrent only sheets
        for procedureObject in self.procedureObjects:
            if procedureObject.getIsConcurrentOnly():
                self.procedureObjectSheets[procedureObject].autofit()
        
        self.executionTime = time() - start
        self.populateSummarySheet(excludedDirs)


    def createMetaSheets(self, addRecommendations):
        self.summarySheet = self.wb.add_worksheet("Summary")
        self.summarySheet.activate()  # view this worksheet on startup

        # For summarySheet, first # rows are used for mainstay metrics. Skip a line, then write a variable number of procedure metrics.
        self.sheetRows[self.summarySheet] = 15

        # create the legend sheet if recommendations are on
        if not addRecommendations: return

        ws = self.wb.add_worksheet("Legend")

        # format maps
        baseFontMap = {"font_name": "Aptos Narrow"}
        baseFontBoldMap = {**baseFontMap, "bold": True}
        wrappedMap = {**baseFontMap, "text_wrap": True}

        blueMap = {**baseFontBoldMap, "bg_color": "#156082", "font_color": "white"}
        redMap = {**baseFontMap, "bg_color": "#FFC7CE", "valign": "vcenter"}
        redCenterMap = {**redMap, "bold": True, "align": "center"}
        yellowMap = {**baseFontMap, "bg_color": "#FFEB9C", "valign": "vcenter"}
        yellowCenterMap = {**yellowMap, "bold": True, "align": "center"}
        purpleMap = {**baseFontMap, "bg_color": "#E49EDD", "valign": "vcenter"}
        purpleCenterMap = {**purpleMap, "bold": True, "align": "center"}
        noColorMap = {**baseFontMap, "valign": "vcenter"}
        noColorCenterMap = {**noColorMap, "bold": True, "align": "center"}
        whiteBackgroundMap = {"bg_color": "white"}

        # width
        ws.set_column("A:A", 30.71) # 283 pixels
        ws.set_column("B:B", 135) # 1221 pixels

        # text
        ##
        blueFont = self.wb.add_format({**blueMap, "border": 1})
        ws.write(0, 0, "ERROR TYPE", blueFont)
        ws.write(0, 1, "RECOMMENDATION", blueFont)

        ##
        redCenterFont = self.wb.add_format({**redCenterMap, "top": 2, "bottom": 2, "right": 1})
        redFont = self.wb.add_format({**redMap, "border": 1})
        ws.merge_range("A2:B2", "RED HIGHLIGHTS", redCenterFont)
        ws.write(2, 0, "Identical File", redFont)
        ws.write(3, 0, "Old File", redFont)
        ws.write(4, 0, "Empty Directory", redFont)
        ws.write(5, 0, "Empty File", redFont)

        # Some of these are not set to wrap, cause then it messes with the height. It's weird.
        wrappedBorderFont = self.wb.add_format({**wrappedMap, "border": 1})
        baseBorderFont = self.wb.add_format({**baseFontMap, "border": 1})
        ws.write(2, 1, "Recommend deletion of identical files that are in the same folder. Change to [YELLOW HIGHLIGHTS] any files that you do not want deleted.", wrappedBorderFont)
        ws.write(3, 1, "If deemed transitory information and not information of business value (IBV), recommend deletion of files that have not been accessed for more than 1095 days / 3 years.  Change to [YELLOW HIGHLIGHTS] any files that you do not want deleted.", wrappedBorderFont)
        ws.write(4, 1, "Recommend deleting directories that contain less than 1 file and moving the less than 1 file to the parent directory until 3 or more files have accumulated. Change to [YELLOW HIGHLIGHTS] any directory that you do not want deleted and the less than 1 file moved to the parent directory.", wrappedBorderFont)
        ws.write(5, 1, "Recommend deleting files that are either empty, corrupt or unable to be accessed. Change to [YELLOW HIGHLIGHTS] any files that you do not want deleted.", baseBorderFont)

        ##
        yellowCenterFont = self.wb.add_format({**yellowCenterMap, "top": 2, "bottom": 2, "right": 1})
        yellowFont = self.wb.add_format({**yellowMap, "border": 1})
        ws.merge_range("A7:B7", "YELLOW HIGHLIGHTS", yellowCenterFont)
        ws.write(7, 0, "Identical File", yellowFont)
        ws.write(8, 0, "Multiple Version", yellowFont) # TODO: Here
        ws.write(9, 0, "Large File Type", yellowFont)
        ws.write(10, 0, "Old Files", yellowFont)
        ws.write(11, 0, "Empty Directory", yellowFont)

        ws.write(7, 1, "Recommend review of identical files in 3 or more folders. Change to [RED HIGHLIGHTS] for deletion.", wrappedBorderFont)
        ws.write(8, 1, "Recommend review of multiple file versions. Change to [RED HIGHLIGHTS] for deletion.", wrappedBorderFont) # TODO: Here
        ws.write(9, 1, "Recommend review of large files types. If deemed transitory information and not information of business value (IBV), change to  [RED HIGHLIGHTS] for deletion.  ", wrappedBorderFont)
        ws.write(10, 1, "Recommend review of files that have not been accessed for more than 730 days / two years but less than 1095 days / 3 years. If deemed transitory information and not information of business value (IBV), change to  [RED HIGHLIGHTS] for deletion.  ", wrappedBorderFont)
        ws.write(11, 1, "Recommend reviewing directories that contain 2 files for directory deletion and moving the 2 files to the parent directory until 3 or more files have accumulated. Change to [RED HIGHLIGHTS] any directories that you want deleted and the 2 files moved to the parent directory.", wrappedBorderFont)
        #for row in range(9, 11): ws.set_row(row, 30) # for setting height, if ever needed

        ##
        purpleCenterFont = self.wb.add_format({**purpleCenterMap, "top": 2, "bottom": 2, "right": 1})
        purpleFont = self.wb.add_format({**purpleMap, "border": 1})
        ws.merge_range("A13:B13", "PURPLE HIGHLIGHTS", purpleCenterFont)
        ws.write(13, 0, "Identical File", purpleFont)
        ws.write(13, 1, "Recommend reviewing the series of identical files that exist in a minimum of three distinct folders. Change to [RED HIGHLIGHTS] for deletion.", wrappedBorderFont)

        ##
        noColorCenterFont = self.wb.add_format({**noColorCenterMap, "top": 2, "bottom": 2, "right": 1})
        noColorFont = self.wb.add_format({**noColorMap, **wrappedMap, "border": 1})
        ws.merge_range("A15:B15", "NO HIGHLIGHTS", noColorCenterFont)
        ws.write(15, 0, "Identical File", noColorFont)
        ws.write(16, 0, "Replace Space with Hyphen", noColorFont)
        ws.write(17, 0, "Replace Bad Character", noColorFont)
        ws.write(18, 0, "Exceed Character Count", noColorFont)

        ws.write(15, 1, "Recommend review of identical files. Change to [RED HIGHLIGHTS] for deletion.", baseBorderFont)
        ws.write(16, 1, "Recommend replacing all spaces in directory and files names with hyphens (-) to align with PAEC file naming convention and IM best practices. Change to [YELLOW HIGHLIGHTS] any directory or file that you do not want modified.", wrappedBorderFont)
        ws.write(17, 1, "Recommend replacing all bad characters in directory and files names with hyphens (-) or other approproriate character to align with PAEC file naming convention and IM best practices. Change to [YELLOW HIGHLIGHTS] any directory or file that you do not want modified.", wrappedBorderFont)
        ws.write(18, 1, "Change to [YELLOW HIGHLIGHTS] priority files that you want to make sure are being backed up on the shared drive. Recommend directory and/or file names will be provided for your review before changes are made. ", wrappedBorderFont)

        ##
        whiteSpaceFont = self.wb.add_format({**whiteBackgroundMap, "top": 2})
        importantTopFont = self.wb.add_format({**wrappedMap, "bold": True, "top": 1, "left": 1, "right": 1})
        importantMiddleFont = self.wb.add_format({**wrappedMap, "bold": True, "left": 1, "right": 1})
        importantBottomFont = self.wb.add_format({**wrappedMap, "bold": True, "bottom": 1, "left": 1, "right": 1})
        ws.write(19, 0, "", whiteSpaceFont)
        ws.write(19, 1, "", whiteSpaceFont)
        ws.write(20, 1, "IMPORTANT - For Identical File, Large Files Type, Old File, Empty Directory, and Empty File; all files or folders that retain [RED HIGHIGHTS] will be modified, as per appropriate fix procedure.", importantTopFont)
        ws.write(21, 1, "IMPORTANT - For Replace Space with Hyphen, Bad Character, and Exceed Character Count; changing folder and file names will break existing links and shortcuts to the files or folders. Links and shortcuts will have to be updated.", importantMiddleFont)
        ws.write(22, 1, "IMPORTANT - For Exceed Character Count; files where the file path exceeds 200 characters are not being backed up on the shared drive.", importantBottomFont)

        # FRENCH
        ##
        ws.write(24, 0, "Français", self.wb.add_format({**baseFontBoldMap}))

        ##
        ws.write(25, 0, "TYPE D’ERREUR", blueFont)
        ws.write(25, 1, "RECOMMANDATION", blueFont)

        ##
        ws.merge_range("A27:B27", "SURBRILLANCE ROUGE", redCenterFont)
        ws.write(27, 0, "Fichier identique", redFont)
        ws.write(28, 0, "Ancien fichier", redFont)
        ws.write(29, 0, "Répertoire vide", redFont)
        ws.write(30, 0, "Dossier vide", redFont)

        ws.write(27, 1, "On recommande la suppression des fichiers identiques se trouvant dans le même dossier. Mettez en [SURBRILLANCE JAUNE] tous les fichiers que vous ne souhaitez pas voir supprimés.", wrappedBorderFont)
        ws.write(28, 1, "Si l’information est considérée comme transitoire et non comme une information à valeur opérationnelle (IVO), on recommande la suppression des fichiers qui n’ont pas été consultés depuis plus de 1095 jours/3 ans. Mettez-en [SURBRILLANCE JAUNE] tous les fichiers que vous ne souhaitez pas voir supprimés.", wrappedBorderFont)
        ws.write(29, 1, "On recommande la suppression des répertoires qui contiennent un fichier ou moins et le déplacement de ce fichier vers le répertoire parent jusqu’à ce que trois (3) fichiers ou plus soient regroupés. Mettez en [SURBRILLANCE JAUNE] tout répertoire que vous ne voulez pas supprimer, et dont vous ne voulez pas que le fichier ou moins qu’il contient soit déplacé vers le répertoire parent.", wrappedBorderFont)
        ws.write(30, 1, "On recommande la suppression des fichiers qui sont soit vides, soit corrompus, soit impossibles d’accès. Mettez en [SURBRILLANCE JAUNE] tous les fichiers que vous ne souhaitez pas voir supprimés.", wrappedBorderFont)

        ##
        ws.merge_range("A32:B32", "SURBRILLANCE JAUNE", yellowCenterFont)
        ws.write(32, 0, "Fichier identique", yellowFont)
        ws.write(33, 0, "Versions multiple", yellowFont) # TODO: Here
        ws.write(34, 0, "Type de fichiers volumineux", yellowFont)
        ws.write(35, 0, "Ancien fichier", yellowFont)
        ws.write(36, 0, "Répertoire vide", yellowFont)

        ws.write(32, 1, "On recommande l’examen des dossiers identiques dans trois (3) dossiers ou plus. Mettez les éléments en [SURBRILLANCE ROUGE] aux fins de suppression.", wrappedBorderFont)
        ws.write(33, 1, "Recommander la révision de plusieurs versions de fichiers. Changer en [HIGHLIGHTS ROUGES] pour suppression.", wrappedBorderFont) # TODO: Here
        ws.write(34, 1, "On recommande l’examen des types de fichiers volumineux. Si l’information est considérée comme transitoire et non comme une information à valeur opérationnelle (IVO), mettez l’élément en [SURBRILLANCE ROUGE] aux fins de suppression.  ", wrappedBorderFont)
        ws.write(35, 1, "On recommande l’examen des dossiers qui n’ont pas été consultés depuis plus de 730 jours/2 ans, mais moins de 1095 jours/3 ans. Si l’information est considérée comme transitoire et non comme une information à valeur opérationnelle (IVO), mettez l’élément en [SURBRILLANCE ROUGE] aux fins de suppression.  ", wrappedBorderFont)
        ws.write(36, 1, "On recommande l’examen des répertoires qui contiennent deux (2) fichiers aux fins de suppression et le déplacement de ces fichiers vers le répertoire parent jusqu’à ce que trois (3) fichiers ou plus soient regroupés. Mettez en [SURBRILLANCE ROUGE] tout répertoire que vous souhaitez voir supprimé et les deux (2) fichiers déplacés vers le répertoire parent.", wrappedBorderFont)

        ##
        ws.merge_range("A38:B38", "SURBRILLANCE MAUVE", purpleCenterFont)
        ws.write(38, 0, "Fichier identique", purpleFont)

        ws.write(38, 1, "On recommande d’examiner les séries de dossiers identiques qui existent dans au moins trois (3) dossiers distincts. Mettez l’élément en [SURBRILLANCE ROUGE] aux fins de suppression.", wrappedBorderFont)

        ##
        ws.merge_range("A40:B40", "AUCUNE SURBRILLANCE ", noColorCenterFont)
        ws.write(40, 0, "Fichier identique", noColorFont)
        ws.write(41, 0, "Remplacement des espaces par des traits d’union ", noColorFont)
        ws.write(42, 0, "Remplacement des caractères non acceptés", noColorFont)
        ws.write(43, 0, "Dépassement du nombre de caractères", noColorFont)

        ws.write(40, 1, "On recommande l’examen des dossiers identiques. Mettez les éléments en [SURBRILLANCE ROUGE] aux fins de suppression.", baseBorderFont)
        ws.write(41, 1, "On recommande de remplacer tous les espaces dans les noms de répertoires et de fichiers par des traits d’union (-) afin de respecter la convention d’appellation des fichiers de l’ECAP et les pratiques exemplaires en matière de gestion de l’information. Mettez en [SURBRILLANCE JAUNE] tout répertoire ou fichier que vous ne souhaitez pas modifier.", wrappedBorderFont)
        ws.write(42, 1, "On recommande de remplacer tous les caractères non acceptés dans les noms de répertoires et de fichiers par des traits d’union (-) ou d’autres caractères appropriés afin de respecter la convention d’appellation des fichiers de l’ECAP et les pratiques exemplaires en matière de gestion de l’information. Mettez en [SURBRILLANCE JAUNE] tout répertoire ou fichier que vous ne souhaitez pas modifier.", wrappedBorderFont)
        ws.write(43, 1, "Mettez en [SURBRILLANCE JAUNE] les fichiers prioritaires qui, à votre avis, doivent être sauvegardés sur le disque partagé. Les noms des répertoires ou des fichiers recommandés vous seront fournis aux fins d’examen avant que les changements ne soient effectués. ", wrappedBorderFont)

        ##
        ws.write(44, 0, "", whiteSpaceFont)
        ws.write(44, 1, "", whiteSpaceFont)
        ws.write(45, 1, "IMPORTANT - Pour Fichier identique, Type de fichiers volumineux, Ancien fichier, Répertoire vide et Fichier vide, tous les fichiers ou dossiers qui sont encore en [SURBRILLANCE ROUGE] seront modifiés, conformément à la procédure de correction appropriée.", importantTopFont)
        ws.write(46, 1, "IMPORTANT - Pour Remplacement des espaces par des traits d’union, Remplacement des caractères non acceptés, Dépassement du nombre de caractères, la modification des noms de dossiers et de fichiers entraînera la rupture des liens et des raccourcis existants vers les fichiers ou les dossiers. Les liens et les raccourcis devront être mis à jour.", importantMiddleFont)
        ws.write(47, 1, "IMPORTANT - Pour Dépassement du nombre de caractères, les fichiers dont le chemin d’accès dépasse 200 caractères ne sont pas sauvegardés sur le disque partagé.", importantBottomFont)

        #


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
        self.summarySheet.write(2, 0, "Excluded Extensions", self.headerFormat)
        self.summarySheet.write(3, 0, "Include Subdirectories", self.headerFormat)
        self.summarySheet.write(4, 0, "Enable Modifications", self.headerFormat)
        self.summarySheet.write(5, 0, "Include Hidden Files", self.headerFormat)
        self.summarySheet.write(6, 0, "Add Recommendations", self.headerFormat)
        self.summarySheet.write(7, 0, "Argument(s)", self.headerFormat)
        self.summarySheet.write(9, 0, "Directory count", self.headerFormat)
        self.summarySheet.write(10, 0, "Directory error count / %", self.headerFormat)
        self.summarySheet.write(11, 0, "File count", self.headerFormat)
        self.summarySheet.write(12, 0, "File error count / %", self.headerFormat)
        self.summarySheet.write(13, 0, "Execution time (s)", self.headerFormat)

        self.summarySheet.write(0, 1, dirAbsolute, self.summaryValueFormat)
        if self.excludedExtensions:
            self.summarySheet.write(2, 1, str(self.excludedExtensions), self.summaryValueFormat)
        self.summarySheet.write(3, 1, str(includeSubdirectories), self.summaryValueFormat)
        self.summarySheet.write(4, 1, str(allowModify), self.summaryValueFormat)
        self.summarySheet.write(5, 1, str(includeHiddenFiles), self.summaryValueFormat)
        self.summarySheet.write(6, 1, str(addRecommendations), self.summaryValueFormat)

        # HARD CODED atip stuff
        # which is okay because the code can't get here without this being agreed upon. Therefore, don't need a separate variable.
        if allowModify:
            self.summarySheet.write(8, 0, "NO LIT-HOLD / ATIP", self.headerFormat)
            self.summarySheet.write(8, 1, str(True))


    def populateSummarySheet(self, excludedDirs):
        col = 1

        for procedureObject in self.procedureObjects:
            arg = self.procedureObjectArgs[procedureObject]
            
            if arg == ():
                continue
            
            self.summarySheet.write(7, col, f"{arg[0] if len(arg) <= 1 else arg} : {procedureObject.name}", self.summaryValueFormat)
            col += 1

        if self.filesScannedCount == 0: fileErrorPercentage = 0
        else: fileErrorPercentage = round(self.fileErrorCount / self.filesScannedCount * 100, 2)
        if self.foldersScannedCount == 0: folderErrorPercentage = 0
        else: folderErrorPercentage = round(self.folderErrorCount / self.foldersScannedCount * 100, 2)

        self.summarySheet.write_number(9, 1, self.foldersScannedCount, self.summaryValueFormat)
        self.summarySheet.write_number(10, 1, self.folderErrorCount, self.summaryValueFormat)
        self.summarySheet.write(10, 2, "{}%".format(folderErrorPercentage), self.summaryValueFormat)
        self.summarySheet.write_number(11, 1, self.filesScannedCount, self.summaryValueFormat)
        self.summarySheet.write_number(12, 1, self.fileErrorCount, self.summaryValueFormat)
        self.summarySheet.write(12, 2, "{}%".format(fileErrorPercentage), self.summaryValueFormat)
        self.summarySheet.write_number(13, 1, round(self.executionTime, 4), self.summaryValueFormat)

        i = 1
        for exDir in excludedDirs:
            self.summarySheet.write_string(1, i, exDir, self.summaryValueFormat)
            i += 1

        i = 0
        for ws in self.getAllProcedureSheets():
            self.summarySheet.write(self.sheetRows[self.summarySheet] + i, 1, self.summaryCounts[ws], self.summaryValueFormat)
            i += 1


    def close(self):
        self.wb.close()
