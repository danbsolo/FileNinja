# from procedureFunctions import *
from findFunctions import *
from fixFunctions import *
from procedureHeader import *
from defs import *

FIND_PROCEDURES = {
    LIST_ALL: FindProcedure(LIST_ALL, listAll, True),
    SPACE_FILE_FIND: FindProcedure(SPACE_FILE_FIND, spaceFileFind, True),
    BAD_CHARACTER_FILE_FIND: FindProcedure(BAD_CHARACTER_FILE_FIND, badCharFileFind, True),
    OLD_FILE_FIND: FindProcedure(OLD_FILE_FIND, oldFileFind, True, startFunction=writeOwnerHeader, recommendMainFunction=oldFileFindRecommend, validatorFunction=minimumIntToInfinityOrMaxValidator, argBoundary=1),
    EMPTY_DIRECTORY_FIND: FindProcedure(EMPTY_DIRECTORY_FIND, emptyDirectory, True, isFileFind=False, recommendMainFunction=emptyDirectoryRecommend),
    EMPTY_FILE_FIND: FindProcedure(EMPTY_FILE_FIND, emptyFileFind, True, startFunction=writeOwnerHeader, recommendMainFunction=emptyFileFindRecommend),
    SPACE_FOLDER_FIND: FindProcedure(SPACE_FOLDER_FIND, spaceFolderFind, True, isFileFind=False),
    BAD_CHARACTER_FOLDER_FIND: FindProcedure(BAD_CHARACTER_FOLDER_FIND, badCharFolderFind, True, isFileFind=False),
    CHARACTER_LIMIT_FIND: FindProcedure(CHARACTER_LIMIT_FIND, overCharLimitFind, True),
    IDENTICAL_FILE_ERROR:FindProcedure(IDENTICAL_FILE_ERROR, duplicateContentConcurrent, False, duplicateContentPost, startFunction=duplicateContentStart, recommendPostFunction=duplicateContentPostRecommend),
    FILE_EXTENSION_SUMMARY: FindProcedure(FILE_EXTENSION_SUMMARY, fileExtensionConcurrent, False, fileExtensionPost, startFunction=fileExtensionStart),
    LIST_ALL_OWNER: FindProcedure(LIST_ALL_OWNER, listAllOwner, True, startFunction=writeOwnerHeader),
}

FIND_PROCEDURES_DISPLAY = [
    LIST_ALL,
    LIST_ALL_OWNER,
    IDENTICAL_FILE_ERROR,
    FILE_EXTENSION_SUMMARY,
    OLD_FILE_FIND,
    EMPTY_DIRECTORY_FIND,
    EMPTY_FILE_FIND,
    SPACE_FOLDER_FIND,
    SPACE_FILE_FIND,
    BAD_CHARACTER_FOLDER_FIND,
    BAD_CHARACTER_FILE_FIND,
    CHARACTER_LIMIT_FIND
]


# NOTE: The functions that require an argument must still be in the order as set in FIX_PROCEDURES_DISPLAY below
# This is so the "/" separated list of arguments from the "Parameter" field are properly distributed.
FIX_PROCEDURES = {
    SPACE_FILE_FIX: FixProcedure(SPACE_FILE_FIX, spaceFileFixLog, spaceFileFixModify, True),
    DELETE_OLD_FILES: FixProcedure(DELETE_OLD_FILES, deleteOldFilesLog, deleteOldFilesModify, True, minimumIntToInfinityOrMaxValidator, 1, "# Days Last Accessed", startFunction=deleteOldFilesStart, recommendLogFunction=deleteOldFilesRecommendLog),
    DELETE_EMPTY_DIRECTORIES_FIX: FixProcedure(DELETE_EMPTY_DIRECTORIES_FIX, deleteEmptyDirectoriesLog, deleteEmptyDirectoriesModify, False, minimumIntToInfinityValidator, 0, "# Files Contained", recommendLogFunction=deleteEmptyDirectoriesRecommendLog),
    DELETE_EMPTY_FILES: FixProcedure(DELETE_EMPTY_FILES, deleteEmptyFilesLog, deleteEmptyFilesModify, True, columnName="Staged for Deletion", startFunction=deleteEmptyFilesStart, recommendLogFunction=deleteEmptyFilesRecommendLog),
    SEARCH_AND_REPLACE_FOLDER: FixProcedure(SEARCH_AND_REPLACE_FOLDER, searchAndReplaceFolderLog, searchAndReplaceFolderModify, False, multiplePairsOfStringsValidator, ">", postFunction=fixfolderModifyPost),
    SEARCH_AND_REPLACE_FILE: FixProcedure(SEARCH_AND_REPLACE_FILE, searchAndReplaceFileLog, searchAndReplaceFileModify, True, multiplePairsOfStringsValidator, ">"),
    SPACE_FOLDER_FIX: FixProcedure(SPACE_FOLDER_FIX, spaceFolderFixLog, spaceFolderFixModify, False, postFunction=fixfolderModifyPost),
    DELETE_IDENTICAL_FILES: FixProcedure(DELETE_IDENTICAL_FILES, deleteIdenticalFilesLogConcurrent, deleteIdenticalFilesLogConcurrent, True, startFunction=deleteIdenticalFilesStart, postFunction=deleteIdenticalFilesPost, recommendPostFunction=deleteIdenticalFilesRecommendPost)
}

FIX_PROCEDURES_DISPLAY = [
    "",
    "",
    DELETE_IDENTICAL_FILES,
    "",
    DELETE_OLD_FILES,
    DELETE_EMPTY_DIRECTORIES_FIX,
    DELETE_EMPTY_FILES,
    SPACE_FOLDER_FIX,
    SPACE_FILE_FIX,
    SEARCH_AND_REPLACE_FOLDER,
    SEARCH_AND_REPLACE_FILE
]