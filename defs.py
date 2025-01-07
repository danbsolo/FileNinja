from findFixFunctions import *
from findFixHeader import *

RESULTS_DIRECTORY = "fileCrawlerResults"

LIST_ALL = "ListAll"
CHARACTER_LIMIT_FIND = "CharLimit-Find"
BAD_CHARACTER_FIND = "BadChar-Find"
SPACE_FIND = "SPC-Find"
FILE_EXTENSION_SUMMARY = "FileExt-Summary"
DUPLICATE_FILE_FIND = "DupFile-Find"

NULL_OPTION = "..."
SPACE_FIX = "SPC-Fix"
DELETE_OLD_FILES = "DelOldFiles-Fix"
DELETE_EMPTY_DIRECTORIES_FIX = "DelEmptyDirs-Fix"
SEARCH_AND_REPLACE = "SearchAndReplace-Fix"

FIND_PROCEDURES = {
    LIST_ALL: FindProcedure(LIST_ALL, listAll, True),
    CHARACTER_LIMIT_FIND: FindProcedure(CHARACTER_LIMIT_FIND, overCharLimitFind, True),
    BAD_CHARACTER_FIND: FindProcedure(BAD_CHARACTER_FIND, badCharFind, True),
    SPACE_FIND: FindProcedure(SPACE_FIND, spaceFind, True),
    DUPLICATE_FILE_FIND: FindProcedure(DUPLICATE_FILE_FIND, duplicateFileConcurrent, False, duplicateFilePost),
    FILE_EXTENSION_SUMMARY: FindProcedure(FILE_EXTENSION_SUMMARY, fileExtensionConcurrent, False, fileExtensionPost)
}

FIX_PROCEDURES = {
    NULL_OPTION: None,
    SPACE_FIX: FixProcedure(SPACE_FIX, spaceFixLog, spaceFixModify, True),
    DELETE_OLD_FILES: FixProcedure(DELETE_OLD_FILES, deleteOldFilesLog, deleteOldFilesModify, True, minimumIntegerValidator, 1, "# Days Last Accessed"),
    DELETE_EMPTY_DIRECTORIES_FIX: FixProcedure(DELETE_EMPTY_DIRECTORIES_FIX, deleteEmptyDirectoriesLog, deleteEmptyDirectoriesModify, False, minimumIntegerValidator, 0, "# Files Contained"),
    SEARCH_AND_REPLACE: FixProcedure(SEARCH_AND_REPLACE, searchAndReplaceLog, searchAndReplaceModify, True, twoStringsValidator, "~")
}