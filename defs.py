from findFixFunctions import *
from findFixHeader import *

RESULTS_DIRECTORY = "fileCrawlerResults"
FILE_CRAWLER = "FileCrawler"

# NOTE: Excel worksheet name must be <= 31 characters
LIST_ALL = "List All Items"
SPACE_FIND = "Space Error"
BAD_CHARACTER_FIND = "Bad Character Error"
OLD_FILE_FIND = "Old File Error"
EMPTY_DIRECTORY_FIND = "Empty Directory"
EMPTY_FILE_FIND = "Empty File"
CHARACTER_LIMIT_FIND = "Character Limit Error"
DUPLICATE_CONTENT_FIND = "Identical Files Error"
DUPLICATE_NAME_FIND = "Duplicate Names Error"
FILE_EXTENSION_SUMMARY = "Summarize File Types"

NULL_OPTION = "---"
SPACE_FIX = "Replace Space with Hyphen"
SEARCH_AND_REPLACE = "Search & Replace Bad Characters"
DELETE_OLD_FILES = "Delete Old Files"
DELETE_EMPTY_DIRECTORIES_FIX = "Delete Empty Directories"
DELETE_EMPTY_FILES = "Delete Empty Files"

FIND_PROCEDURES = {
    LIST_ALL: FindProcedure(LIST_ALL, listAll, True),
    SPACE_FIND: FindProcedure(SPACE_FIND, spaceFind, True),
    BAD_CHARACTER_FIND: FindProcedure(BAD_CHARACTER_FIND, badCharFind, True),
    OLD_FILE_FIND: FindProcedure(OLD_FILE_FIND, oldFileFind, True),
    EMPTY_DIRECTORY_FIND: FindProcedure(EMPTY_DIRECTORY_FIND, emptyDirectoryConcurrent, True, emptyDirectoryPost, False),
    EMPTY_FILE_FIND: FindProcedure(EMPTY_FILE_FIND, emptyFileFind, True),
    CHARACTER_LIMIT_FIND: FindProcedure(CHARACTER_LIMIT_FIND, overCharLimitFind, True),
    DUPLICATE_CONTENT_FIND:FindProcedure(DUPLICATE_CONTENT_FIND, duplicateContentConcurrent, False, duplicateContentPost),
    DUPLICATE_NAME_FIND: FindProcedure(DUPLICATE_NAME_FIND, duplicateNameConcurrent, False, duplicateNamePost),
    FILE_EXTENSION_SUMMARY: FindProcedure(FILE_EXTENSION_SUMMARY, fileExtensionConcurrent, False, fileExtensionPost)
}

FIX_PROCEDURES = {
    NULL_OPTION: None,
    SPACE_FIX: FixProcedure(SPACE_FIX, spaceFixLog, spaceFixModify, True),
    SEARCH_AND_REPLACE: FixProcedure(SEARCH_AND_REPLACE, searchAndReplaceLog, searchAndReplaceModify, True, twoStringsValidator, "~"),
    DELETE_OLD_FILES: FixProcedure(DELETE_OLD_FILES, deleteOldFilesLog, deleteOldFilesModify, True, minimumIntegerValidator, 1, "# Days Last Accessed"),
    DELETE_EMPTY_DIRECTORIES_FIX: FixProcedure(DELETE_EMPTY_DIRECTORIES_FIX, deleteEmptyDirectoriesLog, deleteEmptyDirectoriesModify, False, minimumIntegerValidator, 0, "# Files Contained"),
    DELETE_EMPTY_FILES: FixProcedure(DELETE_EMPTY_FILES, deleteEmptyFilesLog, deleteEmptyFilesModify, True, columnName="Staged for Deletion") # WIP
}
