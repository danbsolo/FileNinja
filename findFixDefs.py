from findFixMethods import *
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

FIND_METHODS = {
    LIST_ALL: FindMethod(LIST_ALL, listAll, True),
    CHARACTER_LIMIT_FIND: FindMethod(CHARACTER_LIMIT_FIND, overCharLimitFind, True),
    BAD_CHARACTER_FIND: FindMethod(BAD_CHARACTER_FIND, badCharFind, True),
    SPACE_FIND: FindMethod(SPACE_FIND, spaceFind, True),
    DUPLICATE_FILE_FIND: FindMethod(DUPLICATE_FILE_FIND, duplicateFileConcurrent, False, duplicateFilePost),
    FILE_EXTENSION_SUMMARY: FindMethod(FILE_EXTENSION_SUMMARY, fileExtensionConcurrent, False, fileExtensionPost)
}

FIX_METHODS = {
    NULL_OPTION: None,
    SPACE_FIX: FixMethod(SPACE_FIX, spaceFixLog, spaceFixModify, "Modification", True),
    DELETE_OLD_FILES: FixMethod(DELETE_OLD_FILES, deleteOldFilesLog, deleteOldFilesModify, "Days Old", True, minimumIntegerValidator, 1),
    DELETE_EMPTY_DIRECTORIES_FIX: FixMethod(DELETE_EMPTY_DIRECTORIES_FIX, deleteEmptyDirectoriesLog, deleteEmptyDirectoriesModify, "# Files Contained", False, minimumIntegerValidator, 0)
}