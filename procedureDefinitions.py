from findFunctions import *
from fixFunctions import *
from procedureHeader import *
from defs import *


FIND_PROCEDURES = {
    LIST_ALL: FindProcedure(
        LIST_ALL,
        True,
        listAll,
    ),

    LIST_ALL_OWNER: FindProcedure(
        LIST_ALL_OWNER,
        True,
        listAllOwner,
        startFunction=writeOwnerHeader
    ),

    IDENTICAL_FILE: FindProcedure(
        IDENTICAL_FILE,
        True,
        duplicateContentConcurrent,
        startFunction=duplicateContentStart,
        postFunction=duplicateContentPost,
        recommendPostFunction=duplicateContentPostRecommend
    ),

    FILE_EXTENSION_SUMMARY: FindProcedure(
        FILE_EXTENSION_SUMMARY,
        True,
        fileExtensionConcurrent,
        startFunction=fileExtensionStart,
        postFunction=fileExtensionPost
    ),

    OLD_FILE: FindProcedure(
        OLD_FILE,
        True,
        oldFileFind,
        validatorFunction=minimumIntToInfinityOrMaxValidator,
        argBoundary=1,
        defaultArgument=(1095,),
        columnName="# Days Last Accessed",
        startFunction=oldFileFindStart,
        recommendBaseFunction=oldFileFindRecommend
    ),

    EMPTY_DIRECTORY: FindProcedure(
        EMPTY_DIRECTORY,
        False,
        emptyDirectory,
        validatorFunction=minimumIntToInfinityValidator,
        argBoundary=0,
        defaultArgument=(0,),
        columnName="# Files Contained",
        startFunction=emptyDirectoryStart,
        recommendBaseFunction=emptyDirectoryRecommend
    ),

    EMPTY_FILE: FindProcedure(
        EMPTY_FILE,
        True,
        emptyFileFind,
        startFunction=writeOwnerHeader,
        recommendBaseFunction=emptyFileFindRecommend
    ),

    SPACE_FILE_FIND: FindProcedure(
        SPACE_FILE_FIND,
        True,
        spaceFileFind,
    ),

    SPACE_DIRECTORY_FIND: FindProcedure(
        SPACE_DIRECTORY_FIND,
        False,
        spaceFolderFind,
    ),

    BAD_CHARACTER_DIRECTORY: FindProcedure(
        BAD_CHARACTER_DIRECTORY,
        False,
        badCharFolderFind,
    ),

    BAD_CHARACTER_FILE: FindProcedure(
        BAD_CHARACTER_FILE,
        True,
        badCharFileFind,
    ),

    CHARACTER_LIMIT: FindProcedure(
        CHARACTER_LIMIT,
        True,
        overCharLimitFind,
    ),
}

FIX_PROCEDURES = {
    DELETE_EMPTY_FILE: FixProcedure(
        DELETE_EMPTY_FILE,
        True,
        deleteEmptyFilesLog,
        deleteEmptyFilesModify,
        columnName="Staged for Deletion",
        startFunction=deleteEmptyFilesStart,
        recommendBaseFunction=deleteEmptyFilesRecommendLog
    ),

    SPACE_FILE_FIX: FixProcedure(
        SPACE_FILE_FIX,
        True,
        spaceFileFixLog,
        spaceFileFixModify,
    ),

    SPACE_FOLDER_FIX: FixProcedure(
        SPACE_FOLDER_FIX,
        False,
        spaceFolderFixLog,
        spaceFolderFixModify,
        postFunction=fixfolderModifyPost
    ),

    REPLACE_CHARACTER_FOLDER: FixProcedure(
        REPLACE_CHARACTER_FOLDER,
        False,
        searchAndReplaceFolderLog,
        searchAndReplaceFolderModify,
        validatorFunction=multiplePairsOfStringsValidator,
        argBoundary=">",
        postFunction=fixfolderModifyPost
    ),

    REPLACE_CHARACTER_FILE: FixProcedure(
        REPLACE_CHARACTER_FILE,
        True,
        searchAndReplaceFileLog,
        searchAndReplaceFileModify,
        validatorFunction=multiplePairsOfStringsValidator,
        argBoundary=">"
    ),
}


# NOTE: The functions that require an argument must still be in the order as set in FIX_PROCEDURES_DISPLAY below
# This is so the "/" separated list of arguments from the "Parameter" field are properly distributed.
FIND_PROCEDURES_DISPLAY = [
    LIST_ALL,
    LIST_ALL_OWNER,
    IDENTICAL_FILE,
    FILE_EXTENSION_SUMMARY,
    OLD_FILE,
    EMPTY_DIRECTORY,
    EMPTY_FILE,
    SPACE_DIRECTORY_FIND,
    SPACE_FILE_FIND,
    BAD_CHARACTER_DIRECTORY,
    BAD_CHARACTER_FILE,
    CHARACTER_LIMIT
]

FIX_PROCEDURES_DISPLAY = [
    "",
    "",
    "",
    "",
    "",
    "",
    DELETE_EMPTY_FILE,
    SPACE_FOLDER_FIX,
    SPACE_FILE_FIX,
    REPLACE_CHARACTER_FOLDER,
    REPLACE_CHARACTER_FILE
]