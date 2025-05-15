from procedureFunctions import *
from procedureHeader import *
from defs import *


ALL_PROCEDURES = {
    LIST_ALL: FindProcedure(
        LIST_ALL,
        True,
        listAll,
        startFunction=writeDefaultHeaders
    ),

    LIST_ALL_OWNER: FindProcedure(
        LIST_ALL_OWNER,
        True,
        listAllOwner,
        startFunction=writeDefaultAndOwnerHeaders
    ),

    IDENTICAL_FILE: FindProcedure(
        IDENTICAL_FILE,
        True,
        duplicateContentConcurrent,
        startFunction=duplicateContentStart,
        postFunction=duplicateContentPost,
        recommendPostFunction=duplicateContentPostRecommend,
        isConcurrentOnly=False
    ),

    FILE_EXTENSION_SUMMARY: FindProcedure(
        FILE_EXTENSION_SUMMARY,
        True,
        fileExtensionConcurrent,
        startFunction=fileExtensionStart,
        postFunction=fileExtensionPost,
        isConcurrentOnly=False
    ),

    OLD_FILE: FindProcedure(
        OLD_FILE,
        True,
        oldFileFind,
        validatorFunction=minimumIntToInfinityOrMaxValidator,
        argBoundary=1,
        defaultArgument=(1095,),
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
        startFunction=emptyDirectoryStart,
        recommendBaseFunction=emptyDirectoryRecommend
    ),

    EMPTY_FILE: FindProcedure(
        EMPTY_FILE,
        True,
        emptyFileFind,
        startFunction=writeDefaultAndOwnerHeaders,
        recommendBaseFunction=emptyFileFindRecommend
    ),

    SPACE_FILE_FIND: FindProcedure(
        SPACE_FILE_FIND,
        True,
        spaceFileFind,
        startFunction=writeDefaultHeaders
    ),

    SPACE_DIRECTORY_FIND: FindProcedure(
        SPACE_DIRECTORY_FIND,
        False,
        spaceFolderFind,
        startFunction=writeDefaultHeaders
    ),

    BAD_CHARACTER_DIRECTORY: FindProcedure(
        BAD_CHARACTER_DIRECTORY,
        False,
        badCharFolderFind,
        startFunction=writeDefaultHeaders
    ),

    BAD_CHARACTER_FILE: FindProcedure(
        BAD_CHARACTER_FILE,
        True,
        badCharFileFind,
        startFunction=writeDefaultHeaders
    ),

    CHARACTER_LIMIT: FindProcedure(
        CHARACTER_LIMIT,
        True,
        overCharLimitFind,
        startFunction=writeDefaultHeaders
    ),

    DELETE_EMPTY_FILE: FixProcedure(
        DELETE_EMPTY_FILE,
        True,
        deleteEmptyFilesLog,
        deleteEmptyFilesModify,
        startFunction=deleteEmptyFilesStart,
        recommendBaseFunction=deleteEmptyFilesRecommendLog
    ),

    SPACE_FILE_FIX: FixProcedure(
        SPACE_FILE_FIX,
        True,
        spaceFileFixLog,
        spaceFileFixModify,
        startFunction=writeDefaultHeaders
    ),

    SPACE_FOLDER_FIX: FixProcedure(
        SPACE_FOLDER_FIX,
        False,
        spaceFolderFixLog,
        spaceFolderFixModify,
        startFunction=writeDefaultHeaders,
        postFunction=fixfolderModifyPost
    ),

    REPLACE_CHARACTER_FOLDER: FixProcedure(
        REPLACE_CHARACTER_FOLDER,
        False,
        searchAndReplaceFolderLog,
        searchAndReplaceFolderModify,
        startFunction=searchAndReplaceFolderStart,
        validatorFunction=multiplePairsOfStringsValidator,
        argBoundary=">",
        postFunction=fixfolderModifyPost
    ),

    REPLACE_CHARACTER_FILE: FixProcedure(
        REPLACE_CHARACTER_FILE,
        True,
        searchAndReplaceFileLog,
        searchAndReplaceFileModify,
        startFunction=searchAndReplaceFileStart,
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