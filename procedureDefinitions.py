from procedureFunctions import *
from procedureClass import *
from defs import *


ALL_PROCEDURES = {
### FIND PROCEDURES ################################################################################
    LIST_ALL: Procedure(
        LIST_ALL,
        True,
        listAllBase,
        startFunction=writeDefaultHeaders
    ),

    LIST_ALL_OWNER: Procedure(
        LIST_ALL_OWNER,
        True,
        listAllOwnerBase,
        startFunction=writeDefaultAndOwnerHeaders
    ),

    IDENTICAL_FILE: Procedure(
        IDENTICAL_FILE,
        True,
        identicalFileBase,
        startFunction=identicalFileStart,
        postFunction=identicalFilePost,
        recommendPostFunction=identicalFilePostRecommend,
        isConcurrentOnly=False
    ),

    FILE_EXTENSION_SUMMARY: Procedure(
        FILE_EXTENSION_SUMMARY,
        True,
        fileExtensionSummaryBase,
        startFunction=fileExtensionSummaryStart,
        postFunction=fileExtensionSummaryPost,
        isConcurrentOnly=False
    ),

    OLD_FILE: Procedure(
        OLD_FILE,
        True,
        oldFileBase,
        validatorFunction=minimumIntToInfinityOrMaxValidator,
        argBoundary=1,
        defaultArgument=(1095,),
        startFunction=oldFileStart,
        recommendBaseFunction=oldFileRecommend
    ),

    EMPTY_DIRECTORY: Procedure(
        EMPTY_DIRECTORY,
        False,
        emptyDirectoryBase,
        validatorFunction=minimumIntToInfinityValidator,
        argBoundary=0,
        defaultArgument=(0,),
        startFunction=emptyDirectoryStart,
        recommendBaseFunction=emptyDirectoryRecommend
    ),

    EMPTY_FILE: Procedure(
        EMPTY_FILE,
        True,
        emptyFileFindBase,
        startFunction=writeDefaultAndOwnerHeaders,
        recommendBaseFunction=emptyFileFindRecommend
    ),

    SPACE_FILE_FIND: Procedure(
        SPACE_FILE_FIND,
        True,
        spaceFileFindBase,
        startFunction=writeDefaultHeaders
    ),

    SPACE_DIRECTORY_FIND: Procedure(
        SPACE_DIRECTORY_FIND,
        False,
        spaceFolderFindBase,
        startFunction=writeDefaultHeaders
    ),

    BAD_CHARACTER_DIRECTORY: Procedure(
        BAD_CHARACTER_DIRECTORY,
        False,
        badCharacterFolderFind,
        startFunction=badCharacterStart
    ),

    BAD_CHARACTER_FILE: Procedure(
        BAD_CHARACTER_FILE,
        True,
        badCharacterFileFind,
        startFunction=badCharacterStart
    ),

    CHARACTER_LIMIT: Procedure(
        CHARACTER_LIMIT,
        True,
        exceedCharacterLimitBase,
        startFunction=writeDefaultHeaders
    ),

### FIX PROCEDURES ################################################################################
    DELETE_EMPTY_FILE: Procedure(
        DELETE_EMPTY_FILE,
        True,
        deleteEmptyFileBase,
        deleteEmptyFileModify,
        startFunction=deleteEmptyFileStart,
        recommendBaseFunction=deleteEmptyFileRecommend
    ),

    SPACE_FILE_FIX: Procedure(
        SPACE_FILE_FIX,
        True,
        spaceFileFixBase,
        spaceFileFixModify,
        startFunction=writeDefaultHeaders
    ),

    SPACE_FOLDER_FIX: Procedure(
        SPACE_FOLDER_FIX,
        False,
        spaceFolderFixBase,
        spaceFolderFixModify,
        startFunction=spaceFolderFixStart,
        postFunction=fixfolderModifyPost
    ),

    REPLACE_CHARACTER_FOLDER: Procedure(
        REPLACE_CHARACTER_FOLDER,
        False,
        searchAndReplaceFolderBase,
        searchAndReplaceFolderModify,
        startFunction=searchAndReplaceFolderStart,
        validatorFunction=multiplePairsOfStringsValidator,
        argBoundary=">",
        postFunction=fixfolderModifyPost
    ),

    REPLACE_CHARACTER_FILE: Procedure(
        REPLACE_CHARACTER_FILE,
        True,
        searchAndReplaceFileBase,
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