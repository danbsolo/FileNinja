from FileNinjaSuite.FileNinja.procedureFunctions import *
from FileNinjaSuite.FileNinja.procedureClass import *
from FileNinjaSuite.Shared.sharedDefs import *

# TODO: Partition this file into multiple defs.

TITLE = "File-Ninja"
RESULTS_DIRECTORY = TITLE + "-Results"
README_ADMIN = "README-Admin.docx"
README_LITE = "README-Lite.docx"
SHARED_UTILS_PATH = "..\\shared\\"

# NOTE: Excel worksheet name must be <= 31 characters
# The following characters are also invalid: / \ ? * : [ ]
LIST_ALL = "List All Files"
LIST_ALL_OWNER = "List All Files (Owner)"
SPACE_FILE_FIND = "Space Error (FILE)"
SPACE_DIRECTORY_FIND = "Space Error (DIR)"
BAD_CHARACTER_FILE = "Bad Character (FILE)"
BAD_CHARACTER_DIRECTORY = "Bad Character (DIR)"
OLD_FILE = "Old File#~"
EMPTY_DIRECTORY = "Empty Directory#~"
EMPTY_FILE = "Empty File~"
CHARACTER_LIMIT = "Exceed Character Limit"
IDENTICAL_FILE = "Identical File~"
FILE_EXTENSION_SUMMARY = "Large File Size#"
MULTIPLE_VERSION = "Multiple Version#"
SEARCH_FUNCTION_FILE = "Search Function (FILE)#"
SEARCH_FUNCTION_DIRECTORY = "Search Function (DIR)#"

DELETE_EMPTY_FILE = "Delete Empty File~"
SPACE_FILE_FIX = "Replace Space w Hyphen (FILE)"
SPACE_FOLDER_FIX = "Replace Space w Hyphen (DIR)"
REPLACE_CHARACTER_FILE = "Replace Character (FILE)#"
REPLACE_CHARACTER_FOLDER = "Replace Character (DIR)#"

DIR_ABSOLUTE_KEY = "dirAbsolute"
EXCLUDED_DIRS_KEY = "excludedDirs"
EXCLUDED_EXTENSIONS_KEY = "excludedExtensions"
SELECTED_FIND_PROCEDURES_KEY = "selectedFindProcedures"
SELECTED_FIX_PROCEDURES_KEY = "selectedFixProcedures"
ARG_UNPROCESSED_KEY = "argUnprocessed"
INCLUDE_SUBDIRECTORIES_KEY = "includeSubdirectories"
ENABLE_MODIFICATIONS_KEY = "enableModifications"
INCLUDE_HIDDEN_FILES_KEY = "includeHiddenFiles"
ADD_RECOMMENDATIONS_KEY = "addRecommendations"
JSON_KEYS = [
    DIR_ABSOLUTE_KEY,
    EXCLUDED_DIRS_KEY,
    SELECTED_FIND_PROCEDURES_KEY,
    SELECTED_FIX_PROCEDURES_KEY,
    ARG_UNPROCESSED_KEY,
    INCLUDE_SUBDIRECTORIES_KEY,
    ENABLE_MODIFICATIONS_KEY,
    INCLUDE_HIDDEN_FILES_KEY,
    ADD_RECOMMENDATIONS_KEY,
    EXCLUDED_EXTENSIONS_KEY
]


### PROCEDURE DEFINITIONS ###
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
        startFunction=writeDefaultAndOwnerAndLastModifiedHeaders
    ),

    IDENTICAL_FILE: Procedure(
        IDENTICAL_FILE,
        True,
        identicalFileBase,
        isConcurrentOnly=False,
        startFunction=identicalFileStart,
        postFunction=identicalFilePost,
        recommendPostFunction=identicalFilePostRecommend,

    ),

    FILE_EXTENSION_SUMMARY: Procedure(
        FILE_EXTENSION_SUMMARY,
        True,
        fileExtensionSummaryBase,
        isConcurrentOnly=False,
        validatorFunction=minimumIntToInfinityValidator,
        argBoundary=0,
        defaultArgument=(100,),
        startFunction=fileExtensionSummaryStart,
        postFunction=fileExtensionSummaryPost,
        recommendPostFunction=fileExtensionSummaryPostRecommend,

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
        startFunction=writeDefaultAndOwnerAndLastModifiedHeaders,
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
        startFunction=exceedCharacterLimitStart
    ),

    MULTIPLE_VERSION: Procedure(
        MULTIPLE_VERSION,
        True,
        multipleVersionBase,
        validatorFunction=numberRangeValidator,
        argBoundary=(1, 100),
        defaultArgument=(98,),
        isConcurrentOnly=False,
        startFunction=multipleVersionStart,
        postFunction=multipleVersionPost,
        recommendPostFunction=multipleVersionPostRecommend
    ),

    SEARCH_FUNCTION_FILE: Procedure(
        SEARCH_FUNCTION_FILE,
        True,
        searchFunctionFileBase,
        validatorFunction=multipleStringValidator,
        startFunction=searchFunctionFileStart
    ),

    SEARCH_FUNCTION_DIRECTORY: Procedure(
        SEARCH_FUNCTION_DIRECTORY,
        False,
        searchFunctionDirBase,
        validatorFunction=multipleStringValidator,
        startFunction=searchFunctionDirStart
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
        validatorFunction=multiplePairReplaceValidator,
        argBoundary=">",
        postFunction=fixfolderModifyPost
    ),

    REPLACE_CHARACTER_FILE: Procedure(
        REPLACE_CHARACTER_FILE,
        True,
        searchAndReplaceFileBase,
        searchAndReplaceFileModify,
        startFunction=searchAndReplaceFileStart,
        validatorFunction=multiplePairReplaceValidator,
        argBoundary=">"
    ),
}


# NOTE: The functions that require an argument must still be in the order as set in FIX_PROCEDURES_DISPLAY below
# This is so the "/" separated list of arguments from the "Parameter" field are properly distributed.
FIND_PROCEDURES_DISPLAY = [
    LIST_ALL,
    LIST_ALL_OWNER,
    SEARCH_FUNCTION_FILE,
    SEARCH_FUNCTION_DIRECTORY,
    IDENTICAL_FILE,
    MULTIPLE_VERSION,
    FILE_EXTENSION_SUMMARY,
    OLD_FILE,
    EMPTY_DIRECTORY,
    EMPTY_FILE,
    SPACE_DIRECTORY_FIND,
    SPACE_FILE_FIND,
    BAD_CHARACTER_DIRECTORY,
    BAD_CHARACTER_FILE,
    CHARACTER_LIMIT,
]

FIX_PROCEDURES_DISPLAY = [
    "",
    "",
    "",
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
