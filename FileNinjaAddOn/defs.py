from FileNinjaSuite.FileNinjaAddOn.procedureFunctions import *
from FileNinjaSuite.FileNinjaAddOn.procedureClass import *
from FileNinjaSuite.Shared.sharedDefs import *


TITLE = "File-Ninja-AddOn"
RESULTS_DIRECTORY = "File-Ninja-AddOn-Results"
NAME_CHOP_AI_QUERY = "Name-Chop AI Query"
NAME_CHOP_MODIFIER = "Name-Chop Modifier"
FILE_SHRED_FLAT = "File-Shred (FLAT)"
FILE_SHRED_TREE = "File-Shred (TREE)"
TYPO_FIXER = "Typo Fixer AI Query"


AI_PROCEDURES = {
    NAME_CHOP_AI_QUERY: Procedure(
        NAME_CHOP_AI_QUERY, 
        nameChopAIQueryFunction, 
        "aiQueryPrompt.txt"),
    
    NAME_CHOP_MODIFIER: Procedure(
        NAME_CHOP_MODIFIER, 
        nameChopModifierFunction),

    FILE_SHRED_FLAT: Procedure(
        FILE_SHRED_FLAT,
        fileShredFlatFunction),

    FILE_SHRED_TREE: Procedure(
        FILE_SHRED_TREE,
        fileShredTreeFunction),
    
    TYPO_FIXER: Procedure(
        TYPO_FIXER, 
        typoFixerFunction, 
        "typoFixerPrompt.txt")
}

PROCEDURES_DISPLAY = {
    NAME_CHOP_AI_QUERY: ("B", "E"),
    NAME_CHOP_MODIFIER: ("B", "E"),
    FILE_SHRED_FLAT: ("C", "B"),
    FILE_SHRED_TREE: ("B", ""),
    TYPO_FIXER: ("A", "B"),
}
