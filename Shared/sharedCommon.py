from FileNinjaSuite.Shared.sharedDefs import *

def interpretError(exitCode, exitMessage):    
    if exitMessage:
        return f"{EXIT_STATUS_CODES[exitCode]} >>> {exitMessage}"
    else:
        return EXIT_STATUS_CODES[exitCode]