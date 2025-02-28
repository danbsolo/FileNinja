class FindProcedure:
    def __init__(self, name, mainFunction, isConcurrentOnly=True, postFunction=None, isFileFind=True):
        self.name = name
        self.mainFunction = mainFunction
        self.isFileFind = isFileFind # If False, isFolderFind
        self.isConcurrentOnly = isConcurrentOnly
        self.postFunction = postFunction

class FixProcedure:
    def __init__(self, name, logFunction, modifyFunction, isFileFix, validatorFunction=None, argBoundary=None, columnName="Modification", postFunction=None):
        self.name = name
        self.logFunction = logFunction
        self.modifyFunction = modifyFunction
        self.columnName = columnName
        self.isFileFix = isFileFix  # If False, isFolderFix
        self.validatorFunction = validatorFunction
        self.argBoundary = argBoundary
        self.postFunction = postFunction


def minimumIntToInfinityOrMaxValidator(arg:str, minimum:int):
    try:
        arg = arg.split("-")

        if len(arg) != 2:  # either too few or too many hyphens (bounds)
            return minimumIntToInfinityValidator(arg[0], minimum)

        lowerBound = int(arg[0].strip())
        upperBound = int(arg[1].strip())

        if (minimum <= lowerBound) and (lowerBound <= upperBound):
            return (lowerBound, upperBound)
    
    except:
        return

def minimumIntToInfinityValidator(arg:str, minimum:int):
    try:
        arg.strip()
        arg = int(arg)
        if (arg >= minimum): return (arg,)
    except:
        return
    
def twoStringsValidator(arg:str, separator:str):
    try:
        separatorIndex = arg.rfind(separator)
        toBeReplaced = arg[0:separatorIndex].strip()
        replacer = arg[separatorIndex+len(separator):].strip()
        
        # toBeReplaced will be stripped (via strip()) to an empty string if it was just whitespace
        if (toBeReplaced) and (toBeReplaced != replacer): return (toBeReplaced, replacer)
    except:
        return
