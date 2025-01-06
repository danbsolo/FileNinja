class FindMethod:
    def __init__(self, name, mainFunction, isStateless, postFunction=None):
        self.name = name
        self.mainFunction = mainFunction
        self.isStateless = isStateless
        self.postFunction = postFunction

class FixMethod:
    def __init__(self, name, logFunction, modifyFunction, isFileFix, validatorFunction=None, argBoundary=None, columnName="Modification"):
        self.name = name
        self.logFunction = logFunction
        self.modifyFunction = modifyFunction
        self.columnName = columnName
        self.isFileFix = isFileFix  # If False, isFolderFix
        self.validatorFunction = validatorFunction
        self.argBoundary = argBoundary


def minimumIntegerValidator(arg:str, minimum:int):
    try:
        arg.strip()
        arg = int(arg)
        if (arg >= minimum): return arg
    except:
        return
    
def twoStringsValidator(arg:str, separator:str):
    try:
        separatorIndex = arg.index(separator)
        toBeReplaced = arg[0:separatorIndex].strip()
        replacer = arg[separatorIndex+len(separator):].strip()
        
        if toBeReplaced != replacer: return (toBeReplaced, replacer)
    except:
        return
