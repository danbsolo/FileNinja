class FindProcedure:
    def __init__(self, name, mainFunction, isConcurrentOnly=True, postFunction=None, isFileFind=True, startFunction=None, recommendLogFunction=None, recommendPostFunction=None):
        self.name = name
        self.mainFunction = mainFunction
        self.isFileFind = isFileFind # If False, isFolderFind
        self.isConcurrentOnly = isConcurrentOnly
        self.postFunction = postFunction
        self.startFunction = startFunction
        self.recommendLogFunction = recommendLogFunction
        self.recommendPostFunction = recommendPostFunction

    def getStartFunction(self):
        return self.startFunction

    def getMainFunction(self):
        return self.mainFunction
    
    def getPostFunction(self):
        return self.postFunction


class FixProcedure:
    def __init__(self, name, logFunction, modifyFunction, isFileFix, validatorFunction=None, argBoundary=None, columnName="Modifications", postFunction=None, startFunction=None, recommendLogFunction=None, recommendPostFunction=None):
        self.name = name
        self.logFunction = logFunction
        self.modifyFunction = modifyFunction
        self.columnName = columnName
        self.isFileFix = isFileFix  # If False, isFolderFix
        self.validatorFunction = validatorFunction
        self.argBoundary = argBoundary
        self.postFunction = postFunction
        self.startFunction = startFunction
        self.recommendLogFunction = recommendLogFunction
        self.recommendPostFunction = recommendPostFunction

    def getStartFunction(self):
        return self.startFunction
    
    def getMainFunction(self, allowModify, addRecommendations):
        if allowModify:
            return self.modifyFunction
        elif addRecommendations and self.recommendLogFunction:
            return self.recommendLogFunction
        else:
            return self.logFunction
        
    def getPostFunction(self, addRecommendations):
        if addRecommendations and self.recommendPostFunction:
            return self.recommendPostFunction
        
        if self.postFunction:
            return self.postFunction
         







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

def pairOfStringsValidator(arg:str, separator:str):
    try:
        separatorIndex = arg.rfind(separator)
        toBeReplaced = arg[0:separatorIndex].strip()
        replacer = arg[separatorIndex+len(separator):].strip()
        
        # toBeReplaced will be stripped (via strip()) to an empty string if it was just whitespace
        if (toBeReplaced) and (toBeReplaced != replacer): return (toBeReplaced, replacer)
    except:
        return

def multiplePairsOfStringsValidator(arg:str, separator:str):
    multiplePairs = tuple()

    try:
        multiplePairs = arg.split("*")

        for i in range(len(multiplePairs)):
            multiplePairs[i] = pairOfStringsValidator(multiplePairs[i], separator)
        
        return multiplePairs
    except:
        return