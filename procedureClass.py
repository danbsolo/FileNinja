class Procedure:
    def __init__(self, name, isFileProcedure, baseFunction, modifyFunction=None, isConcurrentOnly=True, validatorFunction=None, argBoundary=None, defaultArgument=None, startFunction=None, postFunction=None, recommendBaseFunction=None, recommendPostFunction=None):
        self.name = name
        self.isFileProcedure = isFileProcedure # If False, isFolderFind
        self.baseFunction = baseFunction
        self.modifyFunction = modifyFunction
        self.isConcurrentOnly = isConcurrentOnly
        self.startFunction = startFunction
        self.postFunction = postFunction
        self.recommendBaseFunction = recommendBaseFunction
        self.recommendPostFunction = recommendPostFunction
        self.validatorFunction = validatorFunction
        self.argBoundary = argBoundary
        self.defaultArgument = defaultArgument
        self.lastValidatedArgument = None

    def getPostFunction(self, addRecommendations):
        if addRecommendations and self.recommendPostFunction:
            return self.recommendPostFunction
        if self.postFunction:
            return self.postFunction

    def getValidArgument(self, arg):
        if not self.validatorFunction:
            return ()
        
        if (potentialArg := self.validatorFunction(arg, self.argBoundary)):
            return potentialArg
        else:
            return self.defaultArgument
    
    def getMainFunction(self, allowModify, addRecommendations):
        if allowModify and self.modifyFunction:
            return self.modifyFunction
        elif addRecommendations and self.recommendBaseFunction:
            return self.recommendBaseFunction
        else:
            return self.baseFunction

    def isFixFunction(self):
        # As opposed to a "Find" function
        return bool(self.modifyFunction)

    def getIsConcurrentOnly(self):
        return self.isConcurrentOnly

    def getStartFunction(self):
        return self.startFunction
    
    def getDefaultArgument(self):
        return self.defaultArgument
    
    def getIsFileProcedure(self):
        return self.isFileProcedure



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
        arg = int(arg)
        if (arg >= minimum): return (arg,)
    except:
        return


def pairOfStringsValidator(arg:str, separator:str):
    try:
        separatorIndex = arg.rfind(separator)
        
        if (separatorIndex == -1):
            return

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
            pair = pairOfStringsValidator(multiplePairs[i], separator)

            if pair is None:
                return
            
            multiplePairs[i] = pair

        return multiplePairs
    except:
        return