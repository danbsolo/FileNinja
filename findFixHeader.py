class FindMethod:
    def __init__(self, name, mainMethod, isStateless, postMethod=None):
        self.name = name
        self.mainMethod = mainMethod
        self.isStateless = isStateless
        self.postMethod = postMethod

# Callable[[str, str], bool]
class FixMethod:
    def __init__(self, name, logMethod, modifyMethod, columnName, isFileFix, validatorFunction=None, argBoundary=None):
        self.name = name
        self.logMethod = logMethod
        self.modifyMethod = modifyMethod
        self.columnName = columnName
        self.isFileFix = isFileFix
        self.validatorFunction = validatorFunction
        self.argBoundary = argBoundary


def minimumIntegerValidator(arg:int, minimum:int):
    try:
        arg.strip()
        arg = int(arg)
        if (arg >= minimum): return arg
    except:
        return