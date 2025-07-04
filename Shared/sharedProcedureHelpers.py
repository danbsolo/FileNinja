import os


def addLongPathPrefix(dirAbsolute):
    if dirAbsolute.startswith('\\\\'):
        return '\\\\?\\UNC' + dirAbsolute[1:]
    else:
        return '\\\\?\\' + dirAbsolute


def joinDirToFileName(dirAbsolute, fileName):
    return dirAbsolute + "\\" + fileName


def getRootNameAndExtension(itemName):
    rootName, extension = os.path.splitext(itemName)
    return (rootName, extension.lower())


def getDirectoryBaseName(dirAbsolute):
    return os.path.basename(dirAbsolute)


def getDirectoryDirName(dirAbsolute):
    return os.path.dirname(dirAbsolute)
