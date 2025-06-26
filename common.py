from .defs import *
import json


def interpretError(exitPair):
    exitCode, exitMessage = exitPair
    
    if exitMessage:
        return f"{EXIT_STATUS_CODES[exitCode]} >>> {exitMessage}"
    else:
        return EXIT_STATUS_CODES[exitCode]


def loadSettingsFromJSON(filePath):
    if not os.path.exists(filePath):
        return (-12, filePath)

    try:
        with open(filePath, "r") as f:
            settings = json.load(f)
    except json.JSONDecodeError as e:
        return (-10, f"{e}")
    
    missingKeys = []
    for key in JSON_KEYS:
        if key not in settings:
            missingKeys.append(key)
    if missingKeys:
        return (-11, f"{missingKeys}")
    else:
        return (STATUS_SUCCESSFUL, settings)
