from FileNinjaSuite.FileNinja.defs import *
from FileNinjaSuite.Shared.sharedCommon import *
import json

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
