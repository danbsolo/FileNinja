HCS_ASSETS_PATH = "\\\\BNO-files\\NP-GROUPS\\PAEC-ECAP\\PAE-EAP\\Horizontal-Coordination-Support\\Admin\\HCS-Tools-Procedures\\File-Ninja\\Assets\\"
LOGO_PATH = HCS_ASSETS_PATH + "File-Ninja-Logo-Square.png"

STATUS_SUCCESSFUL = 0
STATUS_IDLE = 1
STATUS_RUNNING = 102  # 102 == The HTTP response code for "Still processing"
STATUS_UNEXPECTED = -999

EXIT_STATUS_CODES = {
    STATUS_SUCCESSFUL: "Successful.",
    STATUS_IDLE: "Idle.",
    STATUS_RUNNING: "*Should* still be processing.",
    -1: "ERROR: Could not open file. Close file and try again.",
    -2: "ERROR: Invalid directory.",
    -3: "ERROR: Invalid argument.",
    -4: "ERROR: Invalid excluded directory.",
    -5: "ERROR: Invalid settings. Cannot run multiple Fix Procedures when modify is checked.",
    -6: "ERROR: Invalid arguments. Separate with \"/\"",
    -7: "ERROR: Invalid settings. Cannot run modifications simultaneously with recommendations and/or hidden files.",
    -8: "ERROR: No procedures selected.",
    -9: "ERROR: Directory does not exist.",
    -10: "ERROR: JSON Decode error.",
    -11: "ERROR: Missing keys in JSON.",
    -12: "ERROR: File path does not exist.",
    -13: "ERROR: Invalid excluded extension.",
    -101: "ERROR: Prompt does not exist.",
    -102: "ERROR: Some chunks may be faulty.",
    STATUS_UNEXPECTED: "ERROR: Unexpected condition occurred."
}
