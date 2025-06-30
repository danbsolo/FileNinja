import view
import control
import common
import sys
from concurrent.futures import ThreadPoolExecutor
import time
import filesScannedSharedVar
import ctypes



def hideConsole():
    hwnd = ctypes.windll.kernel32.GetConsoleWindow()
    if hwnd:
        ctypes.windll.user32.ShowWindow(hwnd, 0)  # 0 = SW_HIDE



def main():
    # NOTE: This line is manually changed between True and False for now, hence the seemingly redundant code ahead.
    isAdmin = True

    # If Lite, always launch view.
    # If Admin and no arguments are specified, launch the view.
    # Otherwise, attempt to launch the controller with settings specified.
    if not isAdmin or len(sys.argv) <= 1:
        hideConsole()
        view.launchView(isAdmin)
        return
    
    print("Running File-Ninja-Control...")
    filePath = sys.argv[1]
    print(f"Using {filePath}...")

    with ThreadPoolExecutor(max_workers=1) as tpe:
        future = tpe.submit(control.launchControllerFromJSON, filePath)

        # run until done
        time.sleep(0.1)  # in the case it errors immediately, this wait allows that error to be caught immediately
        lastFilesScanned = -1
        while not future.done():
            time.sleep(0.5)
            if lastFilesScanned != filesScannedSharedVar.FILES_SCANNED:
                lastFilesScanned = filesScannedSharedVar.FILES_SCANNED
                print(lastFilesScanned)

        exitPair = future.result()

    print(common.interpretError(exitPair))



if __name__ == "__main__":
    main()
