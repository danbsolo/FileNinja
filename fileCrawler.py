from tkinter import filedialog, messagebox
import os
from workBookManager import WorkbookManager



def control(dirAbsolute: str, includeSubFolders: bool, renameFiles: bool):
    # TODO: Perhaps workbook name with number appended if name is already taken
    workbookName = dirAbsolute.split("/")[-1] \
    + "FileCrawlSub" + str(includeSubFolders) + "Rename" + str(renameFiles) + ".xlsx"
    print("\nCreating " + workbookName + "...")
    
    try:
        fileHandler = open(workbookName, 'w')
        fileHandler.close()
    except PermissionError:
        return -1
    
    wbm = WorkbookManager(workbookName)

    # TODO: Perhaps create a better means of setting the check method
    if (renameFiles):
        # wbm.setCheckMethod("PC-PAEC-Rename")
        wbm.setCheckMethod(wbm.renameItemPCPAEC)
    else:
        # wbm.setCheckMethod("PC-PAEC")
        wbm.setCheckMethod(wbm.showRenamePCPAEC)

    if (includeSubFolders):
        wbm.folderCrawl(os.walk(dirAbsolute))
    else:
        # mimic os.walk()'s output but only for the current directory
        dirFolders = []
        dirFiles = []
        
        for item in os.listdir(dirAbsolute):
            if os.path.isfile(os.path.join(dirAbsolute, item)):
                dirFiles.append(item)
            else:
                dirFolders.append(item)

        wbm.folderCrawl([(dirAbsolute, dirFolders, dirFiles)])


    # open newly created file for the user
    wbm.close()
    print("Opening " + workbookName + ".")
    os.startfile(workbookName)
    return 0


def view():
    # Directory Selection
    print("Selecting a directory...")
    dirAbsolute = filedialog.askdirectory()
    if (not dirAbsolute): 
        print("Nothing selected. Terminating program.")
        return
    else:
        print("Selected directory: ", dirAbsolute)


    # Subfolder Inclusion
    print("\nInquiring regarding the inclusion of subfolders...")
    includeSubFolders = messagebox.askyesnocancel("Yes or no?", "Include sub folders?")
    if (includeSubFolders):
        print("Including subfolders.")
    elif (includeSubFolders == False):
        print("Excluding subfolders.")
    else:
        print("Cancel selected. Terminating program.")
        return
    

    # Whether to rename files
    print("\nInquiring regarding renaming of files...")
    renameFiles = messagebox.askyesnocancel("Yes or no?", "Rename files? An excel sheet will log would-be changes regardless.")
    if (renameFiles):
        print("Renaming files.")
        if (includeSubFolders and not messagebox.askyesnocancel("Are you sure?", "You have chosen to include subfolders AND rename files. This is an IRREVERSIBLE action. Are you sure?")):
            print("Cancel selected. Terminating program.")
            return
    elif (renameFiles == False):
        print("Not renaming files.")
    else:
        print("Cancel selected. Terminating program.")


    exitCode = control(dirAbsolute, includeSubFolders, renameFiles) 
    if (exitCode == -1):
        print(f"Could not open file. Close file and try again.")

        

def main():
    view()



if __name__ == "__main__":
    main()