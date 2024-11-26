from tkinter import filedialog, messagebox
import os
from workbookManager import WorkbookManager
import checkMethodsPCPAEC


def control(dirAbsolute: str, includeSubFolders: bool, renameFiles: bool):
    # Create fileCrawlerResults if does not exist. If it does, just pass
    try:
        os.mkdir("fileCrawlerResults")
    except:
        pass

    # TODO: Perhaps workbook name with number appended if name is already taken    
    workbookPathName = "fileCrawlerResults\\" + dirAbsolute.split("/")[-1] \
    + "FileCrawlSub" + str(includeSubFolders) + "Rename" + str(renameFiles) + ".xlsx"
    print("\nCreating " + workbookPathName + "...")

    # Errors if this file already exists and is currently opened
    # Also solved by just having more unique file names, such as including the current time of execution    
    try:
        fileHandler = open(workbookPathName, 'w')
        fileHandler.close()
    except PermissionError:
        return -1

    # Initialize objects
    wbm = WorkbookManager(workbookPathName)
    wbm.setDefault()
    checkMethodsPCPAEC.setWorkBookManager(wbm)

    # Set checkMethod function
    if (renameFiles): 
        wbm.appendCheckMethod(checkMethodsPCPAEC.renameItem)
    else: 
        wbm.appendCheckMethod(checkMethodsPCPAEC.hasSpace)
        wbm.appendCheckMethod(checkMethodsPCPAEC.overCharacterLimit)
        wbm.appendCheckMethod(checkMethodsPCPAEC.badCharacters)

    # Distinguish between the inclusion of exclusion of subfolders
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


    # close gracefully and open newly created file for the user
    wbm.close()
    print("Opening " + workbookPathName + ".")
    os.startfile(workbookPathName)

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