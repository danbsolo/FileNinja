from tkinter import filedialog, messagebox
import xlsxwriter
import os
from typing import List, Tuple

class WorkbookManager:

    def __init__(self, workbookName: str):
        self.workbookName = workbookName
        self.wb = xlsxwriter.Workbook(workbookName)
        
        # TODO: add more worksheets here when necessary
        self.mainSheetRow = 1
        self.mainSheet = self.wb.add_worksheet("MainSheet")
        
        ## Set formatting
        # Column indexes
        self.DIR_COL = 0
        self.ITEM_COL = 1
        self.RENAME_COL = 2
        self.ERROR_COL = 3

        # Column width
        self.mainSheet.set_column(self.DIR_COL, self.DIR_COL, 50)
        self.mainSheet.set_column(self.ITEM_COL, self.RENAME_COL, 30)
        self.mainSheet.set_column(self.ERROR_COL, self.ERROR_COL+5, 50)

        # Default cell styles
        self.dirColFormat = self.wb.add_format({
            "bg_color": "#99CCFF", # blueish
            "bold": True
        })

        self.fileErrorFormat = self.wb.add_format({
            "bg_color": "#FF4444", # reddish
            "bold": True
        })

        self.headerFormat = self.wb.add_format({
            "bg_color": "#C0C0C0", # grayish
            "bold": True
        })

        # Write headers
        self.mainSheet.write(0, self.DIR_COL, "Directories", self.headerFormat)
        self.mainSheet.write(0, self.ITEM_COL, "Items", self.headerFormat)
        self.mainSheet.write(0, self.RENAME_COL, "Rename", self.headerFormat)
        self.mainSheet.write(0, self.ERROR_COL, "Errors", self.headerFormat)
        

    def fileCrawl(self, dirItems: List[str]):
        # can be used to write down folders as well
        for itemName in dirItems:
            if (self.pcNamingConventionCheck(itemName)):
                self.mainSheet.write(self.mainSheetRow, self.ITEM_COL, itemName)
            else:
                self.mainSheet.write(self.mainSheetRow, self.ITEM_COL, itemName, self.fileErrorFormat)

            self.mainSheetRow += 1
    

    def folderCrawl(self, dirTree: List[Tuple[str, list, list]]):
        for (dirAbsolute, dirFolders, dirFiles) in dirTree:
            self.mainSheet.write(self.mainSheetRow, self.DIR_COL, dirAbsolute, self.dirColFormat)
            self.fileCrawl(dirFolders + dirFiles)


    def pcNamingConventionCheck(self, itemName: str) -> bool:
        # If good, return True. If error, write down the error(s) and return False.
        return True
    

    def close(self):
        self.wb.close()



def control(dirAbsolute: str, includeSubFolders: bool, renameFiles: bool):
    # if the file already exists, overwrite it
    workbookName = dirAbsolute.split("/")[-1] + "FileCrawl" + ".xlsx"
    wbm = WorkbookManager(workbookName)

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


    wbm.close()

    # start newly created file for the user
    os.startfile(workbookName)



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
    

    # Rename files or just build excel sheet
    print("\nInquiring regarding command...")
    renameFiles = messagebox.askyesnocancel("Yes or no?", "Rename files? An excel sheet will be made regardless.")
    if (renameFiles):
        print("Renaming files.")
        if (includeSubFolders and not messagebox.askyesnocancel("Are you sure?", "You have chosen to include subfolders AND rename files. This is an IRREVERSIBLE action. Are you sure?")):
            print("Cancel selected. Terminating program.")
            return
    elif (renameFiles == False):
        print("Not renaming files.")
    else:
        print("Cancel selected. Terminating program.")


    control(dirAbsolute, includeSubFolders, renameFiles)

        

def main():
    view()



if __name__ == "__main__":
    main()