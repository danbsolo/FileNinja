import tkinter as tk
from tkinter import filedialog
import os
from workbookManager import WorkbookManager
from datetime import datetime
from defs import *
from sys import argv


def control(dirAbsolute:str, includeSubfolders:bool, modify:bool, selectedFindProcedures:list[str], selectedFixProcedure:str, unprocessedArg:str):
    if (not dirAbsolute): return -2
    
    # Create fileCrawlerResults directory name if does not exist
    try: os.mkdir(RESULTS_DIRECTORY)
    except: pass

    # >>> "folderName-20##-##-##.xlsx"
    workbookPathName = RESULTS_DIRECTORY + "\\" + dirAbsolute.split("/")[-1] + "-" \
    + datetime.today().strftime("%Y-%m-%d") + ".xlsx"

    # Initialize objects
    wbm = WorkbookManager(workbookPathName)
    setWorkbookManager(wbm)

    # Errors if this file already exists and is currently opened
    try:
        fileHandler = open(workbookPathName, 'w')
        fileHandler.close()
    except PermissionError:
        return -1

    # Set findProcedures and fixProcedure
    for fm in selectedFindProcedures:
        wbm.addFindProcedure(FIND_PROCEDURES[fm])

    if selectedFixProcedure != NULL_OPTION:
        fixProcedureObject = FIX_PROCEDURES[selectedFixProcedure]

        if not wbm.setFixArg(fixProcedureObject, unprocessedArg):
            return -3
        
        wbm.setFixProcedure(fixProcedureObject, modify)
                    
    wbm.styleSummarySheet(dirAbsolute, includeSubfolders, modify)    

    # print("\nCreating " + workbookPathName + "...")
    if (includeSubfolders):
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
    # print("Opening " + workbookPathName + ".")
    os.startfile(workbookPathName)
    return 0



def view(isAdmin: bool):
    def launchController():
        if modifyState.get() and not tk.messagebox.askyesnocancel("Allow Modify?", "You have chosen to modify items. This is an IRREVERSIBLE action. Are you sure?"):
            # print("Aborted. Continuing selection.")
            return

        root.title(FILE_CRAWLER + ": CURRENTLY RUNNING...")
        exitCode = control(dirAbsoluteVar.get(), bool(includeSubFoldersState.get()), bool(modifyState.get()), 
                           [findListbox.get(fm) for fm in findListbox.curselection()],
                           fixListbox.get(tk.ACTIVE),
                           parameterVar.get())
        root.title(FILE_CRAWLER)
        
        errorMessage = ""
        if (exitCode == 0):
            return
        elif (exitCode == -1):
            errorMessage = "Could not open file. Close file and try again."
        elif (exitCode == -2):
            errorMessage = "Invalid directory."
        elif (exitCode == -3):
            errorMessage = "Invalid argument."
        else:
            errorMessage = "An error has occurred."

        tk.messagebox.showerror("Error: " + str(exitCode), errorMessage)


    def selectDirectory():
        potentialDirectory = filedialog.askdirectory()

        if (potentialDirectory):
            dirAbsoluteVar.set(potentialDirectory)

        
    def openResultsDirectory():
        if os.path.exists(RESULTS_DIRECTORY):
            os.startfile(RESULTS_DIRECTORY)
        else:
            tk.messagebox.showinfo("Directory DNE", "Directory \"" + RESULTS_DIRECTORY + "\" does not exist. Try executing a file crawl first.")


    def checkPassword():
        if passwordEntryVar.get() == PASSWORD:
            addAdminLayout(includeSubfoldersCheckbutton.winfo_height())
            root.unbind("<Control-Key-p>")
        else:
            tk.messagebox.showerror("INCORRECT", "Incorrect password.")
        

    def inquirePassword(event):
        passwordWindow = tk.Toplevel(root)
        passwordWindow.title("Type password")
        passwordWindow.focus_set()

        # passwordWindow.lift(root)  # -> Put over root

        passwordEntry = tk.Entry(passwordWindow, show="*", textvariable=passwordEntryVar, width=30, font=fontGeneral)
        passwordEntry.pack()
        passwordEntry.focus_set()

        passwordEntry.bind("<Return>", lambda e: [
            checkPassword(), passwordEntryVar.set(""), passwordWindow.destroy()])
    

    def addAdminLayout(extraHeight):
        root.geometry("{}x{}".format(rootWidth, rootHeight + extraHeight))
        root.attributes('-topmost', True)  # keeps root window at top layer

        modifyCheckbutton = tk.Checkbutton(frames[5], text="Allow Modify", variable=modifyState, font=fontGeneral)
        modifyCheckbutton.pack(padx=(0, 50))
        

    # root stuff
    root = tk.Tk()
    root.title(FILE_CRAWLER)
    root.resizable(0, 0)
    rootWidth = 500
    rootHeight = 390
    root.geometry("{}x{}".format(rootWidth, rootHeight))
        
    frames = []
    for i in range(8):
        frames.append(tk.Frame(root, bd=0, relief=tk.SOLID))
        frames[i].pack(fill="x", padx=10, pady=3)


    # aesthetic/layout variables
    fontType = "None"
    fontSize = 15
    fontGeneral = (fontType, fontSize)
    fontSmall = (fontType, int(fontSize/3*2))
    listboxHeight = max(len(FIND_PROCEDURES.keys()), len(FIX_PROCEDURES.keys())) +1
    listboxWidth = int(rootWidth/15)
    finalButtonsWidth = 20

    # data variables
    dirAbsoluteVar = tk.StringVar()
    parameterVar = tk.StringVar()
    includeSubFoldersState = tk.IntVar()
    modifyState = tk.IntVar()


    # widgets
    browseButton = tk.Button(frames[0], text="Browse to Select", command=selectDirectory, font=fontGeneral, width=rootWidth)
    browseButton.pack()

    dirHeaderLabel = tk.Label(frames[1], text = "Directory:", font=fontGeneral)
    dirLabel = tk.Label(frames[1], textvariable=dirAbsoluteVar, font=fontSmall, anchor="e") 
    dirHeaderLabel.pack(side=tk.LEFT)
    dirLabel.pack(side=tk.LEFT)
    
    findLabel = tk.Label(frames[2], text="Find", font=fontGeneral)
    fixLabel = tk.Label(frames[2], text="Fix", font=fontGeneral)
    fixLabel.pack(side=tk.RIGHT, padx=(0, rootWidth/5))
    findLabel.pack(side=tk.LEFT, padx=(rootWidth/5, 0))

    frames[2].pack(fill="x", padx=10, pady=(3, 0))  # inadvertently packed twice to have less y padding

    findListbox = tk.Listbox(frames[3], selectmode="multiple", exportselection=0, width=listboxWidth, height=listboxHeight)
    fixListbox = tk.Listbox(frames[3], exportselection=0, width=listboxWidth, height=listboxHeight)
    for findProcedureName in FIND_PROCEDURES.keys():
        findListbox.insert(tk.END, findProcedureName)
    for fixProcedureName in FIX_PROCEDURES.keys():
        fixListbox.insert(tk.END, fixProcedureName)
    findListbox.select_set(0)
    fixListbox.select_set(0)
    findListbox.config(font=fontSmall)
    fixListbox.config(font=fontSmall)
    findListbox.pack(side=tk.LEFT)
    fixListbox.pack(side=tk.RIGHT)
    frames[3].pack(fill="x", padx=10, pady=(0, 3))

    parameterLabel = tk.Label(frames[4], text="Parameter:", font=fontGeneral)
    argumentEntry = tk.Entry(frames[4], textvariable=parameterVar, width=rootWidth, font=fontSmall)
    parameterLabel.pack(side=tk.LEFT)
    argumentEntry.pack(side=tk.LEFT)

    includeSubfoldersCheckbutton = tk.Checkbutton(frames[5], text="Include Subfolders", variable=includeSubFoldersState, font=fontGeneral)
    includeSubfoldersCheckbutton.pack()

    executeButton = tk.Button(frames[6], text="Execute", command=launchController, width=finalButtonsWidth, font=fontGeneral)
    executeButton.pack()
    frames[6].configure(width=rootWidth/2)
    frames[6].pack(side=tk.LEFT, expand=True)

    resultsButton = tk.Button(frames[7], text="Results", command=openResultsDirectory, width=finalButtonsWidth, font=fontGeneral)
    resultsButton.pack()
    frames[7].configure(width=rootWidth/2)
    frames[7].pack(side=tk.LEFT, expand=True)

    root.bind('<Control-Key-w>', lambda e: root.destroy())
    root.bind('<Control-Key-W>', lambda e: root.destroy())

    if isAdmin:
        # root.after(50, addAdminLayout)
        addAdminLayout(35)
    else:
        passwordEntryVar = tk.StringVar()
        root.bind('<Control-Key-p>', inquirePassword)
        root.bind('<Control-Key-P>', inquirePassword)

    root.mainloop()
        

def main():
    if (len(argv) <= 1):
        view(False)
    elif (argv[1] == PASSWORD):
        view(True)
    else:
        tk.messagebox.showerror("FileCrawler: INCORRECT", "Incorrect.")


if __name__ == "__main__":
    main()
