# File Ninja
File Ninja helps you manage your files and automate tasks that are too repetitive to do manually. The Admin version includes everything from the Lite version, and adds the following features:
- Fix procedures
- Parameter# field
- Toggle for *Allow Modify*
- Toggle for *Add Recommendations~*
- Toggle for *Include Hidden Files*


# Installing File Ninja
- Download and open the File Ninja p7m file.
- OR run the Python code after installing XlsxWriter (`pip install XlsxWriter`).


# Basic Elements
| UI Element | Description |
|---|---|
| Browse to Select | Browse to select the directory to perform operations on. |
| Browse to Exclude | Select subdirectories of the main directory to exclude from operation. Hidden folders — and their subdirectories — are automatically excluded. To remove a previously added excluded directory, either double-click or right click it. (If *Include Subfolders* is off, this feature does nothing.)
| Find procedures | Flags any errant files based on the selection. |
| Fix procedures | If *Allow Modify* is off, flags any errant files based on the selection, showing would-be fixes. If *Allow Modify* is on, executes said fixes. |
| Parameter# | Fix procedures that end in *#* require an argument. Input their arguments here. If running multiple Fix procedures, separate arguments using "/", in order from top to bottom of the list of Fix procedures selected. <br> For example: If running "Replace Character (DIR)#" and "Replace Character (FILE)#", you may input (without the quotes) "&>and / @>at". |
| Include Subdirectories | If on, File Ninja will peruse through all subdirectories as well. Otherwise, will only traverse the currently selected directory. Related: "Browse to Exclude". |
| Allow Modify | If on, executes all changes irreversibly. Can not run multiple modifier Fix procedures simultaneously. |
| Add Recommendations~ | If on, procedures that end in "~" will show helpful recommendations. Red highlight indicates a stronger warning -- those files should perhaps be deleted. Yellow highlight indicates a weaker warning -- those files should be looked at, at the very least.
| Execute | Commence execution. |
| Results | Open results folder, containing Excel files of previous executions. File names follow the convention "<<SelectedFolderName>>-YY-mm-DD-HH-MM-SS.xlsx". |
| Summary | A metric detailing various metrics regarding the execution. |

# PLACEHOLDER

# Hints
- Middle-click anywhere in the window to alternate between light and dark mode.
- `ctrl+w` to close the window.