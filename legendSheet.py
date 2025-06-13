import xlsxwriter
import os

wbName = "legendTest.xlsx"
wb = xlsxwriter.Workbook(wbName)
ws = wb.add_worksheet("Legend")

# format maps
baseFontMap = {"font_name": "Aptos Narrow"}
baseFontBoldMap = {**baseFontMap, "bold": True}
wrappedMap = {**baseFontMap, "text_wrap": True}

blueHeaderMap = {**baseFontBoldMap, "bg_color": "#156082", "font_color": "white"}
redMap = {**baseFontMap, "bg_color": "#FFC7CE", "valign": "vcenter"}
redCenterMap = {**redMap, "bold": True, "align": "center"}
yellowMap = {**baseFontMap, "bg_color": "#FFEB9C", "valign": "vcenter"}
yellowCenterMap = {**yellowMap, "bold": True, "align": "center"}
purpleMap = {**baseFontMap, "bg_color": "#E49EDD", "valign": "vcenter"}
purpleCenterMap = {**purpleMap, "bold": True, "align": "center"}
noColorMap = {**baseFontMap, "valign": "vcenter"}
noColorCenterMap = {**noColorMap, "bold": True, "align": "center"}
whiteBackgroundMap = {"bg_color": "white"}

# width
ws.set_column("A:A", 30.71)
ws.set_column("B:B", 135)

# text
##
ws.write(0, 0, "ERROR TYPE", wb.add_format({**blueHeaderMap, "border": 1}))
ws.write(0, 1, "RECOMMENDATION", wb.add_format({**blueHeaderMap, "border": 1}))

##
ws.merge_range("A2:B2", "RED HIGHLIGHTS", wb.add_format({**redCenterMap, "top": 2, "bottom": 2, "right": 1}))
ws.write(2, 0, "Identical File", wb.add_format({**redMap, "border": 1}))
ws.write(3, 0, "Old File", wb.add_format({**redMap, "border": 1}))
ws.write(4, 0, "Empty Directory", wb.add_format({**redMap, "border": 1}))
ws.write(5, 0, "Empty File", wb.add_format({**redMap, "border": 1}))

# Some of these are not set to wrap, cause then it messes with the height. It's weird.
ws.write(2, 1, "Recommend deletion of identical files that are in the same folder. Change to [YELLOW HIGHLIGHTS] any files that you do not want deleted.", wb.add_format({**wrappedMap, "border": 1}))
ws.write(3, 1, " If deemed transitory information and not information of business value (IBV), recommend deletion of files that have not been accessed for more than 1095 days / 3 years.  Change to [YELLOW HIGHLIGHTS] any files that you do not want deleted.", wb.add_format({**wrappedMap, "border": 1}))
ws.write(4, 1, "Recommend deleting directories that contain less than 1 file and moving the less than 1 file to the parent directory until 3 or more files have accumulated. Change to [YELLOW HIGHLIGHTS] any directory that you do not want deleted and the less than 1 file moved to the parent directory.", wb.add_format({**wrappedMap, "border": 1}))
ws.write(5, 1, "Recommend deleting files that are either empty, corrupt or unable to be accessed. Change to [YELLOW HIGHLIGHTS] any files that you do not want deleted.", wb.add_format({**baseFontMap, "border": 1}))

##
ws.merge_range("A7:B7", "YELLOW HIGHLIGHTS", wb.add_format({**yellowCenterMap, "top": 2, "bottom": 2, "right": 1}))
ws.write(7, 0, "Identical File", wb.add_format({**yellowMap, "border": 1}))
ws.write(8, 0, "Large File Type", wb.add_format({**yellowMap, "border": 1}))
ws.write(9, 0, "Old Files", wb.add_format({**yellowMap, "border": 1}))
ws.write(10, 0, "Empty Directory", wb.add_format({**yellowMap, "border": 1}))

ws.write(7, 1, "Recommend review of identical files in 3 or more folders. Change to [RED HIGHLIGHTS] for deletion.", wb.add_format({**wrappedMap, "border": 1}))
ws.write(8, 1, "Recommend review of large files types. If deemed transitory information and not information of business value (IBV), change to  [RED HIGHLIGHTS] for deletion.  ", wb.add_format({**wrappedMap, "border": 1}))
ws.write(9, 1, "Recommend review of files that have not been accessed for more than 730 days / two years but less than 1095 days / 3 years. If deemed transitory information and not information of business value (IBV), change to  [RED HIGHLIGHTS] for deletion.  ", wb.add_format({**wrappedMap, "border": 1}))
ws.write(10, 1, "Recommend reviewing directories that contain 2 files for directory deletion and moving the 2 files to the parent directory until 3 or more files have accumulated. Change to [RED HIGHLIGHTS] any directories that you want deleted and the 2 files moved to the parent directory.", wb.add_format({**wrappedMap, "border": 1}))
#for row in range(9, 11): ws.set_row(row, 30) # for setting height, if ever needed

##
ws.merge_range("A12:B12", "PURPLE HIGHLIGHTS", wb.add_format({**purpleCenterMap, "top": 2, "bottom": 2, "right": 1}))
ws.write(12, 0, "Identical File", wb.add_format({**purpleMap, "border": 1}))
ws.write(12, 1, "Recommend reviewing the series of identical files that exist in a minimum of three distinct folders. Change to [RED HIGHLIGHTS] for deletion.", wb.add_format({**wrappedMap, "border": 1}))

##
ws.merge_range("A14:B14", "NO HIGHLIGHTS", wb.add_format({**noColorCenterMap, "top": 2, "bottom": 2, "right": 1}))
ws.write(14, 0, "Identical File", wb.add_format({**noColorMap, "border": 1}))
ws.write(15, 0, "Replace Space with Hyphen", wb.add_format({**noColorMap, "border": 1}))
ws.write(16, 0, "Replace Bad Character", wb.add_format({**noColorMap, "border": 1}))
ws.write(17, 0, "Exceed Character Count", wb.add_format({**noColorMap, "border": 1}))

ws.write(14, 1, "Recommend review of identical files. Change to [RED HIGHLIGHTS] for deletion.", wb.add_format({**baseFontMap, "border": 1}))
ws.write(15, 1, "Recommend replacing all spaces in directory and files names with hyphens (-) to align with PAEC file naming convention and IM best practices. Change to [YELLOW HIGHLIGHTS] any directory or file that you do not want modified.", wb.add_format({**wrappedMap, "border": 1}))
ws.write(16, 1, "Recommend replacing all bad characters in directory and files names with hyphens (-) or other approproriate character to align with PAEC file naming convention and IM best practices. Change to [YELLOW HIGHLIGHTS] any directory or file that you do not want modified.", wb.add_format({**wrappedMap, "border": 1}))
ws.write(17, 1, "Change to [YELLOW HIGHLIGHTS] priority files that you want to make sure are being backed up on the shared drive. Recommend directory and/or file names will be provided for your review before changes are made. ", wb.add_format({**wrappedMap, "border": 1}))

##
ws.write(18, 0, "", wb.add_format({**whiteBackgroundMap, "top": 2}))
ws.write(18, 1, "", wb.add_format({**whiteBackgroundMap, "top": 2}))
ws.write(19, 1, "IMPORTANT - For Identical File, Large Files Type, Old File, Empty Directory, and Empty File; all files or folders that retain [RED HIGHIGHTS] will be modified, as per appropriate fix procedure.", wb.add_format({**wrappedMap, "bold": True, "top": 1, "left": 1, "right": 1}))
ws.write(20, 1, "IMPORTANT - For Replace Space with Hyphen, Bad Character, and Exceed Character Count; changing folder and file names will break existing links and shortcuts to the files or folders. Links and shortcuts will have to be updated.", wb.add_format({**wrappedMap, "bold": True, "left": 1, "right": 1}))
ws.write(21, 1, "IMPORTANT - For Exceed Character Count; files where the file path exceeds 200 characters are not being backed up on the shared drive.", wb.add_format({**wrappedMap, "bold": True, "bottom": 1, "left": 1, "right": 1}))

#
wb.close()
os.startfile(wbName)