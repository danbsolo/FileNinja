import xlsxwriter
import os

wbName = "legendTest.xlsx"
wb = xlsxwriter.Workbook(wbName)
ws = wb.add_worksheet("MainLegend")

# formats
baseFont = {"font_name": "Aptos Narrow"}
baseFontBold = {**baseFont, "bold": True}

baseFontFormat = wb.add_format({**baseFont})
baseFontBoldFormat = wb.add_format({**baseFontBold})

wrapped = {**baseFont, "text_wrap": True}
wrappedFormat = wb.add_format(wrapped)

blueHeaderFormat = wb.add_format({**baseFontBold, "bg_color": "#156082", "font_color": "white"}) # BLUE
lightRedFormat = wb.add_format({**baseFont, "bg_color": "#FFC7CE", "valign": "vcenter"})
lightRedCenterFormat = wb.add_format({**baseFontBold, "bg_color": "#FFC7CE", "align": "center", "valign": "vcenter"})
yellowFormat = wb.add_format({**baseFont, "bg_color": "#FFEB9C", "valign": "vcenter"})
yellowCenterFormat = wb.add_format({**baseFontBold, "bg_color": "#FFEB9C", "align": "center", "valign": "vcenter"})
purpleFormat = wb.add_format({**baseFont, "bg_color": "#E49EDD", "valign": "vcenter"})
purpleCenterFormat = wb.add_format({**baseFontBold, "bg_color": "#E49EDD", "align": "center", "valign": "vcenter"})
whiteFormat = wb.add_format({**baseFont, "valign": "vcenter"})
whiteCenterFormat = wb.add_format({**baseFontBold, "align": "center", "valign": "vcenter"})

whiteBackgroundFormat = wb.add_format({"bg_color": "white"})
borderFormat = wb.add_format({"border": 1})
# width
ws.set_column("A:A", 30.71)
ws.set_column("B:B", 135)

# text
##
ws.write(0, 0, "ERROR TYPE", blueHeaderFormat)
ws.write(0, 1, "RECOMMENDATION", blueHeaderFormat)

##
ws.merge_range("A2:B2", "RED HIGHLIGHTS", lightRedCenterFormat)
ws.write(2, 0, "Identical File", lightRedFormat)
ws.write(3, 0, "Old File", lightRedFormat)
ws.write(4, 0, "Empty Directory", lightRedFormat)
ws.write(5, 0, "Empty File", lightRedFormat)

ws.write(2, 1, "Recommend deletion of identical files that are in the same folder. Change to [YELLOW HIGHLIGHTS] any files that you do not want deleted.", wrappedFormat)
ws.write(3, 1, " If deemed transitory information and not information of business value (IBV), recommend deletion of files that have not been accessed for more than 1095 days / 3 years.  Change to [YELLOW HIGHLIGHTS] any files that you do not want deleted.", wrappedFormat)
ws.write(4, 1, "Recommend deleting directories that contain less than 1 file and moving the less than 1 file to the parent directory until 3 or more files have accumulated. Change to [YELLOW HIGHLIGHTS] any directory that you do not want deleted and the less than 1 file moved to the parent directory.", wrappedFormat)
ws.write(5, 1, "Recommend deleting files that are either empty, corrupt or unable to be accessed.Change to [YELLOW HIGHLIGHTS] any files that you do not want deleted.", baseFontFormat)
#for row in range(3, 5): ws.set_row(row, 30)

##
ws.merge_range("A7:B7", "YELLOW HIGHLIGHTS", yellowCenterFormat)
ws.write(7, 0, "Identical File", yellowFormat)
ws.write(8, 0, "Large File Type", yellowFormat)
ws.write(9, 0, "Old Files", yellowFormat)
ws.write(10, 0, "Empty Directory", yellowFormat)

ws.write(7, 1, "Recommend review of identical files in 3 or more folders. Change to [RED HIGHLIGHTS] for deletion.", wrappedFormat)
ws.write(8, 1, "Recommend review of large files types. If deemed transitory information and not information of business value (IBV), change to  [RED HIGHLIGHTS] for deletion.  ", baseFontFormat)
ws.write(9, 1, "Recommend review of files that have not been accessed for more than 730 days / two years but less than 1095 days / 3 years. If deemed transitory information and not information of business value (IBV), change to  [RED HIGHLIGHTS] for deletion.  ", wrappedFormat)
ws.write(10, 1, "Recommend reviewing directories that contain 2 files for directory deletion and moving the 2 files to the parent directory until 3 or more files have accumulated. Change to [RED HIGHLIGHTS] any directories that you want deleted and the 2 files moved to the parent directory.", wrappedFormat)
#for row in range(9, 11): ws.set_row(row, 30)

##
ws.merge_range("A12:B12", "PURPLE HIGHLIGHTS", purpleCenterFormat)
ws.write(12, 0, "Identical File", purpleFormat)
ws.write(12, 1, "Recommend reviewing the series of identical files that exist in a minimum of three distinct folders. Change to [RED HIGHLIGHTS] for deletion.", wrappedFormat)

##
ws.merge_range("A14:B14", "NO HIGHLIGHTS", whiteCenterFormat)
ws.write(14, 0, "Identical File", whiteFormat)
ws.write(15, 0, "Replace Space with Hyphen", whiteFormat)
ws.write(16, 0, "Replace Bad Character", whiteFormat)
ws.write(17, 0, "Exceed Character Count", whiteFormat)

ws.write(14, 1, "Recommend review of identical files. Change to [RED HIGHLIGHTS] for deletion.", baseFontFormat)
ws.write(15, 1, "Recommend replacing all spaces in directory and files names with hyphens (-) to align with PAEC file naming convention and IM best practices. Change to [YELLOW HIGHLIGHTS] any directory or file that you do not want modified.", wrappedFormat)
ws.write(16, 1, "Recommend replacing all bad characters in directory and files names with hyphens (-) or other approproriate character to align with PAEC file naming convention and IM best practices. Change to [YELLOW HIGHLIGHTS] any directory or file that you do not want modified.", wrappedFormat)
ws.write(17, 1, "Change to [YELLOW HIGHLIGHTS] priority files that you want to make sure are being backed up on the shared drive. Recommend directory and/or file names will be provided for your review before changes are made. ", wrappedFormat)

##
ws.write(18, 0, "", whiteBackgroundFormat)
ws.write(18, 1, "", whiteBackgroundFormat)
ws.write(19, 1, "IMPORTANT - For Identical File, Large Files Type, Old File, Empty Directory, and Empty File; all files or folders that retain [RED HIGHIGHTS] will be modified, as per appropriate fix procedure.",
        wb.add_format({**wrapped, "bold": True, "top": 1, "left": 1, "right": 1}))
ws.write(20, 1, "IMPORTANT - For Replace Space with Hyphen, Bad Character, and Exceed Character Count; changing folder and file names will break existing links and shortcuts to the files or folders. Links and shortcuts will have to be updated.",
        wb.add_format({**wrapped, "bold": True, "left": 1, "right": 1}))
ws.write(21, 1, "IMPORTANT - For Exceed Character Count; files where the file path exceeds 200 characters are not being backed up on the shared drive.", 
        wb.add_format({**wrapped, "bold": True, "bottom": 1, "left": 1, "right": 1}))

#
wb.close()
os.startfile(wbName)