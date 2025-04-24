class ExcelWritePackage:
    def __init__(self, row, col, text, worksheet): # TODO: Add "format" option.
        self.row = row
        self.col = col
        self.text = text
        self.worksheet = worksheet

    def executeWrite(self):
        self.worksheet.write(self.row, self.col, self.text)