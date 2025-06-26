class ExcelWritePackage:
    def __init__(self, row, col, text, worksheet, format=None):
        self.row = row
        self.col = col
        self.text = text
        self.worksheet = worksheet
        self.format = format

    def executeWrite(self):
        if self.format:
            self.worksheet.write(self.row, self.col, self.text, self.format)
        else:
            self.worksheet.write(self.row, self.col, self.text)
