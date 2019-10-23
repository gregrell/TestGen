
import pyexcel as pe

class xlsxWrapper():
    """
    XLS Wrapper
    Author: Rell
    Date: Oct 22 2019
    """
    def __init__(self, file):
        self.book=readFile(file)
        self.sheet_names=self.book.sheet_names()
        self.components=self.getSheet('COMPONENTS')
        self.fmea=self.getSheet('FMEA')




    def getSheet(self, sheet_name):
        try:
            return self.book.sheet_by_name(sheet_name)
        except:
            print("could not read sheet")

    def getCols(self, sheet, row):
        "returns all columns for a given row as list"
        return sheet[0]


def readFile(file):
    "reads an excel file, returns book type from pyexcel"
    try:
        book = pe.get_book(file_name=file)

        return book
    except:
        print('could not read from ',file)



