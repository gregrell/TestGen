"""
XLS Wrapper
"""
import pyexcel as pe

class xlsxWrapper():
    def __init__(self, file):
        self.book=readFile(file)
        self.sheet_names=self.book.sheet_names()
        self.main_sheet=self.getSheet(self.sheet_names[1])


    def getSheet(self, sheet_name):
        try:
            return self.book.sheet_by_name(sheet_name)
        except:
            print("could not read sheet")


def readFile(file):
    "reads an excel file, returns book type from pyexcel"
    try:
        book = pe.get_book(file_name=file)

        return book
    except:
        print('could not read from ',file)



