
from xlsxWrapper import xlsxWrapper

if __name__=="__main__":
    """
    Test Generator, parser, pdf creator
    """

    template = xlsxWrapper("Template.xlsx")
    print(template.main_sheet)

