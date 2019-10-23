
from xlsxWrapper import xlsxWrapper
from xlsxWrapper import readFile as rf

if __name__=="__main__":
    """
    Test Generator, parser, pdf creator
    Author: Rell
    Date: Oct 22 2019
    """

    template = xlsxWrapper("Template.xlsx")
    fmea_cols=template.getCols(template.fmea,0)
    print(fmea_cols)


