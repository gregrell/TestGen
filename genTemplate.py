import xlsxwriter as xls
workbook=xls.Workbook('Template.xlsx')
worksheet=workbook.add_worksheet('FMEA')

expenses =(
    ['rent',100],
    ['gas',100],
)

row=0
col=0
for item, cost in expenses:
    #worksheet.write(row,col, item)
    #worksheet.write(row,col+1, cost)
    row+=1


workbook.close()

import pyexcel as pe
sheet=pe.get_sheet(file_name="Template.xlsx")
print(sheet)

