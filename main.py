# Writing to an excel
# sheet using Python
import xlwt
import xlrd
from xlwt import Workbook

# Workbook is created
wb = Workbook()

# add_sheet is used to create sheet.
sheet1 = wb.add_sheet('Sheet 1')


row = 0
column = 0

# sheet1.write(1, 0, 'ISBT DEHRADUN')
# sheet1.write(2, 0, 'SHASTRADHARA')
# sheet1.write(3, 0, 'CLEMEN TOWN')
# sheet1.write(4, 0, 'RAJPUR ROAD')
# sheet1.write(5, 0, 'CLOCK TOWER')
# sheet1.write(0, 1, 'ISBT DEHRADUN')
# sheet1.write(0, 2, 'SHASTRADHARA')
# sheet1.write(0, 3, 'CLEMEN TOWN')/
# sheet1.write(0, 4, 'RAJPUR ROAD')
# sheet1.write(0, 5, 'CLOCK TOWER')

#wb.save('xlwt example.xls')
location = "excelsheet_example\venv\Scripts"
workbook = xlrd.open_workbook(location)
sheet = workbook.sheet_by_index(0)
print(sheet.cell_value(0, 0))





#C:\Users\Sohum\Documents\PythonProject\excelsheet_example\venv\Scripts