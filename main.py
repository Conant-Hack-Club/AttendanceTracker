# Writing to an excel
# sheet using Python
import xlwt
import xlrd
from xlwt import Workbook\


location = "blah.xls"

workbook = xlrd.open_workbook(location)
sheet = workbook.sheet_by_index(0)
print(sheet.cell_value(0, 0))