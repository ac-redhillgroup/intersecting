from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

direction = {'N' : 1,'S' : 2, 'E' : 3, 'W' : 4, 'CW' : 5, 'CL' : 6, 'IN' : 7, 'OB' : 8}
#loading worksheet
wb = load_workbook(filename = 'westcat.xlsx')
ws = wb.worksheets[0]
#new worksheet
nwb = Workbook()
nws = nwb.worksheets[0]
nws.title = "Westcat QM data"