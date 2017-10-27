from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

#direction = {'N' : 1,'S' : 2, 'E' : 3, 'W' : 4, 'CW' : 5, 'CL' : 6, 'IN' : 7, 'OB' : 8}
#loading worksheet
def export():
	wb = load_workbook(filename = 'westcat.xlsx')
	ws = wb.worksheets[0] 
	rows = ws.max_row
	cols = ws.max_column
	arr = []
	i = 1
	for row in ws.rows:
		r = []
		for cell in row:
			r.append(cell.value)
		arr.append(r)
	    # i = i+1
	    # if i == 4:
	    # 	break
	print arr[5]

export()
#new worksheet
# nwb = Workbook()
# nws = nwb.worksheets[0]
# nws.title = "Westcat QM data"
