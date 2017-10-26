from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
wb = load_workbook(filename = 'test.xlsx')
ws = wb.worksheets[0]
rows = ws.max_row
#finding the cell number
cols = ws.max_column
for i in range(1,cols+1):
	try:
		if ws.cell(row=1,column=i).value == "Surveyor Comments:":
			col = get_column_letter(i)
	except ValueError:
		print "cannot code"
#update cells
for i in range(1,rows):
	if ws[col][i].value:
		ws[col][i].value = ws['QC'][i].value.replace('\n',' ')
		print ws[col][i].value

wb.save('test.xlsx')