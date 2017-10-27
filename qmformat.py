from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

direction = {'N' : 1,'S' : 2, 'E' : 3, 'W' : 4, 'CW' : 5, 'CL' : 6}
coming_from = {'work': 1,'' }

wb = load_workbook(filename = 'results-survey458251.xlsx')
ws = wb.worksheets[0]
rows = ws.max_row
