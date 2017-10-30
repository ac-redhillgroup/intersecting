import os
from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter


#loading worksheet
def export():
	wb = load_workbook(filename = 'westcat.xlsx')
	ws = wb.worksheets[0] 
	direction = {'N' : 1,'S' : 2, 'E' : 3, 'W' : 4, 'CW' : 5, 'CL' : 6, 'IN' : 7, 'OB' : 8}
	ags = {10 : 1, 11 : 2, 12 : 3, 15 : 4, 16 : 5, 17 : 6, 18 : 7, 19 : 8, "30Z" : 9, "C3" : 10, "JR" : 11, "JL" : 12, "JX" : 13, "JPX" : 14, "LYNX" : 15, "-oth-" : 16}
	tags = {}
	# new worksheet
	nwb = Workbook()
	nws = nwb.worksheets[0]
	nws.title = "Westcat QM data"
	rows = ws.max_row
	cols = ws.max_column
	arr = []
	i = 1
	for row in ws.rows:
		r = []
		for cell in row:
			r.append(cell.value)
		#print r[27]
		arr.append(r)
	    # i = i+1
	    # if i == 4:
	    # 	break
	dic = {}
	headers = arr[1]
	for header in range(0,len(headers)):
		dic[headers[header]] = header
	val = arr[6]
	crow = 1
	for idx,v in enumerate(val):
		if idx == dic["id"]:
			nws.cell(row=1,column= 1).value = v
		elif idx == dic["InterviewersInitials"]:
			nws.cell(row=1,column= 2).value = v
		elif idx == dic["g1xRoutexSWCx0"]:
			nws.cell(row=1,column= 3).value = ags[v]
		elif idx == dic["g1xRoutexSWCx0[other]"]:
			nws.cell(row=1,column= 4).value = v
		elif idx == dic["xTimeofDayx0"]:
			nws.cell(row=1,column= 5).value = v
		elif idx == dic["xDirectionx0"]:
			nws.cell(row=1,column= 6).value = direction[v]
		elif idx == dic["NumberOfRiders"]:
			nws.cell(row=1,column= 7).value = v
		elif idx == dic["startLocOfDevice"]:
			#check if location is there or not
			if v:
				loc = [x.strip() for x in v.split(',')]
				nws.cell(row=1,column= 8).value = loc[0]
				nws.cell(row=1,column= 9).value = loc[1]
		#origin related columns
		elif idx == dic["g1xOriginx0"]:
			nws.cell(row=1,column= 10).value = 14 if (v == "-oth-") else v
		elif idx == dic["g1xOriginx0[other]"]:
			nws.cell(row=1,column= 11).value = v
		elif idx == dic["g1xMapx10[1]"]:
			#check if location is there or not
			if v:
				loc = [x.strip() for x in v.split(',')]
				nws.cell(row=1,column= 12).value = loc[0]
				nws.cell(row=1,column= 13).value = loc[1]
		elif idx == dic["g1xMapx10[2]"]:
			nws.cell(row=1,column= 14).value = v
		#destination related columns
		elif idx == dic["g1xDestix0"]:
			nws.cell(row=1,column= 15).value = 14 if (v == "-oth-") else v
		elif idx == dic["g1xDestix0[other]"]:
			nws.cell(row=1,column= 16).value = v
		elif idx == dic["g1xMapx10[6]"]:
			#check if location is there or not
			if v:
				loc = [x.strip() for x in v.split(',')]
				nws.cell(row=1,column= 17).value = loc[0]
				nws.cell(row=1,column= 18).value = loc[1]
		elif idx == dic["g1xMapx10[7]"]:
			nws.cell(row=1,column= 19).value = v
		elif idx == dic["AccessMode"]:
			#20 column is for roundtrip
			nws.cell(row=1,column= 21).value = 9 if (v == "-oth-") else v
		elif idx == dic["AccessMode[other]"]:
			nws.cell(row=1,column= 22).value = v
		elif idx == dic["AccessMinutes"]:
			nws.cell(row=1,column= 23).value = v
		elif idx == dic["AccessMiles"]:
			nws.cell(row=1,column= 24).value = v
		elif idx == dic["FirstB4Trans"]:
			#25 column is for calculating transfers
			nws.cell(row=1,column= 26).value = v[1:]
		elif idx == dic["g1xAgencyx1"]:
			nws.cell(row=1,column= 27).value = v

		nwb.save("qm.xlsx")
	os.system("start " + "qm.xlsx")
		
export()