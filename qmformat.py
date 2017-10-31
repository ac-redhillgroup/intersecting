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
	transit_agencies = ["3D","AC","EM","BA","CC","FS","GG","SF","SM","ST","VN","WC","-oth-"]
	transit_agencies_dict = {}
	for idx,i in enumerate(transit_agencies):
		transit_agencies_dict[i] = idx+1
	print transit_agencies_dict
	bart_dict = {}
	bart_wb = load_workbook(filename = 'BART_CodeDict.xlsx')
	bart_ws = bart_wb.worksheets[0]
	for row in bart_ws.rows:
		bart_dict[row[0].value] = row[1].value
	print bart_dict['16SM']
	emgo_dict = {"Holli" : "Hollis ","NHoll" : "North-Hollis","NShel" : "North-Shellmound","ShPo" : "Shellmound-Powell","SSPM" : "SoShell-Powell - The Marina","SSPT" : "SoShell-Powell - The Towers","SHoll" : "South-Hollis","WExp" : "Watergate-Express"}
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
		arr.append(r)
	dic = {}
	headers = arr[1]
	arr = arr[1:]
	for header in range(0,len(headers)):
		dic[headers[header]] = header

	for i in range(1,len(arr)):
		val = arr[i]
		for idx,v in enumerate(val):
			if idx == dic["id"]:
				nws.cell(row=i,column= 1).value = v
			elif idx == dic["InterviewersInitials"]:
				nws.cell(row=i,column= 2).value = v
			elif idx == dic["g1xRoutexSWCx0"]:
	 			nws.cell(row=i,column= 3).value = ags[v] 
			elif idx == dic["g1xRoutexSWCx0[other]"]:
				nws.cell(row=i,column= 4).value = v
			elif idx == dic["xTimeofDayx0"]:
				nws.cell(row=i,column= 5).value = v
			elif idx == dic["xDirectionx0"]:
				nws.cell(row=i,column= 6).value = direction[v]
			elif idx == dic["NumberOfRiders"]:
				nws.cell(row=i,column= 7).value = v
			elif idx == dic["startLocOfDevice"]:
				#check if location is there or not
				if v:
					loc = [x.strip() for x in v.split(',')]
					nws.cell(row=i,column= 8).value = loc[0]
					nws.cell(row=i,column= 9).value = loc[1]
			#origin related columns
			elif idx == dic["g1xOriginx0"]:
				nws.cell(row=i,column= 10).value = 14 if (v == "-oth-") else v
			elif idx == dic["g1xOriginx0[other]"]:
				nws.cell(row=i,column= 11).value = v
			elif idx == dic["g1xMapx10[1]"]:
				#check if location is there or not
				if v:
					loc = [x.strip() for x in v.split(',')]
					nws.cell(row=i,column= 12).value = loc[0]
					nws.cell(row=i,column= 13).value = loc[1]
			elif idx == dic["g1xMapx10[2]"]:
				nws.cell(row=i,column= 14).value = v
			#destination related columns
			elif idx == dic["g1xDestix0"]:
				nws.cell(row=i,column= 15).value = 14 if (v == "-oth-") else v
			elif idx == dic["g1xDestix0[other]"]:
				nws.cell(row=i,column= 16).value = v
			elif idx == dic["g1xMapx10[6]"]:
				#check if location is there or not
				if v:
					loc = [x.strip() for x in v.split(',')]
					nws.cell(row=i,column= 17).value = loc[0]
					nws.cell(row=i,column= 18).value = loc[1]
			elif idx == dic["g1xMapx10[7]"]:
				nws.cell(row=i,column= 19).value = v
			elif idx == dic["AccessMode"]:
				#20 column is for roundtrip
				nws.cell(row=i,column= 21).value = 9 if (v == "-oth-") else v
			elif idx == dic["AccessMode[other]"]:
				nws.cell(row=i,column= 22).value = v
			elif idx == dic["AccessMinutes"]:
				nws.cell(row=i,column= 23).value = v
			elif idx == dic["AccessMiles"]:
				nws.cell(row=i,column= 24).value = v
			elif idx == dic["FirstB4Trans"]:
				#25 column is for calculating transfers
				global xc
				xc = v
				nws.cell(row=i,column= 26).value = v[1:] if v else None
			elif idx == dic["g1xAgencyx1"] and xc == 'A1':
				nws.cell(row=i,column= 27).value = transit_agencies_dict[v]
				global bagc1
				bagc1 = v
				#print bagc1
			elif idx == dic["g1xAgencyx1[other]"]:
				nws.cell(row=i,column= 28).value = v
			#AC Route
			elif idx == dic["g1xRoutexACx1"] and bagc1 == 'AC':
				global ac_route1
				ac_route1 = v
				nws.cell(row=i,column= 29).value = v
			elif idx == dic["g1xRoutexACx1[other]"] and bagc1 == 'AC' and ac_route1 == None:
				nws.cell(row=i,column= 29).value = v
			#Bart Route
			elif idx == dic["g1xRoutexBARTx1"] and bagc1 == 'BA':
				global ba_route1
				ba_route1 = v
				try:
					if  bart_dict[v]:
						nws.cell(row=i,column= 29).value = bart_dict[v]
				except Exception as e:
					print "Error.."
				else:
					pass
			elif idx == dic["g1xRoutexBARTx1[other]"] and bagc1 == 'BA' and ba_route1 == None:
				nws.cell(row=i,column= 30).value = v#bart_dict[v]
			elif idx == dic["g1xRoutexBAx1"] and bagc1 == 'BA':
				pass
			#CC Route
			elif idx == dic["g1xRoutexCCx1"] and bagc1 == 'CC':
				global cc_route1
				cc_route1 = v
				nws.cell(row=i,column= 29).value = v
			elif idx == dic["g1xRoutexCCx1[other]"] and bagc1 == 'CC' and cc_route1 == None:
				nws.cell(row=i,column= 30).value = v
			#Emgo Route
			elif idx == dic["g1xRoutexEMGOx1"] and bagc1 == 'EM':
				global em_route1
				em_route1 = v
				nws.cell(row=i,column= 29).value = v
			elif idx == dic["g1xRoutexEMGOx1[other]"] and bagc1 == 'EM' and em_route1 == None:
				nws.cell(row=i,column= 30).value = v
			elif idx == dic["g1xRoutexEMx1"] and bagc1 == 'EM':
				pass
			#Fast Route
			elif idx == dic["g1xRoutexFSx1"] and bagc1 == 'FS':
				global fs_route1
				fs_route1 = v
				nws.cell(row=i,column= 29).value = v
			elif idx == dic["g1xRoutexFSx1[other]"] and bagc1 == 'FS' and fs_route1 == None:
				nws.cell(row=i,column= 30).value = v
			#Golden Gate
			elif idx == dic["g1xRoutexGGx1"] and bagc1 == 'GG':
				global gg_route1
				gg_route1 = v
				nws.cell(row=i,column= 29).value = v
			elif idx == dic["g1xRoutexGGx1[other]"] and bagc1 == 'GG' and gg_route1 == None:
				nws.cell(row=i,column= 30).value = v
			#Samtrans
			elif idx == dic["g1xRoutexSMx1"] and bagc1 == 'SM':
				global sm_route1
				sm_route1 = v
				nws.cell(row=i,column= 29).value = v
			elif idx == dic["g1xRoutexSMx1[other]"] and bagc1 == 'SM' and sm_route1 == None:
				nws.cell(row=i,column= 30).value = v
			#MUNI
			elif idx == dic["g1xRoutexSFx1"] and bagc1 == 'SF':
				global sf_route1
				sf_route1 = v
				nws.cell(row=i,column= 29).value = v
			elif idx == dic["g1xRoutexSFx1[other]"] and bagc1 == 'SF' and sf_route1 == None:
				nws.cell(row=i,column= 30).value = v
			#Soltrans
			elif idx == dic["g1xRoutexSTx1"] and bagc1 == 'ST':
				global st_route1
				st_route1 = v
				nws.cell(row=i,column= 29).value = v
			elif idx == dic["g1xRoutexSTx1[other]"] and bagc1 == 'ST' and st_route1 == None:
				nws.cell(row=i,column= 30).value = v
			#Napa Vine
			elif idx == dic["g1xRoutexVNx1"] and bagc1 == 'VN':
				global vn_route1
				vn_route1 = v
				nws.cell(row=i,column= 29).value = v
			elif idx == dic["g1xRoutexVNx1[other]"] and bagc1 == 'VN' and vn_route1 == None:
				nws.cell(row=i,column= 30).value = v
			#WC
			elif idx == dic["g1xRoutexWCx1"] and bagc1 == 'WC':
				global wc_route1
				wc_route1 = v
				nws.cell(row=i,column= 29).value = v
			elif idx == dic["g1xRoutexWCx1[other]"] and bagc1 == 'WC' and wc_route1 == None:
				nws.cell(row=i,column= 30).value = v
		nwb.save("qm.xlsx")
	os.system("start " + "qm.xlsx")
		
export()