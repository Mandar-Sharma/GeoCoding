import xlrd
import xlwt
import geocoder
import json

def truncate(f, n):
    '''Truncates/pads a float f to n decimal places without rounding'''
    s = '{}'.format(f)
    if 'e' in s or 'E' in s:
        return '{0:.{1}f}'.format(f, n)
    i, p, d = s.partition('.')
    return '.'.join([i, (d+'0'*n)[:n]])


def workon_file(path):
	workbook = xlwt.Workbook() 
	sheet = workbook.add_sheet("Matched Output") 
	with open('JsonData.txt', 'r') as outfile:
		json_data = json.load(outfile)
	book = xlrd.open_workbook(path)
	first_sheet = book.sheet_by_index(0)
	#range(1501,3001)
	sheet_index = 1
	for row_index in range(1501,3001):
		value = first_sheet.cell(row_index,1).value
		value_list = value.split(";")
		for i in range(len(value_list)):
			check_loco = geocoder.google(value_list[i])
			if check_loco.latlng:
				check_loco_lat = truncate(check_loco.latlng[0],2)
				check_loco_long = truncate(check_loco.latlng[1],2)
				for i in range(len(json_data)):
					json_list_ele= json_data[i]
					for key in json_list_ele:
						if json_list_ele[key]:
							lati = truncate(json_list_ele[key][0], 2)
							longi = truncate(json_list_ele[key][1], 2)
							if check_loco_lat == lati and check_loco_long == longi:
								sheet.write(sheet_index, 0, value)
								sheet.write(sheet_index, 1, key)
								sheet_index = sheet_index + 1
	workbook.save("Output.xls")						

workon_file("Address-datacleaning.xlsx")