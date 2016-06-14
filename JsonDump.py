import xlrd
import geocoder
import json

def workon_file(path):
	book = xlrd.open_workbook(path)
	first_sheet = book.sheet_by_index(1)
	second_sheet = book.sheet_by_index(1)
	listofunis = []
	for row_index in range(second_sheet.nrows):
		uni = second_sheet.cell(row_index,1).value
		g = geocoder.google(uni)
		geoco = g.latlng
		eachdict = {uni : geoco}
		listofunis.append(eachdict)
		print(eachdict)
	with open('JsonData.txt', 'w') as outfile:
		json.dump(listofunis, outfile)

workon_file("Address-datacleaning.xlsx")