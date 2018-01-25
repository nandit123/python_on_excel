import numpy
import xlrd
import xlwt
from xlutils.copy import copy

workbook = xlrd.open_workbook('data2.xlsx')
worksheet = workbook.sheet_by_index(0)

workbook2 = copy(workbook)
worksheet2 = workbook2.get_sheet(0)


#for internal component 1

for j in range(5, 10):
	l = []
	for i in range(2, 63): # i goes from 2 to 63
		if worksheet.cell(i, j).value != xlrd.empty_cell.value:	
			l.append(float(worksheet.cell(i, j).value))
	a = numpy.array(l)
	m = numpy.median(a)
	print('median is ',m)
	for i in range(2,63):
		if worksheet.cell(i, j).value == xlrd.empty_cell.value:
			worksheet2.write(i, j, m)
workbook2.save('data2.xlsx')