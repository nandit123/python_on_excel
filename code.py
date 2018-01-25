import xlrd
import xlwt
from xlutils.copy import copy

workbook = xlrd.open_workbook('data.xlsx')
worksheet = workbook.sheet_by_index(0)

workbook2 = copy(workbook)
worksheet2 = workbook2.get_sheet(0)


for j in range(5, 10):
	count = 0
	total = 0
	for i in range(2, 63): # i goes from 2 to 63
		if worksheet.cell(i, j).value != xlrd.empty_cell.value:	
			total += float(worksheet.cell(i, j).value)
			count+=1
	average = total/count
	print(average)

	for i in range(2,63):
		if worksheet.cell(i, j).value == xlrd.empty_cell.value:
			worksheet2.write(i, j, average)
workbook2.save('data.xlsx')