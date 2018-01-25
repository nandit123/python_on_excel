import xlrd
import xlwt
from xlutils.copy import copy

workbook = xlrd.open_workbook('data.xlsx')
worksheet = workbook.sheet_by_index(0)

workbook2 = copy(workbook)
worksheet2 = workbook2.get_sheet(0)


#y = worksheet.cell(4,6).value
#worksheet2.write(4, 6, 2.05)
#workbook2.save('data.xlsx')
#print(y)


#for internal component 1
count = 0
total = 0
for i in range(2, 63): # i goes from 2 to 63
	if worksheet.cell(i, 5).value != xlrd.empty_cell.value:	
		total += float(worksheet.cell(i, 5).value)
		count+=1
average = total/count
print(average)

for i in range(2,63):
	if worksheet.cell(i, 5).value == xlrd.empty_cell.value:
		worksheet2.write(i, 5, average)
workbook2.save('data.xlsx')



#for internal component 2
count = 0
total = 0
for i in range(2, 63): # i goes from 2 to 63
	if worksheet.cell(i, 6).value != xlrd.empty_cell.value:	
		total += float(worksheet.cell(i, 6).value)
		count+=1
average = total/count
print(average)

for i in range(2,63):
	if worksheet.cell(i, 6).value == xlrd.empty_cell.value:
		worksheet2.write(i, 6, average)
workbook2.save('data.xlsx')


#for internal component 3
count = 0
total = 0
for i in range(2, 63): # i goes from 2 to 63
	if worksheet.cell(i, 7).value != xlrd.empty_cell.value:	
		total += float(worksheet.cell(i, 7).value)
		count+=1
average = total/count
print(average)

for i in range(2,63):
	if worksheet.cell(i, 7).value == xlrd.empty_cell.value:
		worksheet2.write(i, 7, average)
workbook2.save('data.xlsx')


#for internal component 4
count = 0
total = 0
for i in range(2, 63): # i goes from 2 to 63
	if worksheet.cell(i, 8).value != xlrd.empty_cell.value:	
		total += float(worksheet.cell(i, 8).value)
		count+=1
average = total/count
print(average)

for i in range(2,63):
	if worksheet.cell(i, 8).value == xlrd.empty_cell.value:
		worksheet2.write(i, 8, average)
workbook2.save('data.xlsx')


#for mid-semester
count = 0
total = 0
for i in range(2, 63): # i goes from 2 to 63
	if worksheet.cell(i, 9).value != xlrd.empty_cell.value:	
		total += float(worksheet.cell(i, 9).value)
		count+=1
average = total/count
print(average)

for i in range(2,63):
	if worksheet.cell(i, 9).value == xlrd.empty_cell.value:
		worksheet2.write(i, 9, average)
workbook2.save('data.xlsx')