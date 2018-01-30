import numpy
import xlrd
import xlwt
from xlutils.copy import copy
from scipy import stats
import matplotlib.pyplot as plt

workbook = xlrd.open_workbook('data.xlsx')
worksheet = workbook.sheet_by_index(0)

workbook2 = copy(workbook)
worksheet2 = workbook2.get_sheet(0)


#for internal component 1
stdBefore = []
stdAfter = []
option = int(input('Type 1 for Mean, 2 for Mode and 3 for Median based method : '))
for j in range(5, 10):
	l = []
	for i in range(2, 63): # i goes from 2 to 63
		if worksheet.cell(i, j).value != xlrd.empty_cell.value:	
			l.append(float(worksheet.cell(i, j).value))
	a = numpy.array(l)
	m = numpy.mean(a)
	mo = stats.mode(a)
	me = numpy.median(a)
	print('mean for column is ',m)
	print('mode for column is ',mo[0][0])
	print('median for column is ',me)

	s = numpy.std(a)
	stdBefore.append(s)

	print('--- Column ', j-4, ' ---')
	# filling missing values
	if(option == 1):
		for i in range(2,63):
			if worksheet.cell(i, j).value == xlrd.empty_cell.value:
				worksheet2.write(i, j, m)
	elif(option == 2):
		for i in range(2,63):
			if worksheet.cell(i, j).value == xlrd.empty_cell.value:
				worksheet2.write(i, j, mo[0][0])
	elif(option == 3):
		for i in range(2,63):
			if worksheet.cell(i, j).value == xlrd.empty_cell.value:
				worksheet2.write(i, j, me)
	else:
		print('Invalid Option, Try Again :(')			
workbook2.save('data2.xlsx')

workbook = xlrd.open_workbook('data2.xlsx')
worksheet = workbook.sheet_by_index(0)

for j in range (5, 10):
	l2 = []
	for i in range(2, 63):
		l2.append(float(worksheet.cell(i, j).value))
	a = numpy.array(l2)
	s = numpy.std(a)
	stdAfter.append(s)
print('************')

differenceInStd = 0;
for i in range(0, 5):
	print('Column ', i+1, ' stdBefore = ', stdBefore[i], ' & stdAfter = ', stdAfter[i])
	differenceInStd += (stdBefore[i]-stdAfter[i])

# average difference is now calculated and is termed as fluctuation
fluctuation = float(differenceInStd/5)
print('###Fluctuation = ', fluctuation, ' ###')

# for plotting graph using matplotlib of python
columns = [1,2,3,4,5]
plt.plot(columns, stdBefore, 'ro')
plt.plot(columns, stdAfter, 'o')
plt.ylabel('Standard Deviation')
plt.xlabel('Columns of Data-> red - before :: blue - after')
plt.show()