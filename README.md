# python_on_excel
Using python libraries to perform operations on excel data.
Here the given data is in data.xlsx excel file which contains some missing values. We need to fill up those missing values with some values so that data as whole starts making more sense.

What is done ->
I have created a code_numpy.py file that operates on data.xlsx file to fill up missing values and save the result on data2.xlsx file. Missing values are filled with either mean, mode or median depending on the method user selects when the python program is run. Then standard deviation is found before and after the filling of missing values as per the method(mean, mode or median). A graph is then plotted for standard deviation vs column number (containing values).

Libraries used - 
numpy for mean and median
scipy for mode
matplotlib for plotting the graph
xlrd and xlwt for reading and writing on excel file (else directly xlutils can also be used)

Important: Here excel file have been parsed and operations are the conducted. Other way could be exporting CSV file from excel and then conducting operations on the CSV file. CSV file will automatically eliminate the empty cells. :)



