#!python2.7
import sys
import os
# pip install pypiwin32
import win32com
import win32com.client as win32
from win32com.client import Dispatch
# pip install Pillow
from PIL import ImageGrab

# get fullpath
f = open(sys.argv[1], "r")
fullpath = os.path.realpath(f.name)
directory = os.path.dirname(fullpath)
f.close()

# Export pngs
xlApp = Dispatch('Excel.Application')
workbook = xlApp.Workbooks.Open(fullpath, None, True)
for sheet in xlApp.Sheets:
	
	print "Processing", sheet.Name
	column = 1
	table_counter = 1
	while column < 20:
		value = sheet.Cells(1 , column)
		start_column = column
		start_line = 1
		if str(value) != "None":
			end_column = column
			while str(value) != "None":
				value = sheet.Cells(1 , column)
				column = column + 1
			end_column = column - 2
			line = 1
			value = sheet.Cells(line , start_column)
			while str(value) != "None":
				value = sheet.Cells(line , start_column)
				line = line + 1
			end_line = line - 2
			win32c = win32com.client.constants
			sheet.Range(sheet.Cells(start_line, start_column),sheet.Cells(end_line,end_column)).CopyPicture(Format= 2)    
			workbook = xlApp.Workbooks.Add()
			tmpsheet = workbook.ActiveSheet
			tmpsheet.Paste()
			tmpsheet.Shapes('Picture 1').Copy()
			img = ImageGrab.grabclipboard()
			imgFile = os.path.join(directory, sheet.Name + "_table_" + str(table_counter) + ".png")
			img.save(imgFile,'PNG')
			workbook.Close(False)
			table_counter = table_counter + 1
		else:
			column = column + 1

	for chart in sheet.ChartObjects():
		graphFile = os.path.join(directory, sheet.Name + "_graph_" + chart.Name + ".png")
		chart.Chart.Export(graphFile)

