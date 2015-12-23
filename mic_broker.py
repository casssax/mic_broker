import sys
import csv
import pywintypes
import win32com.client as win32
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox



def aggr_counts(input_file):
	# read in VOIDDATE.TXT file to get counts.
	data = [row for row in input_file]
	#print len(data)
	key_dict = {}
	for row in data:
		# get original key from pos.366(4)
		orig_key = row[365:369]
		# get package code from pos. 278(3)
		package = row[277:280]
		# add keys/quantities to dictionary if there are non ZLD records
		if package != 'ZLD':
			if orig_key in key_dict:
				key_dict[orig_key] += 1
			else:
				key_dict[orig_key] = 1
	return key_dict

def get_broker_codes(list):
	row = 4
	lol_dict = {}
	key_cell = 'x'
	while key_cell != 'None':
		key_cell = str(list.Cells(row,4))
		key_cell = key_cell.strip()
		broker_cell = str(list.Cells(row,3))
		broker_cell = broker_cell.strip()
		#print(key_cell, broker_cell)
		if key_cell in lol_dict:
			pass
		else:
			if broker_cell != 'SUPP' and broker_cell != 'None':
				lol_dict[key_cell] = broker_cell
		row += 1
	print(lol_dict)
	return lol_dict



def populate_excel(input_file,broker_report,list_of_lists):
	#read in list of lists
	try:
		lol = excel.Workbooks.Open(list_of_lists)
	except:
		print("Failed to open list of lists")
		sys.exit(1)
	lols = lol.Sheets('List of Lists')
	lol_dict = get_broker_codes(lols)
	lol.Close(True)
	#read in excel file
	try:
		wb = excel.Workbooks.Open(broker_report)
	except:
		print("Failed to open broker report")
		sys.exit(1)
	ws = wb.Sheets('PURGE DROPS')
	ws.Range("D1").EntireColumn.Clear()
	ws.Range("L1").EntireColumn.Delete()
	ws.Range("L1").EntireColumn.Delete()
	ws.Range("C1").EntireColumn.Insert()
	key_dict = aggr_counts(input_file)
	val = ''
	# first usefull row in excel sheet is A7
	row = 7
	# add column header to output column
	ws.Cells(4,14).Value = 'QTY'
	ws.Cells(5,14).Value = 'MAILED'
	# add column header to broker column
	ws.Cells(5,3).Value = 'VENDER'
	ws.Cells(5,5).Value = 'REJECTS'
	# add column header to adj qty
	ws.Cells(4,13).Value = 'ADJ'
	ws.Cells(5,13).Value = 'QTY'
	# loop through keycodes
	while val != 'TOTALS':
		val = ws.Cells(row,1).Value
		#print 'val: ', val
		if val != None:
			val = val.strip()
		# add quantity of non ZLD records on output, if any.
		if val in key_dict:
			ws.Cells(row,14).Value = key_dict[val]
			ws.Cells(row,3).Value = lol_dict[val]
			ws.Cells(row,5).Value = ws.Cells(row,13).Value - ws.Cells(row,14).Value
		row += 1
	row = row - 1
	total_sum = '=SUM(O7:O' + str(row - 2) + ')'
	#ws.Cells(row,14).Formula = '=SUM(N7:N' + str(row - 2) + ')'

	# format cells with font, size, alignment
	ws.Range(ws.Cells(4,14),ws.Cells(row,3)).Font.Name = "Courier"
	ws.Range(ws.Cells(4,14),ws.Cells(row,3)).Font.Size = 8
	ws.Range(ws.Cells(4,14),ws.Cells(row,3)).HorizontalAlignment = win32.constants.xlRight
	#ws.Range(ws.Cells(5,14),ws.Cells(row,3)).HorizontalAlignment = win32.constants.xlRight
	#ws.Range(ws.Cells(5,14),ws.Cells(row,3)).NumberFormat = "###,##0"

	ws.Cells(row,5).Formula = '=SUM(E7:E' + str(row - 2) + ')'
	ws.Range(ws.Cells(4,13),ws.Cells(row,3)).HorizontalAlignment = win32.constants.xlRight
	ws.Range(ws.Cells(5,13),ws.Cells(row,3)).HorizontalAlignment = win32.constants.xlRight
	ws.Columns.AutoFit()
	# delete DE INPUT column
	ws.Range("F1").EntireColumn.Delete()
	# delete TOTALS row
	ws.Rows(row).EntireRow.Delete()
	wb.Close(True)
	excel.Quit()





excel = win32.gencache.EnsureDispatch('Excel.Application')
excel.Visible = True

# output_file_name = sys.argv[1]
# broker_report = sys.argv[2]
# prevent root window from appearing on screen
root = tk.Tk()
root.withdraw()
# select voiddate file dialog box
messagebox.showinfo(message = "Select the voiddate.txt file:")
output_file_name = filedialog.askopenfilename()
print('voiddate file: ', output_file_name)
# select broker report file dialog box
messagebox.showinfo(message = "Select the EXCEL.XLS file:")
broker_report = filedialog.askopenfilename()
print('broker_report: ', broker_report)
# select list of lists file dialog box
messagebox.showinfo(message = "Select the list of lists file:")
list_of_lists = filedialog.askopenfilename()
print('list_of_lists: ', list_of_lists)

with open(output_file_name, 'r') as input_file:
                populate_excel(input_file, broker_report,list_of_lists)
