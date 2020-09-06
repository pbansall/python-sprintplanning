from openpyxl import load_workbook
import numpy as np
import math
import pandas as pd
import fire

import warnings
warnings.filterwarnings("ignore")

columns = {}
# Correction of headers for Sprint Sheets
def headerCorrection(sheetname,sheet):
    # moving row 1 to header
	columns[sheetname] = sheet[sheetname].columns
	sheet[sheetname].columns = sheet[sheetname].iloc[0]
	sheet[sheetname] = sheet[sheetname].iloc[1:]

#Changing column names as desired
def changeColumnName(sheetname, oldname, newname,sheet):
	sheet[sheetname].rename(columns = {oldname:newname}, inplace = True)

# Deleting the existing sheet
# Loading new sheet with new data
def loadSheet(bookname,sheetname,sheet):
	#print(sheet[sheetname].loc[0])
	book = load_workbook(bookname)
	deleteSheet = book.get_sheet_by_name(sheetname)
	book.remove_sheet(deleteSheet)
	with pd.ExcelWriter(bookname,engine = 'openpyxl') as writer:
		writer.book = book
		#writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
		sheet[sheetname].to_excel(writer,sheet_name = sheetname,index = False)

def buildSprint(sprint,df,bookname,sheet):
	#print (f'inside buildSprint {sprint}')
	#print (sheet.keys())
	headerCorrection(sprint,sheet)
	changeColumnName(sprint,'Billing & Revenue Mgmt','Billing', sheet)
	cols = list(sheet[sprint].columns)
	#print ('After that')
	sheet[sprint] = sheet[sprint][0:0]
	newrow = {}
	for index, row in df.iterrows():
		components = str(row['Impacted Product']).split(',')
		components = list(map(str.strip, components))
		components = list(map(str.lower,components))
		for col in cols:
			if col.lower() in components:
				newrow[col] = 'X'
			elif col.lower().strip() == 'feature name' :
				newrow[col] = row['Feature Name']
			else:
				newrow[col] = ''
		sheet[sprint] = sheet[sprint].append(newrow,ignore_index = True)
	sheet[sprint].sort_values(["Feature Name"],ascending = True, inplace = True)
	sheet[sprint].drop_duplicates(subset=None, keep='first', inplace=True)
	changeColumnName(sprint,'Billing','Billing & Revenue Mgmt',sheet)
	sheet[sprint] = sheet[sprint].columns.to_frame().T.append(sheet[sprint], ignore_index=True)
	sheet[sprint].columns = columns[sprint]
	loadSheet(bookname,sprint,sheet)	


def run(bookname):
	sheet = pd.read_excel(bookname,sheet_name= None)
	#print(sheet.keys())
	for i in set(sheet['User Story']['Sprint ']):
		if not math.isnan(i):
			df = sheet['User Story'].loc[sheet['User Story']['Sprint '] == int(i)].iloc[:,:9]
			sprint = f'Sprint {int(i)}'
			print(sprint)
			buildSprint(sprint,df,bookname,sheet)
	print(columns)

if __name__ == '__main__':
	fire.Fire(run)