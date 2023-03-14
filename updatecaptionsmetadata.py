#! /usr/bin/python3

import re
import sys
import os
import openpyxl
import warnings

warnings.filterwarnings('ignore', category=DeprecationWarning)


sheet = '1. Master Metadata'
hns   = []

START_ROW = 4
START_COL = 2

#get house numbers from the command line and return the list
def get_HouseNumbers():

	hnList = []

	for n in range(1,len(sys.argv)):

		capf = sys.argv[n]

		if os.path.isfile(capf):
			strList = capf.split('.')
			hn  = strList[0]
			m   = re.match("^BUZ_[A-Z0-9]+",hn)
			if m:
				hnList.append(hn)

	return hnList

#get inputfrom the user and return it
def get_input(question):

	print()
	data = input(question)
	return data

def getColNumNum(xlf,sheet,colname):


	stop = True

	for c in range(START_COL,100):

		txt = str(ws.cell(row=START_ROW,column=c).value)
		if txt == colname:
			return c

	print('\n   ~~~Could not find Column:', colname, '~~~')

	wb.save(xlf)
	wb.close()

	sys.exit(1)

	return c


#get info from the user 
vidpth = get_input('Give me your video s3 path: ')
cappth = get_input('Give me your captn s3 path: ')
xlf    = get_input('Give me your xl file name : ')
hns    = get_HouseNumbers()

#open the Metadata xlf for read/write
try:
	wb = openpyxl.load_workbook(xlf)
except:
	print('\n','   ~~~Cannot open Metadata Sheet~~~\n')
	sys.exit(1)

#open the '1. Master Metadata' sheet for read/write
try:
	ws = wb[sheet]
except:
	print('\n','   ~~~Cannot open Metadata Sheet~~~\n')
	sys.exit(1)

#Supplier.OriginalName
#Fremantle.HouseNumber
#TWK.AncillaryName
#getColNumNum(xlf,sheet,colname):

movcol = getColNumNum(xlf,ws,'Supplier.OriginalName')
hncol  = getColNumNum(xlf,ws,'Fremantle.HouseNumber')
capcol = getColNumNum(xlf,ws,'TWK.AncillaryName')

print(movcol,':',capcol,':',hncol)
wb.save(xlf)
wb.close()


