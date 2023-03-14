#! /usr/bin/python3

import re
import sys
import os
import openpyxl

sheet = '1. Master Metadata'
hns   = []

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


#get info from the user 
vidpth = get_input('Give me your video s3 path: ')
cappth = get_input('Give me your captn s3 path: ')
xlf    = get_input('Give me your xl file name : ')
hns    = get_HouseNumbers()

#open the Metadata sheet for read/write
try:
	wb = openpyxl.load_workbook(xlf)
except:
	print('\n','   ~~~Cannot open Metadata Sheet~~~\n')
	sys.exit(1)

wb.save(xlf)
wb.close()




