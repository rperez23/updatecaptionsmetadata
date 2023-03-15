#! /usr/bin/python3

import re
import sys
import os
import openpyxl
import warnings

warnings.filterwarnings('ignore', category=DeprecationWarning)
warnings.simplefilter(action='ignore', category=UserWarning)


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


	for c in range(START_COL,100):

		txt = str(ws.cell(row=START_ROW,column=c).value)
		if txt == colname:
			return c

	print('\n   ~~~Could not find Column:', colname, 'in Metadata Sheet~~~\n')

	wb.save(xlf)
	wb.close()

	sys.exit(1)

	return c

def getxldata(ws,hn,epc,hnc,capc):

	r = START_ROW + 2

	fname     = ''
	capprefix = ''
	counter = 0 

	while counter < 10:

		txt = str(ws.cell(row=r,column=hnc).value)

		if txt == hn:

			fname     = str(ws.cell(row=r,column=epc).value)
			capprefix = str(ws.cell(row=r,column=capc).value)

			return fname, capprefix

		elif txt == 'None':

			counter += 1

		r += 1

	return fname, capprefix

def getversion(txt):

	newfname = ''

	m = re.search('v(\d+)_(\d{8}\.[a-zA-Z]+)$',txt)

	if m:
		n       = int(m.group(1)) + 1
		
		matched = m.group(0)
		before   = txt[:m.start()]
		newfname = before + 'v' + str(n) + '_' + m.group(2)
		#print(txt, ':', newfname)

	else:
		n        = 2
		p        = re.search('(_\d{8}\.[a-zA-Z]+)$',txt)
		matched  = p.group(0)
		before   = txt[:p.start()]
		newfname = before + '_v' + str(n) + matched
		#print(txt, ':', newfname)

	return newfname

def getcaptiontype(capf):

	scc = capf + '.scc'
	srt = capf + '.srt'

	if os.path.isfile(scc):
		return 'scc'
	elif os.path.isfile(srt):
		return 'srt'
	
	return ''


#get info from the user 
vidpth  = get_input('Give me your video s3 path: ')
cappth  = get_input('Give me your captn s3 path: ')
xlf     = get_input('Give me your xl file name : ')
hns     = get_HouseNumbers()

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


epcol  = getColNumNum(xlf,ws,'Supplier.OriginalName')
hncol  = getColNumNum(xlf,ws,'Fremantle.HouseNumber')
capcol = getColNumNum(xlf,ws,'TWK.AncillaryName')

for i in range(0,len(hns)):

	#get a house number from the hns list
	hn = hns[i]

	epname, prefix = getxldata(ws,hn,epcol,hncol,capcol)

	if epname == '' or prefix == '':
		print(hn,': SKIPPING')
	else:
				
		newepname   = getversion(epname)
		capext      = getcaptiontype(hn)
		capfname    = prefix + '.' + capext
		parts       = newepname.split('.')
		newcapname  = parts[0] + '.' + capext

		#print(hn,':',newepname,':',newcapname)
		#BUZ_LMAD03247 : LetsMakeADeal_s2012_e4074_20230227.mxf : LetsMakeADeal_s2012_e4074_v2_20230227.mxf : LetsMakeADeal_s2012_e4074_20230227.scc
		#BUZ_LMAD03248 : LetsMakeADeal_s2012_e4075_20230227.mxf : LetsMakeADeal_s2012_e4075_v2_20230227.mxf : LetsMakeADeal_s2012_e4075_20230227.scc

		#link the caption file
		lncmd = 'ln ' + hn + '.' + capext + ' ' + newcapname
		print(lncmd)
		statln = os.system(lncmd)

		



wb.save(xlf)
wb.close()


