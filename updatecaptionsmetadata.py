#! /usr/bin/python3

import os
import openpyxl

#get inputfrom the user and return it
def get_input(question):

	print()
	data = input(question)
	return data


#get info from the user 
vidpth = get_input('Give me your video s3 path: ')
cappth = get_input('Give me your captn s3 path: ')
xlf    = get_input('Give me your xl file name : ')

#print('results:', '\n', vidpth, '\n', cappth, '\n', xlf)


