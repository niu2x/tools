#-*- coding: utf8 -*- 

import xlrd
import sys
import os

def convert(filename):
	if filename == None:
		return
	if not os.path.isfile(filename):
		return
	workbook = xlrd.open_workbook(filename)
	sheet = workbook.sheet_by_index(0)

	ret = []
	keyNames = []
	for i in xrange(sheet.ncols):
		cell = sheet.cell(0, i)
		keyNames.append(cell.value)

	for r in xrange(1, sheet.nrows):
		item = {}

		for c in xrange(0, sheet.ncols):
			cell = sheet.cell(r, c)
			item[keyNames[c]] = cell.value
			if cell.ctype == 4:
				item[keyNames[c]] = [False, True][cell.value]

		ret.append(item)
	return ret


def generateLua(array):
	luaStr = '''return {'''

	for item in array:
		luaStr = luaStr + '''{'''
		for key in item:
			value = item[key]
			if type(value) == int:
				value = str(value)
			elif type(value) == float:
				value = str(value)
			elif type(value) == bool:
				if value == 1:
					value = 'true'
				else:
					value = 'false'
			else:
				value = '"' + value.encode('utf8') + '"'

			luaStr = luaStr + '''[\"''' + key.encode('utf8') + '''\"] = ''' + value + ''' ,'''
		luaStr = luaStr + '''},'''
	luaStr = luaStr + '''}'''
	return luaStr

if __name__ =='__main__':
	filename = sys.argv[1]
	print generateLua(convert(filename))
