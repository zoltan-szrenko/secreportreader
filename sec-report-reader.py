#!/usr/bin/python3

from dns import reversename, resolver
import sys
import csv
import socket
import xlrd
from xlutils.copy import copy
#import xlwt

print("start")

fname = "report2.xlsx"
sheetname = "in"
rnd_nw = "10.12."
firstrow = 1
dnscolumn = 1

def reversedns(cella):
	revdns = reversename.from_address(cella)
	try:
		print('reversedns: '+dname)
		return str(resolver.query(revdns,"PTR")[0])
	except resolver.NXDOMAIN as e:
		return "NXDOMAIN"
		#print('NXDOMAIN')


xbook1 = xlrd.open_workbook(fname)
xsheet1 = xbook1.sheet_by_index(0)
nrows = xsheet1.nrows

xbook2 = copy(xbook1)
xsheet2 = xbook2.get_sheet(0)

dname = "dnstest"
#num_cols = xsheet.ncols
for row_idx in range(firstrow, nrows):
	celly = xsheet1.cell_value(row_idx, 0)
	if len(celly) >= 6:
		if celly[:len(rnd_nw)] == rnd_nw:
			dname = reversedns(celly)
			print('celly: '+celly+' dname: '+dname)
			xsheet2.write(row_idx, dnscolumn, dname)

xbook2.save(fname)
print("fin")
