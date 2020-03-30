#!/usr/bin/python3

from dns import reversename, resolver
import sys
import argparse
import csv
import socket
import xlrd
from xlutils.copy import copy

fnamei = ""
fnameo = ""
rnd_nw = "10.12." # default network pattern
firstrow = 1 # default second row
dnscolumn = 1 # default second column

parser = argparse.ArgumentParser(description='This script read the first column of the given excel sheet and write back the reverse DNS to the second column')
parser.add_argument('i', metavar='inputfile', action="store", help="input excel file name")
parser.add_argument('-o', metavar='', action="store", help="output excel file name")
parser.add_argument('--nw', metavar='', action="store", help="network pattern to filter IP")
parser.add_argument('--first_row', metavar='', action="store", help="first row where we start to process the IPs", type=int)
parser.add_argument('--dns_column', metavar='', action="store", help="dns column where we put the DNS names", type=int)

if parser.parse_args().i:
	fnamei = parser.parse_args().i
if parser.parse_args().o:
	fnameo = parser.parse_args().o
else:
	fnameo = fnamei
if parser.parse_args().nw:
	rnd_nw = parser.parse_args().nw
if parser.parse_args().first_row:
	firstrow = parser.parse_args().first_row
if parser.parse_args().dns_column:
	dnscolumn = parser.parse_args().dns_column


def reversedns(cella):
	revdns = reversename.from_address(cella)
	try:
		return str(resolver.query(revdns,"PTR")[0])
	except resolver.NXDOMAIN as e:
		return "NXDOMAIN"

def validateip(ipaddr):
	try:
		socket.inet_aton(ipaddr)
	except socket.error:
		return False
	return True

xbook1 = xlrd.open_workbook(fnamei)
xsheet1 = xbook1.sheet_by_index(0)
nrows = xsheet1.nrows

xbook2 = copy(xbook1)
xsheet2 = xbook2.get_sheet(0)

dname = "reversednsfailed"
#num_cols = xsheet.ncols
for row_idx in range(firstrow, nrows):
	ipvalid = True
	celly = xsheet1.cell_value(row_idx, 0)
	
	if validateip(celly):
		if celly[:len(rnd_nw)] == rnd_nw:
			dname = reversedns(celly)
			#print(celly+';'+dname)
			xsheet2.write(row_idx, dnscolumn, dname)

xbook2.save(fnameo)
