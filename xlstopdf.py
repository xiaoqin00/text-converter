#!/usr/bin/env python
# -*- coding: utf-8 -*-
# Created by xiaoqin00 on 2017/6/26

from win32com import client
import os
from optparse import OptionParser

def convert():
    xlApp = client.Dispatch("Excel.Application")
    input=os.path.abspath(options.input)
    output=os.path.abspath(options.output)
    books = xlApp.Workbooks.Open(input)
    ws = books.Worksheets[0]
    ws.Visible = 1
    ws.ExportAsFixedFormat(0, output)

if __name__ == '__main__':
  parser = OptionParser(usage='%prog [options]')
  parser.add_option('-i', '--in', dest='input', help='input file')
  parser.add_option('-o', '--out', dest='output', help='output file')
  (options, args) = parser.parse_args()
  convert()