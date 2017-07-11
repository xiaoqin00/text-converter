#!/usr/bin/env python
# -*- coding: utf-8 -*-
# Created by xiaoqin00 on 2017/7/10

import xlrd
from optparse import OptionParser

def convert():
    try:
        data = xlrd.open_workbook(input)
        table = data.sheets()[0]
        nrows = table.nrows
        ncols=table.ncols
        print nrows, type(nrows)
        f = open(output, 'w')

        for i in range(nrows):
            tmp=''
            for j in range(ncols):
                print table.cell(i,j)
                tmp = tmp + str(table.cell(i, j)).split(':')[1] + ' '
            f.write(tmp + '\n')
        f.close()
    except Exception,e:
        print e

if __name__ == '__main__':
    parser=OptionParser(usage='%prog [options]')
    parser.add_option('-i','--in',dest='input',help='input file')
    parser.add_option('-o','--out',dest='output',help='output file')
    (options,args)=parser.parse_args()
    input=options.input
    output=options.output
    convert()