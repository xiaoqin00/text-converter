#!/usr/bin/env python
# -*- coding: utf-8 -*-
# Created by xiaoqin00 on 2017/6/22

import sys
from pdfminer.pdfinterp import PDFResourceManager,PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage
from optparse import OptionParser
#main
def convert(argv) :
    #输出文件名，这里只处理单文档，所以只用了argv［1］
    outfile = argv[1] + '.txt'
    args = [argv[1]]

    debug = 0
    pagenos = set()
    password = ''
    maxpages = 0
    rotation = 0
    codec = 'utf-8'   #输出编码
    caching = True
    imagewriter = None
    laparams = LAParams()
    #
    PDFResourceManager.debug = debug
    PDFPageInterpreter.debug = debug

    rsrcmgr = PDFResourceManager(caching=caching)
    outfp = file(outfile,'w')
#pdf转换
    device = TextConverter(rsrcmgr, outfp, codec=codec, laparams=laparams,
                imagewriter=imagewriter)

    for fname in args:
        fp = file(fname,'rb')
        interpreter = PDFPageInterpreter(rsrcmgr, device)
#处理文档对象中每一页的内容
        for page in PDFPage.get_pages(fp, pagenos,
                          maxpages=maxpages, password=password,
                          caching=caching, check_extractable=True) :
            page.rotate = (page.rotate+rotation) % 360
            interpreter.process_page(page)
        fp.close()
    device.close()
    outfp.close()
    return

if __name__ == '__main__':
    parser=OptionParser(usage='%prog [options]')
    parser.add_option('-i','--in',dest='input',help='input file')
    parser.add_option('-o','--out',dest='output',help='output file')
    (options,args)=parser.parse_args()
    # print options.input
    convert()