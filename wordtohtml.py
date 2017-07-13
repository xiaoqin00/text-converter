#!/usr/bin/env python
# -*- coding: utf-8 -*-
# Created by xiaoqin00 on 2017/7/11


from win32com import client as wc
import os
from optparse import OptionParser

word = wc.Dispatch('Word.Application')


def wordsToHtml(input,output):
    # for path, subdirs, files in os.walk(dir):
    #     for wordFile in files:
    # wordFullName = os.path.join(path, wordFile)
    # print "word:" + wordFullName
    doc = word.Documents.Open(os.path.abspath(input))

    wordFile2 = unicode(input, "gbk")
    dotIndex = wordFile2.rfind(".")
    if (dotIndex == -1):
        print "********************ERROR: 未取得后缀名！"

    fileSuffix = wordFile2[(dotIndex + 1):]
    if (fileSuffix == "doc" or fileSuffix == "docx"):
        # fileName = wordFile2[: dotIndex]
        # htmlName = fileName + ".html"
        htmlFullName = os.path.join(unicode(os.getcwd(), "gbk"), output)
        # htmlFullName = unicode(path, "gbk") + "\\" + htmlName
        print "generate html:" + htmlFullName
        doc.SaveAs(htmlFullName, 10)   #将word转存为html
        doc.Close()

    word.Quit()
    print ""
    print "Finished!"


if __name__ == '__main__':
    # import sys
    #
    # if len(sys.argv) != 2:
    #     print "Usage: python funcName.py rootdir"
    #     sys.exit(100)
    # wordsToHtml(sys.argv[1])
    parser=OptionParser(usage='%prog [options]')
    parser.add_option('-i','--in',dest='input',help='input file')
    parser.add_option('-o','--out',dest='output',help='output file')
    (options,args)=parser.parse_args()
    input=options.input
    output=options.output
    wordsToHtml(input,output)