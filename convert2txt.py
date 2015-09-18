# -*- coding: utf-8 -*-

import os
import errno
from time import sleep
import sys
import subprocess
from subprocess import Popen, PIPE
import shutil
import csv
import xlrd
import docx
import pdfminer
import slate


from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.pdfpage import PDFPage
from pdfminer.converter import XMLConverter, HTMLConverter, TextConverter
from pdfminer.layout import LAParams
from cStringIO import StringIO


def convert_pdf_to_txt(path):
    rsrcmgr = PDFResourceManager()
    retstr = StringIO()
    codec = 'utf-8'
    laparams = LAParams()
    device = TextConverter(rsrcmgr, retstr, codec=codec, laparams=laparams)
    fp = file(path, 'rb')
    interpreter = PDFPageInterpreter(rsrcmgr, device)
    password = ""
    maxpages = 0
    caching = True
    pagenos=set()
    for page in PDFPage.get_pages(fp, pagenos, maxpages=maxpages, password=password,caching=caching, check_extractable=True):
        interpreter.process_page(page)
    fp.close()
    device.close()
    str = retstr.getvalue()
    retstr.close()
    return str


def doc2txt(filename):
    document = docx.Document(filename)
    docText = '\n'.join([
        paragraph.text.encode('utf-8') for paragraph in document.paragraphs
    ])
    return docText


def xls_to_csv(src, dst):
    x =  xlrd.open_workbook(src)
    sheets = x.sheet_names()
    for sheet in sheets:
        savefile = os.path.splitext(dst)[0] + ' - ' + sheet + '.txt'
        print savefile
        x1 = x.sheet_by_name(sheet)
        with open(savefile, 'wb') as out:
            writecsv = csv.writer(out, quoting=csv.QUOTE_ALL)
            for rownum in xrange(x1.nrows):
                row = [unicode(i).encode( 'utf-8') for i in x1.row_values(rownum)]
                print rownum, row
                out.write(' '.join(row) + '\n')


def mkdir_p(path):
    try:
        os.makedirs(path)
    except OSError as exc: # Python >2.5
        if exc.errno == errno.EEXIST and os.path.isdir(path):
            pass
        else: raise


rootDir = '/home/sim/git/CMA/root/Tender1'
convertedDir = os.path.abspath(rootDir) + '-converted'


failed_conversion = []

if os.path.exists(convertedDir):
    shutil.rmtree(convertedDir)

for path, dirs, files in os.walk(rootDir):
    for f in files:
        src = os.path.join(path, f)
        dst = src.replace(rootDir, convertedDir)
        dst = os.path.splitext(dst)[0] + '.txt'
        filetype = os.path.splitext(src)[1].lower()
        if filetype == '.pdf':
            print src, dst
            mkdir_p(os.path.dirname(dst))
            with open(dst, 'wb') as out:
                out.write(convert_pdf_to_txt(src))
        if filetype == '.docx':
            print src, dst
            with open(dst, 'wb') as out:
                out.write(doc2txt(src))
        if filetype in ['.xls', '.xlsx']:
            print src, dst
            try:
                xls_to_csv(src, dst)
            except:
                failed_conversion.append(src)

print 'Failed\n', failed_conversion

failedFile = os.path.join(rootDir, '../failedConversions.txt')
with open(failedFile, 'wb') as out:
    for item in failed_conversion:
        out.write(item)