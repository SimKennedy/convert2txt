# -*- coding: utf-8 -*-

import os
import sys
import errno
import shutil
import csv
import xlrd
import docx
from subprocess import Popen, PIPE

from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.pdfpage import PDFPage
from pdfminer.converter import XMLConverter, HTMLConverter, TextConverter
from pdfminer.layout import LAParams
from cStringIO import StringIO


def convert_pdf_to_txt(path):
    rsrcmgr = PDFResourceManager()
    retstr = StringIO()
    codec = 'utf-8'
    laparams = None
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


def docx2txt(filename):
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
        x1 = x.sheet_by_name(sheet)
        with open(savefile, 'wb') as out:
            writecsv = csv.writer(out, quoting=csv.QUOTE_ALL)
            for rownum in xrange(x1.nrows):
                row = [unicode(i).encode('utf-8') for i in x1.row_values(rownum)]
                out.write(' '.join(row) + '\n')


def mkdir_p(path):
    try:
        os.makedirs(path)
    except OSError as exc: # Python >2.5
        if exc.errno == errno.EEXIST and os.path.isdir(path):
            pass
        else: raise


def main(srcDir):
    srcDir = '/home/sim/git/convert2txt/root/Tender1'
    #srcDir = '/home/sim/git/convert2txt/examples/test'
    convertedDir = os.path.abspath(srcDir) + '-converted'

    failed_conversion = []

    if os.path.exists(convertedDir):
        shutil.rmtree(convertedDir)

    for path, dirs, files in os.walk(srcDir):
        for f in files:
            src = os.path.join(path, f)
            try:
                dst = src.replace(srcDir, convertedDir)
                dst = os.path.splitext(dst)[0] + '.txt'
                mkdir_p(os.path.dirname(dst))
                filetype = os.path.splitext(src)[1].lower()
                if filetype == '.doc':
                    with open(dst, 'wb') as out:
                        cmd = ['antiword', src]
                        p = Popen(cmd, stdout=PIPE)
                        stdout, stderr = p.communicate()
                        text = stdout.decode('ascii', 'ignore')
                        text = unicode(text).encode('utf-8')
                        out.write(text)
                        if not os.path.isfile(dst):
                            raise
                if filetype == '.pdf':
                    print 'Converting:', src
                    with open(dst, 'wb') as out:
                        out.write(convert_pdf_to_txt(src))
                if filetype == '.docx':
                    print 'Converting:', src
                    with open(dst, 'wb') as out:
                        out.write(docx2txt(src))
                if filetype in ['.xls', '.xlsx']:
                    print 'Converting:', src
                    xls_to_csv(src, dst)
            except:
                failed_conversion.append(src)

    print failed_conversion
    if failed_conversion:
        print 'Failed:\n', '\n'.join(failed_conversion)
        failedFile = os.path.abspath(os.path.join(srcDir, os.pardir, 'failed_to_convert.txt'))

        with open(failedFile, 'wb') as out:
            for item in failed_conversion:
                out.write(item)

if __name__ == '__main__':
    args = sys.argv

    if len(args) < 2:  # need script
        print 'Provide path to directory to convert.'
        sys.exit(1)

    path = args[1]

    main(path)