# -*- coding: utf-8 -*-

import os
import sys
import errno
import shutil
import csv
import re
from subprocess import Popen, PIPE

import xlrd
import docx

from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.pdfpage import PDFPage
from pdfminer.converter import XMLConverter, HTMLConverter, TextConverter
from pdfminer.layout import LAParams
from cStringIO import StringIO


def pdf_to_txt(path):
    rsrcmgr = PDFResourceManager()
    retstr = StringIO()
    codec = 'utf-8'
    #laparams = LAParams()
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


def docx_to_txt(filename):
    document = docx.Document(filename)
    docText = '\n'.join([
        paragraph.text.encode('utf-8') for paragraph in document.paragraphs
    ])
    return docText


def xls_to_txt(src, dst):
    x = xlrd.open_workbook(src)
    sheets = x.sheet_names()
    text = ''
    for sheet in sheets:
        #savefile = os.path.splitext(dst)[0] + ' - ' + sheet + '.txt'
        x1 = x.sheet_by_name(sheet)
        #with open(savefile, 'wb') as out:
            # writecsv = csv.writer(out, quoting=csv.QUOTE_ALL) # For csv format
        for rownum in xrange(x1.nrows):
            row = [unicode(i).encode('utf-8') for i in x1.row_values(rownum)]
            text += ' '.join(row) + '\n'
            #out.write(' '.join(row) + '\n')
    return text


def mkdir_p(path):
    try:
        os.makedirs(path)
    except OSError as exc: # Python >2.5
        if exc.errno == errno.EEXIST and os.path.isdir(path):
            pass
        else: raise


def getUniqueName(path):
    i = 2
    newpath = path
    while os.path.exists(newpath):
        if os.path.isfile(path):
            filename, ext = os.path.splitext(path)
            newpath = getUniqueName(filename) + str(i) + ext
        else:
            newpath = path + str(i)
        i += 1
    return newpath


def main(srcDir):

    convertedDir = os.path.abspath(srcDir) + '-converted'
    dstDir = getUniqueName(convertedDir)

    failed_conversion = []

    for path, dirs, files in os.walk(srcDir):
        for f in files:
            src = os.path.join(path, f)
            copydst = src.replace(srcDir, dstDir)
            mkdir_p(os.path.dirname(copydst))
            shutil.copyfile(src, src.replace(srcDir, dstDir))
            slash = os.path.sep
            found = re.search('.*[Bb]uyer([0-9]+)'+slash+'([Tt]ender[0-9]+)'+slash+'([A-z0-9]+)', src)
            combinedFile = None
            if found:
                bxtenderx = 'b' + found.group(1) + found.group(2).lower()
                bxtenderxappend = bxtenderx + '-' + found.group(3).lower()
                saveDir = os.path.join(found.group(0), bxtenderxappend)
                convertDst = os.path.join(saveDir, f)
                combinedFile = os.path.join(saveDir.replace(srcDir, dstDir), bxtenderxappend) + '.txt'
            else:
                convertDst = src

            convertDst = convertDst.replace(srcDir, dstDir)
            convertDst = convertDst + '.txt'
            if os.path.isfile(convertDst):
                convertDst = getUniqueName(convertDst)

            text = None
            try:
                mkdir_p(os.path.dirname(convertDst))
                filetype = os.path.splitext(src)[1].lower()
                if filetype == '.doc':
                    print 'Converting:', src
                    cmd = ['antiword', src]
                    p = Popen(cmd, stdout=PIPE)
                    stdout, stderr = p.communicate()
                    text = stdout.decode('ascii', 'ignore')
                    text = unicode(text).encode('utf-8')
                    if p.returncode != 0:
                        raise
                if filetype == '.pdf':
                    print 'Converting:', src
                    text = pdf_to_txt(src)
                if filetype == '.docx':
                    print 'Converting:', src
                    text = docx_to_txt(src)
                if filetype in ['.xls', '.xlsx']:
                    print 'Converting:', src
                    text = xls_to_txt(src, convertDst)
                if filetype in ['.txt', '.text', '.csv']:
                    with open(src, 'rb') as input:
                        text = input.read()
                        convertDst = convertDst.strip(filetype) + '.txt'
                if text:
                    with open(convertDst, 'ab') as out:
                        out.write(text)
                    if combinedFile:
                        with open(combinedFile, 'ab') as combout:
                            combout.write(text)
            except:
                failed_conversion.append(src)


    if failed_conversion:
        print '\nFailed to convert:\n\n', '\n'.join(failed_conversion)

        failedFilename = os.path.basename(srcDir) + '-failed_to_convert.txt'
        failedFilepath = dstDir + '-failed_to_convert.txt'
        with open(failedFilepath, 'wb') as out:
            out.write('Failed to convert:\n\n' + '\n'.join(failed_conversion))


if __name__ == '__main__':
    args = sys.argv

    if len(args) < 2:  # need script
        print 'Provide path to directory to convert.'
        sys.exit(1)

    path = args[1]

    #path = '/home/sim/git/convert2txt/files/examples/test'

    main(path)
