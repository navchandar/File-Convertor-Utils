#! python3
# -*- coding: utf-8 -*-
"""
PDF2string  extracts string content from PDF files even from locked PDFs
PDF2str.convertPdfs(folder_path, Optional PDF pwd, Optional Txt filename) all PDFs in a folder
PDF2str.convertPDFList(pdf_List, Optional PDF pwd, Optional Txt filename) single to multiple PDFs
"""
import PyPDF2
import glob
import os


def convertPDFs(folder_path, pwd='', txt_filename):
    '''
    Input should be a folder path with single or multiple pdfs.
    This will automatically get all pdf files in the folder.
    Output will be a single txt file.
    '''
    CurrentDir = os.getcwd()
    os.chdir(folder_path)
    pdf_List = glob.glob("*.pdf")
    if not pdf_List:
        print('No pdf files found to convert')
    else:
        print('Found %s pdf files to convert' % len(pdf_List))
        convertPDFList(pdf_List, pwd, txt_filename)

    os.chdir(CurrentDir)


def convertPDFList(pdf_List, pwd='', txt_filename):
    '''
    Input should be a pdf list with one or multiple PDFs.
    Output will be a single txt file.
    '''
    for pdfFile in pdf_List:
        if not os.path.exists(pdfFile):
            print('%s doesnt exist' % pdfFile)

        elif not pdfFile.endswith(('.pdf', '.PDF')):
            print('%s is not a PDF file' % pdfFile)

        else:
            pdfFileObj = open(pdfFile, 'rb')
            pdfReader = PyPDF2.PdfFileReader(pdfFileObj)

            try:
                if pdfReader.isEncrypted:
                    if not pwd:
                        print('%s is locked and no password entered' % pdfFile)
                    else:
                        pdfReader.decrypt(pwd)
            except Exception as e:
                print('Error while decrypting %s : %s ' % (pdfFile, e))

            if not pdfReader:
                print('Error while reading the file : %s ' % pdfFile)
            else:
                try:
                    pages = pdfReader.numPages
                    print('%s pages in %s' % (pages, pdfFile))

                    textFromPDF = ''
                    for pageNum in range(pages):
                        pageObj = pdfReader.getPage(pageNum)
                        text = pageObj.extractText()
                        textFromPDF += text

                    if not textFromPDF:
                        print('No Txt found in  %s. You may want to use PDF2TXT.convertPdfs()' % pdfFile)
                    else:
                        with open(txt_filename, 'a') as f:
                            f.write(textFromPDF)
                        print('Txt saved in %s.' % txt_filename)

                except Exception as e:
                    print('Error while converting %s : %s ' % (pdfFile, e))

