#! python3
# -*- coding: utf-8 -*-
"""
Extracts Txt content from PDF files with images/scans
Needs Tesseract and ImageMagick installations to work

PDF2TXT.convertPdfs(folder_path, Optional Txt filename) all PDFs in a folder
PDF2TXT.convertPDFList(PDFfiles_list, Optional Txt filename) single to multiple PDFs
"""
import os
import glob
import subprocess


def convertPdfs(folder_path, txt_filename):
    '''
    Input should be a folder path with single or multiple pdfs.
    This will automatically get all PDF files in the folder.
    Output will be a single Txt file.
    '''
    CurrentDir = os.getcwd()
    os.chdir(folder_path)
    pdf_List = glob.glob("*.pdf")
    convertPDFList(pdf_List, txt_filename)
    os.chdir(CurrentDir)


def convertPDFList(pdf_List, txt_filename):
    '''
    Input should be a LIST of PDF filenames.
    Output will be a single TXT file
    '''
    if isinstance(pdf_List, list):
        if not pdf_List:
            print('No pdf files found to convert')
        else:
            print('Found %s pdf files to convert' % len(pdf_List))
            maketxt(pdf_List, txt_filename)
    else:
        print('Input %s is not a list' % pdf_List)


def maketxt(pdf_List, txt_filename):
    # convert pdf into set of png pages with imagemagick
    for index, pdfFile in enumerate(pdf_List):
        try:
            ImageFileName = str(index + 1) + '.png'
            command = 'convert -density %s %s %s' % ('400', pdfFile, ImageFileName)
            subprocess.call(command, shell=True)
        except Exception as e:
            print('Error converting PDF to Image : %s ' % e)

    ImageList = glob.glob('*.png')
    if not ImageList:
        print('No png files found to convert to txt')
    else:
        print('Found %s png files to convert' % len(pdf_List))
        try:
            # extract text into text file
            for ImageFileName in ImageList:
                subprocess.call("tesseract %s stdout > text.txt " % ImageFileName, shell=True)
                textFromImage = open('text.txt', 'r').read()
                with open(txt_filename, 'a') as f:
                    f.write(textFromImage)

                # delete the images and unnecessary txt files
                os.remove(ImageFileName)
                os.remove('text.txt')
        except Exception as e:
            print('Error converting Image to Txt : %s ' % e)

