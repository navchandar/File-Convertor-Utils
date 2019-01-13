#! python3
# -*- coding: utf-8 -*-
"""
Converts/Joins single/multiple images into 1 PDF file using PIL
IMG2PDF.convertImages(folder_path, pdf_filename) converts all images in a folder
IMG2PDF.convertImageList(image_list, pdf_filename) converts one/more images
"""
import os
import glob
from PIL import Image


def convertImages(folder_path, pdf_filename='Pdf_File.pdf'):
    '''
    Input should be a folder path with single or multiple images.
    This will automatically get all JPG/JPEG/PNG files in the folder.
    Output will be a single PDF file.
    '''
    CurrentDir = os.getcwd()
    os.chdir(folder_path)
    jpg_list = glob.glob("*.jpg")
    jpeg_list = glob.glob("*.jpeg")
    png_list = glob.glob("*.png")
    ImageList = jpg_list + jpeg_list + png_list
    if not ImageList:
        print('No image files found to convert')
    else:
        print('Found %s image files to convert' % len(ImageList))
        makePdf(ImageList, pdf_filename)

    os.chdir(CurrentDir)


def convertImageList(ImageList, pdf_filename='Pdf_File.pdf'):
    '''
    Input should be a LIST of Image filenames.
    Output will be a single PDF file
    '''
    if isinstance(ImageList, list):
        if not ImageList:
            print('No image files found to convert')
        else:
            print('Found %s image files to convert' % len(ImageList))
            makePdf(ImageList, pdf_filename)
    else:
        print('Input %s is not a list' % ImageList)


def makePdf(ImageList, pdf_filename):
    '''
    Generated PDF files from a list of image files
    '''
    im_list = []
    for ImageFile in ImageList:
        im_list.append(Image.open(ImageFile))

    if len(im_list) == 1:
        try:
            im_list[0].save(pdf_filename, "PDF", resolution=100.0, quality=100)
            print('%s is saved as %s' % (ImageFile, pdf_filename))
        except Exception as e:
            print('Error converting Images to PDF : %s ' % e)

    elif len(im_list) > 1:
        try:
            im1 = im_list[0]
            im_list.remove(im1)
            im1.save(pdf_filename, "PDF", resolution=100.0, quality=100,
                     save_all=True, append_images=im_list)
            print('%s is saved as  %s' % (ImageList, pdf_filename))
        except Exception as e:
            print('Error converting Images to PDF : %s ' % e)

