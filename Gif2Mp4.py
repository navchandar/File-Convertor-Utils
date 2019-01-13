#! python3
# -*- coding: utf-8 -*-
"""
Converts single/multiple gif files into mp4 files using FFMPEG
Gif2Mp4.convertGifs(foldername) converts all gif files in a folder to mp4
Gif2Mp4.convertGif(GifFilename) converts one gif file to mp4
"""
import os
import subprocess
from subprocess import PIPE


def convertGif(GifFileName):
    VidFileName = GifFileName + ".mp4"
    command = "ffmpeg -f gif -i " + GifFileName + " -pix_fmt yuv420p -c:v libx264 -movflags +faststart -filter:v crop='floor(in_w/2)*2:floor(in_h/2)*2' " + VidFileName

    p = subprocess.Popen(command, shell=True,
                         stdout=PIPE, stderr=PIPE)
    stdout, stderr = p.communicate()

    if stdout:
        print(stdout.decode('ascii'))
    if stderr:
        print(stderr.decode('ascii'))


def convertGifs(folderPath):
    os.chdir(folderPath)
    for files in os.listdir(folderPath):
        if files.endswith('.gif'):
            print('Converting : %s' % files)
            convertGif(files)
