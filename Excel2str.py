#! python3
# -*- coding: utf-8 -*-
"""
Excel2str  extracts string content from Excel files
"""
import pandas as pd


def convertExcel(xl_filename, txt_filename):
    '''
    Input should be a single excel file.
    Output will be a single txt file with all content as string.
    '''
    xl = pd.ExcelFile(xl_filename)
    sheet_names = xl.sheet_names      # see all sheet names
    print('Found %s sheets in the file %s' % (sheet_names, xl_filename))
    for sheet in sheet_names:
        df = xl.parse(sheet, header=None, index_col=None)
        # print(df.head())
        df = df.fillna('')
        row_count, column_count = (df.shape)
        print('Rows = %s, Cols = %s in the sheet %s' % (row_count, column_count, sheet))
        all_lines = []
        for index, row in df.iterrows():
            line = []
            for i in range(column_count):
                line.append(row[i])
            all_lines.append(line)
        print('%s lines in %s' % (len(all_lines), xl_filename))

        if len(all_lines) > 0:
            with open(txt_filename, 'a') as f:
                for line in all_lines:
                    s = ', '.join(str(e) for e in line)
                    f.write('%s\n' % (s))
            print('%s created successfully' % txt_filename)


def convertBigExcel(xl_filename, txt_filename):
    '''
    Input should be a single excel file which is very big for processing.
    Output will be a single txt file with all content as string.
    '''
    xl = pd.ExcelFile(xl_filename)
    sheet_names = xl.sheet_names      # see all sheet names
    print('Found %s sheets in the file %s' % (sheet_names, xl_filename))
    with open(txt_filename, 'a') as f:
        for sheet in sheet_names:
            df = xl.parse(sheet, header=None, index_col=None)
            # print(df.head())
            df = df.fillna('')
            row_count, column_count = (df.shape)
            print('Rows = %s, Cols = %s in the sheet %s' % (row_count, column_count, sheet))
            for index, row in df.iterrows():
                line = []
                for i in range(column_count):
                    line.append(row[i])
                s = ', '.join(str(e) for e in line)
                f.write('%s\n' % (s))
        print('%s created successfully' % txt_filename)

