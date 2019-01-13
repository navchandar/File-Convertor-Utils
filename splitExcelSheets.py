import pandas as pd
import os


def SplitXlSheets(xlFile):
    xl = pd.ExcelFile(xlFile)
    Sheets = xl.sheet_names  # see all sheet names
    for Sheet in Sheets:
        df = xl.parse(Sheet)
        FileName = os.path.dirname(xlFile) + '\\' + Sheet + '.xlsx'
        print('Creating %s' % FileName)
        SavetoExcel(df, FileName)


def SavetoExcel(df, xlPath):
    try:
        writer = pd.ExcelWriter(xlPath, engine='xlsxwriter',
                                date_format='mm/dd/yyy',
                                datetime_format='mm/dd/yyyy')
        df.to_excel(writer, index=False)
        # Close the Pandas Excel writer and output the Excel file.
        writer.save()

    except Exception as e:
        print('Error: %s.' % e)


if __name__ == '__main__':
    SplitXlSheets(input('Enter Excel file path to split into sep sheets: '))
