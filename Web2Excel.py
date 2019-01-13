import pandas as pd
import sys


def Convert2XL(URLofSiteWithTable, xlPath='ConvertedFile.xlsx'):
    '''
    Downloads the first table from the URL provided
    and saves it an XLSX file.
    '''
    try:
        df = pd.read_html(URLofSiteWithTable)[0]
        df = df.fillna('')
        row_count, column_count = (df.shape)
        print('Rows found = %s, Columns found = %s' %
              (row_count, column_count))
        print(df.head())
        SavetoExcel(df, xlPath)

    except Exception as e:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        lineNo = str(exc_tb.tb_lineno)
        print('Error : %s : %s at Line %s.' % (type(e), e, lineNo))


def SavetoExcel(df, xlPath='ConvertedFile.xlsx'):
    try:
        writer = pd.ExcelWriter(xlPath, engine='xlsxwriter',
                                date_format='mm/dd/yyy', datetime_format='mm/dd/yyyy')
        df.to_excel(writer, index=False)
        writer.save()
        print('File saved at %s' % xlPath)

    except Exception as e:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        lineNo = str(exc_tb.tb_lineno)
        print('Error : %s : %s at Line %s.' % (type(e), e, lineNo))


if __name__ == '__main__':
    Convert2XL('https://www.bls.gov/cew/cewedr10.htm')
