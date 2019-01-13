#! python3
# -*- coding: utf-8 -*-
"""
XML Data Generator - Generates XML files from Excel DataSheet and a sample XML file
"""
import os
import sys
import copy
import time
import pandas as pd
from lxml import etree
from datetime import datetime


def issues_create(dict_issues, ListOfLists, ColumnList, xml_template_path):
    # Parser used to keep written XML human readable
    parser = etree.XMLParser(remove_blank_text=True)
    for issue in dict_issues:
        issue_key = issue
        Row = ListOfLists[issue - 1]
        if len(Row) != len(ColumnList):
            print('Error: No. of Rows %s and Column Names %s Mismatch!' % (len(Row), len(ColumnList)))
        else:
            try:
                xml_template = open(xml_template_path, 'r')
                filename = "Test Data-" + str(issue_key) + ".xml"
                xml_issue_path = (filename)
                if os.path.isfile(xml_issue_path) is False:
                    print('Creating XML : %s' % filename)
                    # Create / open XML for writing issue details
                    xml_issue = open(xml_issue_path, 'wb')
                    # Get copy of template 000.xml and write out copy of.
                    baseline_tree = etree.parse(xml_template, parser)
                    # Edit Tree after copy is made
                    new_tree = xml_issues_update(baseline_tree, issue, Row, ColumnList)
                    # Write out tree
                    new_tree.write(xml_issue, pretty_print=True,
                                   xml_declaration=True, encoding="UTF-8")
                    xml_issue.close()

                    print('XML created    : %s' % filename)

            except IndexError as i:
                print('IndexError : ' + str(i))
                pass

            except Exception as e:
                exc_type, exc_obj, exc_tb = sys.exc_info()
                lineNo = str(exc_tb.tb_lineno)
                print('Error while creating XMLs : %s : %s at Line %s.' % (type(e), e, lineNo))

    return


def xml_issues_update(baseline_tree, issue, ListOfValues, ColumnList):
    new_tree = copy.deepcopy(baseline_tree)

    # Below is used to specific what Elements to check in the XML. field[0]
    # denotes the Element and Field[1] is the unique data to populate.
    ls_xml_fields = []
    for index, column in enumerate(ColumnList):
        ls_xml_fields.append([str(column), str(ListOfValues[index])])

    for field in ls_xml_fields:
        # Try wraps around all attempts to check if field exists and 'do' and action.
        try:
            tag = field[0]
            Element = new_tree.find('.//' + tag)
            if Element is not None:
                Element.text = field[1]
                print('Updated the Element : %s with %s' % (tag, field[1]))
            else:
                tagAndAttrib = tag.split('.')
                if tagAndAttrib != []:
                    Element = new_tree.find('.//' + tagAndAttrib[0])
                    if Element is not None:
                        attributes = Element.attrib
                        attributes[tagAndAttrib[1]] = field[1]
                        print('Updated the Attribute : %s with %s' % (tag, field[1]))
                    else:
                        print('%s doesn\'t have the Element : %s' % (xml_filename, tag))

        except (AttributeError, KeyError) as e:
            print('last except', ":", e)
            continue

    # Return the XML to be written
    return new_tree
# -------------------------------------------------------------------------


def folderCreate(FolderName):
    try:
        if not os.path.isdir(FolderName):
            os.mkdir(FolderName)
            print('Folder creation successful : %s' % FolderName)
        else:
            print('Folder already exists : %s' % FolderName)

        os.chdir(FolderName)
    except Exception as e:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        lineNo = str(exc_tb.tb_lineno)
        print('Error: %s : %s at Line %s.' % (type(e), e, lineNo))


def folderDelete(FolderName):
    try:
        if os.path.isdir(FolderName):
            if os.listdir(FolderName) == []:
                os.rmdir(FolderName)
                print('Folder deletion successful : %s' % FolderName)
            else:
                print('Folder not empty and not Deleted : %s' % FolderName)
        else:
            print("Folder doesn't exist : %s" % FolderName)

    except Exception as e:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        lineNo = str(exc_tb.tb_lineno)
        print('Error: %s : %s at Line %s.' % (type(e), e, lineNo))


def emptyFilesDelete(path):
    fileList = os.listdir(path)
    for files in fileList:
        filePath = os.path.join(path, files)
        if os.path.isfile(filePath):
            if os.path.getsize(filePath) == 0:
                try:
                    os.remove(filePath)
                    print('Empty File Deletion successful : %s' % files)
                except Exception as e:
                    exc_type, exc_obj, exc_tb = sys.exc_info()
                    lineNo = str(exc_tb.tb_lineno)
                    print('Error deleting file: %s : %s at Line %s.' % (type(e), e, lineNo))

        elif os.path.isdir(filePath):
            emptyFilesDelete(filePath)


def ReadExcelData(xl_path):
    try:
        df = pd.read_excel(xl_path, sheet_name=0, index_col=None, dtype=str)  # 0 Reads first Sheet
        df = df.fillna('')
        # print(df.head())
        row_count, column_count = (df.shape)
        print('Rows = %s, Columns = %s' % (row_count, column_count))

        ColumnNames = df.columns.values.tolist()

        # Convert Data in all Rows into a list of lists
        AllRowValueList = df.values.tolist()

        # Convert Data in all columns into a dict of lists
        AllColValueList = dict()
        for column in ColumnNames:
            _ls = df[column].tolist()
            AllColValueList[column] = _ls
        print('Excel to List conversion successful.')
        # print(ColumnNames)
        # print(AllRowValueList)
        return (AllRowValueList, ColumnNames, row_count, True)

    except Exception as e:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        lineNo = str(exc_tb.tb_lineno)
        print('Error: %s : %s at Line %s.' % (type(e), e, lineNo))
        return (None, 0, 0, False)


def convert(xml_filename, xl_path):
    xml_template_dir = os.path.dirname(xml_filename)
    os.chdir(xml_template_dir)
    initalTime = time.time()

    if os.path.isfile(xml_filename):
        if os.path.isfile(xl_path):
            print('XML and XL Files found')
            AllRowsList, ColList, RowCount, success = ReadExcelData(xl_path)
            if success is True:
                    if RowCount > 0:
                        ColumnList = []
                        for Col in ColList:
                            ColumnList.append(Col.replace(" ", ""))

                        folderdate = datetime.now().strftime('%m%d%y %M%S%p')
                        DataFolderName = 'DataGen ' + str(folderdate)
                        folderCreate(DataFolderName)

                        issues_create(range(1, RowCount + 1), AllRowsList, ColumnList, xml_filename)

                        # Go back to the original folder
                        os.chdir(xml_template_dir)

                        # Clean Up
                        emptyFilesDelete(DataFolderName)
                        folderDelete(DataFolderName)

                        validFileCount = len(os.listdir(DataFolderName))
                        if validFileCount >= 1:
                            os.startfile(DataFolderName)
                        print('Test Data Generation Complete.')

                        finalTime = time.time()
                        TotalTime = "{0:.2f}".format(finalTime - initalTime)
                        XMLsPerSec = "{0:.2f}".format(validFileCount / float(TotalTime))

                        Time_Saved = (validFileCount * 60) - float(TotalTime)
                        Amount_Saved = "{0:.2f}".format(float((Time_Saved / (60 * 60)) * 20.89))

                        Savings_Text = ''
                        if Time_Saved <= 60:
                            Savings_Text = ' %s seconds and %s $.' % ("{0:.2f}".format(float(Time_Saved)), Amount_Saved)
                        elif 60 < Time_Saved <= 3600:
                            Savings_Text = ' %s minutes and %s $.' % ("{0:.2f}".format(float(Time_Saved / 60)), Amount_Saved)
                        elif Time_Saved > 3600:
                            Savings_Text = ' %s hours and %s $.' % ("{0:.2f}".format(float(Time_Saved / (60 * 60))), Amount_Saved)

                        print("\nTime taken to create %s XMLs: %s seconds." % (validFileCount, TotalTime))
                        print("Approx No. of XMLs created per second : %s " % XMLsPerSec)
                        print("\nAssuming 1min/XML for manual data creation, Savings generated by this script run : %s" % Savings_Text)

                    else:
                        print("No Excel Data found. No of rows = %s" % RowCount)
            else:
                print("Unknown Error: Unable to Read Excel Data.")
        else:
            print('File Not found at %s' % xl_path)
    else:
        print('File Not found at %s' % xml_filename)
