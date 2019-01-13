#! python3
# -*- coding: utf-8 -*-
"""
XML Converter - Converts XML file to Excel data sheet based on tags required
"""
import os
import sys
import time
import logging
import openpyxl
import subprocess
import pandas as pd
from datetime import datetime
import xml.etree.ElementTree as ET


if xml_filename.endswith(('.xml', '.XML', '.txt', '.TXT')):
    if os.path.exists(xml_filename):
        print("XML file selection success")
        os.chdir(os.path.dirname(xml_filename))
        xl_path = "Export-" + datetime.now().strftime('%m%d%y-%H%M%S%f') + ".xlsx"

        try:
            Prog_start_time = time.time()
            status = ProcessXML(xml_filename)

        except Exception as e:
            exc_type, exc_obj, exc_tb = sys.exc_info()
            lineNo = str(exc_tb.tb_lineno)
            print('Error: %s : %s at Line %s.' % (type(e), e, lineNo))

        Prog_run_time = (time.time() - Prog_start_time)
        Prog_run_time = "{0:.2f}".format(Prog_run_time)
        print("Script run time = %s seconds" % Prog_run_time)


def ParseXml(XmlFile):
    '''
    Parse any XML File or XML content and get the output as a list of
    (ListOfTags, ListOfValues, ListOfAttribs)
    Tags will only have the sub child Tag names and not the parent names.
    Attributes will be list of dictionary items.
    '''
    try:
        if os.path.exists(XmlFile):
            tree = ET.parse(XmlFile)
            root = tree.getroot()
            logging.inf('Trying to Parse %s' % XmlFile)
        else:
            root = ET.fromstring(XmlFile)

        # Use the below to print & see if XML is actually loaded
        # print(etree.tostring(tree, pretty_print=True))
        ListOfTags, ListOfValues, ListOfAttribs = [], [], []
        for elem in root.iter('*'):
            Tag = elem.tag
            if ('}' in Tag):
                Tag = Tag.split('}')[1]
            ListOfTags.append(Tag)

            value = elem.text
            if value is not None:
                ListOfValues.append(value)
            else:
                ListOfValues.append('')

            attrib = elem.attrib
            if attrib:
                ListOfAttribs.append([attrib])
            else:
                ListOfAttribs.append([])

        # print(ListOfTags, ListOfValues, ListOfAttribs)
        print('XML File content parsed successfully')
        print('XML File content parsed successfully')

        return (ListOfTags, ListOfValues, ListOfAttribs)

    except Exception as e:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        lineNo = str(exc_tb.tb_lineno)
        print('Error while parsing XMLs : %s : %s at Line %s.' % (type(e), e, lineNo))
        return ([], [], [])


def GetAttributeValues(ListOfTags, ListOfValues, ListOfAttribs):
    try:
        ListOfAttrs = []
        for index, Attributes in enumerate(ListOfAttribs):
            if Attributes:
                if type(Attributes) == list:
                    Dictionary = Attributes[0]
                elif type(Attributes) == dict:
                    Dictionary = Attributes

        ListOfAtriValues = []
        for index, Attributes in enumerate(ListOfAttribs):
            if Attributes:
                if type(Attributes) == list:
                    Dictionary = Attributes[0]
                elif type(Attributes) == dict:
                    Dictionary = Attributes
                DictLength = len(Dictionary)

                if DictLength == 1:
                    for key, value in Dictionary.items():
                        # To separate the remaining Keys and Values into two columns
                        ListOfAttrs.append(key)
                        ListOfAtriValues.append([value])

                elif DictLength > 1:
                    for key, value in Dictionary.items():
                        # To separate the remaining Keys and Values into two columns
                        ListOfAttrs.append(key)
                        ListOfAtriValues.append([value])
                    for i in range(len(Dictionary) - 1):
                        ListOfTags[index:index] = [ListOfTags[index]]
                        ListOfValues[index:index] = [ListOfValues[index]]

            else:
                ListOfAtriValues.append([])
                ListOfAttrs.append('')
        print('Attributes and values successfully extracted.')
        print('Attributes and values successfully extracted.')

        while len(ListOfAttrs) < len(ListOfTags):
            ListOfAttrs.append('')
        while len(ListOfAtriValues) < len(ListOfValues):
            ListOfAtriValues.append([])

        return ListOfTags, ListOfValues, ListOfAttrs, ListOfAtriValues

    except Exception as e:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        lineNo = str(exc_tb.tb_lineno)
        print('Error while extracting Attribute values : %s : %s at Line %s.' % (type(e), e, lineNo))
        return []


def exportExcel(dictionary, columnsList, xl_path):

    df = pd.DataFrame.from_dict(dictionary, orient='columns')[columnsList]
    # convert to excel file
    writer = pd.ExcelWriter(xl_path, engine='xlsxwriter')
    df.to_excel(writer, index=False)
    writer.save()
    writer.close()
    print('Excel Exported successfully at %s' % xl_path)


def makeExcel(ListOfTags, ListOfValues, ListOfAttribs, ListOfAtriValues):
    xl_path = "Export-" + datetime.now().strftime('%m%d%y%H%M%S%f') + ".xlsx"

    # Make each column separately
    df1 = pd.DataFrame(ListOfTags, columns=['XML Tag Names'])
    df2 = pd.DataFrame(ListOfValues, columns=['XML Tag Values'])
    df3 = pd.DataFrame(ListOfAttribs, columns=['Attribute Names'])
    df4 = pd.DataFrame(ListOfAtriValues, columns=['Attribute Values'])
    # concatenate all the columns into one DataFrame
    df = pd.concat([df1, df2, df3, df4], axis=1)

    # convert to excel file
    writer = pd.ExcelWriter(xl_path, engine='xlsxwriter')
    df.to_excel(writer, index=False)
    writer.save()
    writer.close()
    print('Excel Exported successfully at %s' % xl_path)


def ProcessXML(xml_filename):
    status = 1
    try:
        values = {}
        lineCount = 0
        with open(xml_filename) as f:
            for line in f:
                # print(line)
                lineCount += 1

                ListOfTags, ListOfValues, ListOfAttribs = ParseXml(line)

                TagCount, ValueCount = len(ListOfTags), len(ListOfValues)

                if (TagCount == ValueCount):
                    print('Count of Tags and Values retrieved : %s ' % TagCount)
                    for tag in RequiredTags[1:]:    # Excluding 1st so as to add line count once
                        if tag in ListOfTags:
                            value = ListOfValues[ListOfTags.index(tag)]
                            if value:
                                values.setdefault(tag, []).append(value)
                            else:
                                values.setdefault(tag, []).append('NILL')
                        else:
                            values.setdefault(tag, []).append('')
                else:
                    status = 0
                    print('Error: Count mismatch in Tags and Values retrieved.\n\TagCount = %s. ValueCount = %s.'
                          % (TagCount, ValueCount))
                    logging.error('Error: Count mismatch. TagCount = %s. ValueCount = %s.' % (TagCount, ValueCount))

        for k, v in values.items():
            print(k, len(v))
            print('%s Count = %s ' % (k, len(v)))

        exportExcel(values, RequiredTags, xl_path)

    except Exception as e:
        status = 0
        exc_type, exc_obj, exc_tb = sys.exc_info()
        lineNo = str(exc_tb.tb_lineno)
        print('Error: %s : %s at Line %s.' % (type(e), e, lineNo))

    return status
