# -*- coding: utf-8 -*-
"""
Created on Fri Sep 21 16:05:38 2018
@author: phlip (ppp)
2018/10/08: ppp; use PyInstaller to compile the python program
            pyinstaller --onefile <your_script_name>.py
2018/10/09: ppp; Add formatting of header line: bold and underline
2018/10/11: ppp; Add log file of runs in json format with dict
          : ppp; add program version
          : ppp; read tag clean up tags from v_striplist in tagCleanup
2018/10/24: ppp; fixed while v_rowStartLineNum < v_rowEndLineNum:
          :      to while v_rowStartLineNum <= v_rowEndLineNum:
                 as the '<' lost the last field in the XML file
2018/11/04: PPP; v2 - Change to xlsx format
2018/11/10: PPP; v3 - Add configuration information
"""

import os
import datetime
import json
import xlsxwriter

import toml
# from win32com.client import Dispatch

__version__ = 'v00.00.00.03'

v_numRows = 0
v_rowStartPositionList = []
v_rowEndPositionList = []
v_headingList = []
v_tagList = []
v_contentList = []
v_histDict = {}
v_histDict['report'] = []
v_configExists = True

v_stripList = ['<', '>', '/']

# check if the configuration file is in the same dir as excutable
try:
    v_config = toml.load('genrepconfig.toml')
except IOError:
    print("The configuration file does not exists. Please add to the same dir")
    v_configExists = False

v_curDir = os.getcwd()
if v_configExists:
    v_pathLoad = v_config['default_load_path']
    v_pathLoad = v_pathLoad['load_path_name']
    v_loadDir = v_pathLoad
    v_pathSave = v_config['default_save_path']
    v_pathSave = v_pathSave['save_path_name']
    v_saveDir = v_pathSave
    v_sameionames = v_config['iofilenames']
    v_sameionames = v_sameionames['same_names']
    v_historysave = v_config['history']
    v_historysave = v_historysave['hist_to_save_path']
else:
    v_loadDir = v_saveDir = v_curDir

# A list of tags that should be ignore (for now)
V_EXCLUDETAGS = ['<INSTITUTION_NAME>', '<CURDATE>', '<BATCH_REQUEST_NUMBER>',
                 '<REPORT_DESCRIPTION>', '<PROGRAM_NAME>',
                 '<PROGRAM_VERSION>', '<OUTPUT_URL>', '<OUTPUT_DAD>',
                 '<ROWSET>', '</ROWSET>', '<ROW>', '</ROW>',
                 '<?xml version="1.0" encoding="UTF-8"?>',
                 '<FROM_LETTER/>', '<FROM_LETTER_DESC/>',
                 '<TO_LETTER/>']


def tagCleanup(x_tag):
    """A procedure to remove certain characters from a tag"""
    for r_tag in v_stripList:
        x_tag = x_tag.strip(r_tag)
    return x_tag


def lFind(x_string, x_word, x_start=0):
    """A function to find a string in a larger string"""
    position = x_string.find(x_word, x_start)
    return position


def getConfig(x_confvalue):
    """Get the configurations from the genrepconfig.toml file"""
    global v_config
    v_configvalue = v_config[x_confvalue]
    return v_configvalue


v_totalLines = 0

# get the filename that must be used for the input file
v_fileIn = input("Type the name of the file that you want to read from: ")

# Check if the user has entered the xml extention; if not add it
v_extention = v_fileIn.find('.xml')
if v_extention == -1:
    v_fileIn = v_fileIn+'.xml'

# get the filename that must be used for the output file only if same_names!=Y
if v_sameionames != 'Y':
    v_fileOut = input("Type the name of the file that you want to save to: ")
else:
    v_fileOut = v_fileIn.strip('.xml')
# Check if the user has entered the xlsx extention; if not add it
v_extention = v_fileOut.find('.xlsx')
if v_extention == -1:
    v_fileOut = v_fileOut + '.xlsx'

with open(v_loadDir + '/' + v_fileIn) as file:
    v_data = file.readlines()

for v_line in v_data:
    v_posStart = v_line.find('<ROW>')
    v_posEnd = v_line.find('</ROW>')
    if v_posStart > -1:
        v_rowStartPositionList.append(v_totalLines)
    if v_posEnd > -1:
        v_numRows += 1
        v_rowEndPositionList.append(v_totalLines)
    v_totalLines += 1

v_listPosition = 0

while v_numRows > 0:
    # work through the data starting witht the first data ROW
    v_rowStartLineNum = v_rowStartPositionList[v_listPosition] + 1
    # until the end of the last ROW
    v_rowEndLineNum = v_rowEndPositionList[v_listPosition] - 1
    v_listPosition += 1
    while v_rowStartLineNum <= v_rowEndLineNum:
        v_word = v_data[v_rowStartLineNum]

        # Determining the start and end of each tag
        v_startTagStart = lFind(v_word, '<')
        v_startTagEnd = lFind(v_word, '>')
        v_endTagStart = lFind(v_word, '</')
        v_endTagEnd = lFind(v_word, '>', v_endTagStart)

        # get the tag text
        v_tagText = v_word[v_startTagStart: v_startTagEnd+1]

        # Skip tags that should not be in the report
        if v_tagText in V_EXCLUDETAGS:
            v_headingList.append(v_tagText)
            v_rowStartLineNum += 1
            continue
        # get the tag text and add to the tag text list if not yet in the list
        if v_tagText not in v_tagList:
            v_tagList.append(v_tagText)

        # get the tag content and add to the content list; always
        v_tagContent = v_word[v_startTagEnd + 1:v_endTagStart]
        v_contentList.append(v_tagContent)

        v_rowStartLineNum += 1

    # mark the last set of tags for the current ROW
    v_contentList.append('ROWEND')
    v_numRows -= 1


workbook = xlsxwriter.Workbook(v_saveDir + '/' + v_fileOut)

worksheet = workbook.add_worksheet()

t_row = 0
t_col = 0
now = datetime.datetime.now()

v_year = "Date: " + str(now.year)
v_month = "/" + str(now.month)
if now.day < 10:
    v_day = "/0" + str(now.day)
else:
    v_day = "/" + str(now.day)
v_date = v_year + v_month + v_day

worksheet.write(t_row, t_col, v_date)
t_row += 1

# Add a bold_underline format to use to highlight cells.
bold_underline = workbook.add_format({'bold': True, 'underline': True})

workbook.set_properties({
        'title': 'Data from ITS Forms Report option that is in XML',
        'subject': v_fileIn,
        'author': 'Phlip Pretorius',
        'manager': 'The one-and-only PPP',
        'company': 'Phlip Pretorius Consulting/Konsultasie',
        'category': 'ITS Data',
        'keywords': 'Sample, Example, Properties',
        'created': datetime.date(now.year, now.month, now.day),
        'comments': 'Created with Python and XlsxWriter',
        'hyperlink': 'Some URL'})

# write the heading row to the Excel spreadsheet
for tag in v_tagList:
    tag = tagCleanup(tag)
    # worksheet.write(t_row, t_col, tag, v_style)
    worksheet.write(t_row, t_col, tag, bold_underline)
    t_col += 1

row = 2
col = 0
for word in v_contentList:
    if word == 'ROWEND':
        row += 1
        col = 0
        continue
    worksheet.write(row, col, word)
    col += 1

workbook.close()

v_histDict['report'].append({
        '0_Program_version': __version__,
        '1_Date_Time': str(now),
        '2_InFile': v_fileIn,
        '3_RepeatingRecords': len(v_rowStartPositionList),
        '4_OutFile': v_fileOut,
        '5_Rows_to_OutFile': row,
        '6_Columns_to_OutFile': t_col
        })

if v_historysave == 'Y':
    V_HISTFNAME = v_pathSave + '/history.txt'
else:
    V_HISTFNAME = 'history.txt'

with open(V_HISTFNAME, 'a') as outfile:
    json.dump(v_histDict, outfile, sort_keys=True, indent=4)

print("\nCurrent date and time: " + str(now))
print("\nFile in", v_fileIn)
print("\tNumber of repeating records = ", len(v_rowStartPositionList))
print("\nFile out", v_fileOut)
print("\tNumber of rows written to the file = ", row)
print("\tNumber of columns written to the file = ", t_col)
