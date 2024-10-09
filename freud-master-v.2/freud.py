#Script to Analyse DROID reports - Freud
#This script will analyse the DROID report to look for common issues which effect the ability to ingest into DRI. It can be run easily by running the batch file and then dragging the file which you want to analyse
#this will then create an excel spreadsheet in the directory which the program is running from with a different worksheet for each issue highlighted.

import pandas as pd
import numpy as np
import os

#below loads the csv file into pandas and takes required columns, additional columns for multiple identification are not taken yet as this breaks the csv read. It also loads a copy of the white list formats accepted into DRI
csvraw = input("Enter filepath of DROID csv to analyse: ")
csvraw = csvraw.strip('"')
columns_needed = ['ID', 'PARENT_ID', 'URI', 'FILE_PATH', 'NAME', 'METHOD', 'STATUS', 'SIZE', 'TYPE', 'EXT', 'LAST_MODIFIED', 'EXTENSION_MISMATCH', 'SHA256_HASH', 'FORMAT_COUNT', 'PUID', 'MIME_TYPE', 'FORMAT_NAME', 'FORMAT_VERSION']
csv = pd.read_csv(csvraw, usecols=columns_needed)
droidname = os.path.basename(csvraw)
droidname = droidname.rstrip('.csv')
results = pd.ExcelWriter(droidname+'_freudresults.xlsx', engine='xlsxwriter')
originalWhiteList = pd.read_csv('approved_formats.csv')
originalFlaggedFormats = pd.read_csv('Flagged_Formats.csv')
originalFurtherResearch = pd.read_csv('Further_Research.csv')


def unidentified(): #function run to add a worksheet which selects all files which have not been identified by DROID or show up as OLE 2 files, also adds a new title row and makes it blue

    unidentified = csv
    unidentified = unidentified.loc[(((unidentified['FORMAT_COUNT'] == 0) & (unidentified['SIZE'] > 0))) | (unidentified['PUID'] == 'fmt/111' ), :]
    unidentified = unidentified.sort_values('EXT')
    resultbook = results.book
    format = resultbook.add_format({
        'bold': True,
        'fg_color': '#4c9df7'})

    unidentified.to_excel(results, sheet_name='Unidentified_Formats',index=False, startcol = 0, startrow = 1)
    sheet1 = results.sheets['Unidentified_Formats']
    sheet1.write('A1', 'UNIDENTIFIED FORMATS (These should be shown to the digital archivists and in some cases samples will be needed from the government department. For large files or database files such as mdb check that DROID was set to -1 for byte scanning in settings.)', format)
    sheet1.write_row('B1:X1',['','','','','','','','','','','','','','','','','','','','','','','','','','',''],format)

unidentified()


def extension_only(): #function run to add a worksheet which selects all files which have only been identified by their extensions by DROID, also adds a new title row and makes it blue

    extension = csv
    extension = extension.loc[(extension['METHOD'] == "Extension") & (extension['SIZE'] > 0), :]
    extension = extension.sort_values('PUID')
    resultbook = results.book
    format = resultbook.add_format({
        'bold': True,
        'fg_color': '#4c9df7'})

    extension.to_excel(results, sheet_name='Extension_Only_ID',index=False, startcol = 0, startrow = 1)
    sheet1 = results.sheets['Extension_Only_ID']
    sheet1.write('A1', 'EXTENSION ONLY IDENTIFICATION (These are less securely identified file formats so good to scan the list and check if there is anything unusual)', format)
    sheet1.write_row('B1:X1',['','','','','','','','','','','','','','','','','','','','','','','','','','',''],format)

extension_only()

def multiple(): #function run to add a worksheet which selects all files which have been identified as multiple formats by DROID, also adds a new title row and makes it blue

    multiple = csv
    multiple = multiple.loc[(multiple['FORMAT_COUNT'] > 1) & (multiple['SIZE'] > 0), :]
    multiple = multiple.sort_values('EXT')
    resultbook = results.book
    format = resultbook.add_format({
        'bold': True,
        'fg_color': '#4c9df7'})

    multiple.to_excel(results, sheet_name='Multiple_ID',index=False, startcol = 0, startrow = 1)
    sheet1 = results.sheets['Multiple_ID']
    sheet1.write('A1', 'MULTIPLE IDENTIFICATION (Check original CSV for additional identifications. Do not worry about if identification is by extension only. Alert digital archivists if it is a signature identification and multiple id.)', format)
    sheet1.write_row('B1:X1',['','','','','','','','','','','','','','','','','','','','','','','','','','',''],format)

multiple()

def mismatch(): #function run to add a worksheet which selects all files which it identifies as having mismatched extensions to their format identification by DROID, also adds a new title row and makes it blue

    mismatch = csv
    mismatch = mismatch.loc[(mismatch['EXTENSION_MISMATCH'] == True) & (mismatch['SIZE'] > 0), :]
    mismatch = mismatch.sort_values('PUID')
    resultbook = results.book
    format = resultbook.add_format({
        'bold': True,
        'fg_color': '#4c9df7'})

    mismatch.to_excel(results, sheet_name='Extension_Mismatch',index=False, startcol = 0, startrow = 1)
    sheet1 = results.sheets['Extension_Mismatch']
    sheet1.write('A1', 'EXTENSION MISMATCH (A note will be generated for these- check if there are any obvious mistakes such as two extensions or an obvious data entry error that the department are happy to change. If not we are able to take these and will produce a note.)', format)
    sheet1.write_row('B1:X1',['','','','','','','','','','','','','','','','','','','','','','','','','','',''],format)

mismatch()

def container(): #function run to add a worksheet which selects all files which it identifies as compressed container formats by DROID, also adds a new title row and makes it blue

    container = csv
    container = container.loc[(container['TYPE'] == 'Container')]
    container = container.sort_values('PUID')
    resultbook = results.book
    format = resultbook.add_format({
        'bold': True,
        'fg_color': '#4c9df7'})

    container.to_excel(results, sheet_name='Compressed_Container_Formats',index=False, startcol = 0, startrow = 1)
    sheet1 = results.sheets['Compressed_Container_Formats']
    sheet1.write('A1', 'COMPRESSED CONTAINER FORMATS (No further action needed)', format)
    sheet1.write_row('B1:X1',['','','','','','','','','','','','','','','','','','','','','','','','','','',''],format)

container()

def zerobyte(): #function run to add a worksheet which selects all files which it identifies any zero byte files, also adds a new title row and makes it blue

    zerobyte = csv
    zerobyte = zerobyte.loc[(zerobyte['SIZE'] <= 0)]
    zerobyte = zerobyte.sort_values('EXT')
    resultbook = results.book
    format = resultbook.add_format({
        'bold': True,
        'fg_color': '#4c9df7'})

    zerobyte.to_excel(results, sheet_name='Zero_Byte_Files',index=False, startcol = 0, startrow = 1)
    sheet1 = results.sheets['Zero_Byte_Files']
    sheet1.write_row('A1:X1',['ZERO BYTE FILES','','','','These files are empty or can not be found, check that the department can find the original files','','','','','','','','','','','','','','','','','','','','','',''],format)

zerobyte()

def duplicates(): #function run to add a worksheet which selects all files which it identifies as having mismatched extensions to their format identification by DROID, also adds a new title row and makes it blue

    duplicates = csv
    duplicates = (duplicates.loc[duplicates['TYPE'].isin(['File','Container'])])
    duplicates = duplicates.loc[duplicates['SHA256_HASH'].duplicated(keep=False) & (duplicates['SIZE'] > 0), :]
    duplicates = duplicates.sort_values('SHA256_HASH')
    resultbook = results.book
    format = resultbook.add_format({
        'bold': True,
        'fg_color': '#4c9df7'})

    duplicates.to_excel(results, sheet_name='Duplicate_Files',index=False, startcol = 0, startrow = 1)
    sheet1 = results.sheets['Duplicate_Files']
    sheet1.write_row('A1:X1',['DUPLICATE FILES', '','','','Check with department if these need deaccessioning','','','','','','','','','','','','','','','','','','','','','','',''],format)

duplicates()

def FlaggedFormats(): #function run to add a worksheet which selects all files of formats identified by DROID which are on the flagged formats list, also adds a new title row and makes it blue

    FlaggedFormats = csv
    FlaggedFormats= FlaggedFormats.loc[(FlaggedFormats['FORMAT_COUNT'] > 0) & (FlaggedFormats['SIZE'] > 0), :]
    Flagged = {}
    Flagged = originalFlaggedFormats["label"].values.tolist()
    FlaggedFormats = FlaggedFormats.loc[FlaggedFormats.PUID.isin(Flagged)]
    FlaggedFormats = FlaggedFormats.sort_values('PUID')
    resultbook = results.book
    format = resultbook.add_format({
        'bold': True,
        'fg_color': '#4c9df7'})

    FlaggedFormats.to_excel(results, sheet_name='Flagged_Formats',index=False, startcol = 0, startrow = 1)
    sheet1 = results.sheets['Flagged_Formats']
    sheet1.write_row('A1:X1',['FLAGGED FORMATS','','','','These file formats are recommended for deaccessioning check with government department if they are happy to deaccession','','','','','','','','','','','','','','','','','','','','','',''],format)

FlaggedFormats()

def FurtherResearch(): #function run to add a worksheet which selects all files of formats identified by DROID which are on the flagged formats list, also adds a new title row and makes it blue

    FurtherResearch = csv
    FurtherResearch= FurtherResearch.loc[(FurtherResearch['FORMAT_COUNT'] > 0) & (FurtherResearch['SIZE'] > 0), :]
    Research = {}
    Research = originalFurtherResearch["label"].values.tolist()
    FurtherResearch = FurtherResearch.loc[FurtherResearch.PUID.isin(Research)]
    FurtherResearch = FurtherResearch.sort_values('PUID')
    resultbook = results.book
    format = resultbook.add_format({
        'bold': True,
        'fg_color': '#4c9df7'})

    FurtherResearch.to_excel(results, sheet_name='Further_Research',index=False, startcol = 0, startrow = 1)
    sheet1 = results.sheets['Further_Research']
    sheet1.write_row('A1:X1',['FURTHER RESEARCH','','','','This list of file formats is to be shown to a digital archivist for further research and guidance','','','','','','','','','','','','','','','','','','','','','',''],format)

FurtherResearch()

def whiteListFormats(): #function run to add a worksheet which selects all files of formats identified by DROID which are not in any lists, also adds a new title row and makes it blue

    whiteListFormats = csv
    whiteListFormats = whiteListFormats.loc[(whiteListFormats['FORMAT_COUNT'] > 0) & (whiteListFormats['SIZE'] > 0), :]
    whiteList = {}
    whiteList = originalWhiteList["label"].values.tolist()
    Flagged = {}
    Flagged = originalFlaggedFormats["label"].values.tolist()
    Research = {}
    Research = originalFurtherResearch["label"].values.tolist()
    whiteListFormats = whiteListFormats.loc[(~whiteListFormats.PUID.isin(whiteList)) & (~whiteListFormats.PUID.isin(Flagged)) & (~whiteListFormats.PUID.isin(Research)), :]
    whiteListFormats = whiteListFormats.sort_values('PUID')
    resultbook = results.book
    format = resultbook.add_format({
        'bold': True,
        'fg_color': '#4c9df7'})

    whiteListFormats.to_excel(results, sheet_name='Unlisted_Formats',index=False, startcol = 0, startrow = 1)
    sheet1 = results.sheets['Unlisted_Formats']
    sheet1.write_row('A1:X1',['FORMATS NOT ON LISTS','','','','We have to take all files that government departments ask us to but these should be sent to the digital archivists for further research','','','','','','','','','','','','','','','','','','','','','',''],format)

whiteListFormats()

results.close()