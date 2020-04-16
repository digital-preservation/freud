#Script to Analyse DROID reports - Freud
#This script will analyse the DROID report to look for common issues which effect the ability to ingest into DRI. It can be run easily by running the batch file and then dragging the file which you want to analyse
#this will then create an excel spreadsheet in the directory which the program is running from with a different worksheet for each issue highlighted.

import pandas as pd
import numpy as np
import os

#below loads the csv file into pandas and takes required columns, additional columns for multiple identification are not taken yet as this breaks the csv read. It also loads a copy of the white list formats accepted into DRI
csvraw = input("Enter filepath of DROID csv to analyse: ")
csvraw = csvraw.strip('"')
columns_needed = ['ID','PARENT_ID','URI','FILE_PATH','NAME','METHOD','STATUS','SIZE','TYPE','EXT','LAST_MODIFIED','EXTENSION_MISMATCH','SHA256_HASH','FORMAT_COUNT','PUID','MIME_TYPE','FORMAT_NAME','FORMAT_VERSION']
csv = pd.read_csv(csvraw, usecols=columns_needed)
droidname = os.path.basename(csvraw)
droidname = droidname.rstrip('.csv')
results = pd.ExcelWriter(droidname+'_freudresults.xlsx', engine='xlsxwriter')
originalWhiteList = pd.read_csv('formats-whitelist.csv')

def unidentified(): #function run to add a worksheet which selects all files which have not been identified by DROID, also adds a new title row and makes it blue

    unidentified = csv
    unidentified = unidentified.loc[(unidentified['FORMAT_COUNT'] == 0)]
    unidentified = unidentified.sort_values('EXT')
    resultbook = results.book
    format = resultbook.add_format({
        'bold': True,
        'fg_color': '#4c9df7'})

    unidentified.to_excel(results, sheet_name='Unidentified_Formats',index=False, startcol = 0, startrow = 1)
    sheet1 = results.sheets['Unidentified_Formats']
    sheet1.write('A1', 'UNIDENTIFIED FORMATS', format)
    sheet1.write_row('B1:X1',['','','','','','','','','','','','','','','','','','','','','','','','','','',''],format)

unidentified()


def extension_only(): #function run to add a worksheet which selects all files which have only been identified by their extensions by DROID, also adds a new title row and makes it blue

    extension = csv
    extension = extension.loc[(extension['METHOD'] == "Extension")]
    extension = extension.sort_values('PUID')
    resultbook = results.book
    format = resultbook.add_format({
        'bold': True,
        'fg_color': '#4c9df7'})

    extension.to_excel(results, sheet_name='Extension_Only_ID',index=False, startcol = 0, startrow = 1)
    sheet1 = results.sheets['Extension_Only_ID']
    sheet1.write('A1', 'EXTENSION ONLY IDENTIFICATION', format)
    sheet1.write_row('B1:X1',['','','','','','','','','','','','','','','','','','','','','','','','','','',''],format)

extension_only()

def multiple(): #function run to add a worksheet which selects all files which have been identified as multiple formats by DROID, also adds a new title row and makes it blue

    multiple = csv
    multiple = multiple.loc[(multiple['FORMAT_COUNT'] > 1)]
    multiple = multiple.sort_values('EXT')
    resultbook = results.book
    format = resultbook.add_format({
        'bold': True,
        'fg_color': '#4c9df7'})

    multiple.to_excel(results, sheet_name='Multiple_ID',index=False, startcol = 0, startrow = 1)
    sheet1 = results.sheets['Multiple_ID']
    sheet1.write('A1', 'MULTIPLE IDENTIFICATION (Check original CSV for additional identifications)', format)
    sheet1.write_row('B1:X1',['','','','','','','','','','','','','','','','','','','','','','','','','','',''],format)

multiple()

def mismatch(): #function run to add a worksheet which selects all files which it identifies as having mismatched extensions to their format identification by DROID, also adds a new title row and makes it blue

    mismatch = csv
    mismatch = mismatch.loc[(mismatch['EXTENSION_MISMATCH'] == True)]
    mismatch = mismatch.sort_values('PUID')
    resultbook = results.book
    format = resultbook.add_format({
        'bold': True,
        'fg_color': '#4c9df7'})

    mismatch.to_excel(results, sheet_name='Extension_Mismatch',index=False, startcol = 0, startrow = 1)
    sheet1 = results.sheets['Extension_Mismatch']
    sheet1.write('A1', 'EXTENSION MISMATCH', format)
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
    sheet1.write('A1', 'COMPRESSED CONTAINER FORMATS', format)
    sheet1.write_row('B1:X1',['','','','','','','','','','','','','','','','','','','','','','','','','','',''],format)

container()

def zerobyte(): #function run to add a worksheet which selects all files which it identifies any zero byte files, also adds a new title row and makes it blue

    zerobyte = csv
    zerobyte = zerobyte.loc[(zerobyte['SIZE'] == 0)]
    zerobyte = zerobyte.sort_values('EXT')
    resultbook = results.book
    format = resultbook.add_format({
        'bold': True,
        'fg_color': '#4c9df7'})

    zerobyte.to_excel(results, sheet_name='Zero_Byte_Files',index=False, startcol = 0, startrow = 1)
    sheet1 = results.sheets['Zero_Byte_Files']
    sheet1.write('A1', 'ZERO BYTE FILES', format)
    sheet1.write_row('B1:X1',['','','','','','','','','','','','','','','','','','','','','','','','','','',''],format)

zerobyte()

def duplicates(): #function run to add a worksheet which selects all files which it identifies as having mismatched extensions to their format identification by DROID, also adds a new title row and makes it blue

    duplicates = csv
    duplicates = (duplicates.loc[duplicates['TYPE'].isin(['File','Container'])])
    duplicates = duplicates.loc[duplicates['SHA256_HASH'].duplicated(keep=False), :]
    duplicates = duplicates.sort_values('SHA256_HASH')
    resultbook = results.book
    format = resultbook.add_format({
        'bold': True,
        'fg_color': '#4c9df7'})

    duplicates.to_excel(results, sheet_name='Duplicate_Files',index=False, startcol = 0, startrow = 1)
    sheet1 = results.sheets['Duplicate_Files']
    sheet1.write('A1', 'DUPLICATE FILES', format)
    sheet1.write_row('B1:X1',['','','','','','','','','','','','','','','','','','','','','','','','','','',''],format)

duplicates()

def whiteListFormats(): #function run to add a worksheet which selects all files of formats identified by DROID which are not on the DRI white list, also adds a new title row and makes it blue

    whiteListFormats = csv
    whiteListFormats = whiteListFormats.loc[(whiteListFormats['FORMAT_COUNT'] > 0)]
    whiteList = {}
    whiteList = originalWhiteList["label"].values.tolist()
    whiteListFormats = whiteListFormats.loc[~whiteListFormats.PUID.isin(whiteList)]
    whiteListFormats = whiteListFormats.sort_values('PUID')
    resultbook = results.book
    format = resultbook.add_format({
        'bold': True,
        'fg_color': '#4c9df7'})

    whiteListFormats.to_excel(results, sheet_name='Formats_Not_On_White_List',index=False, startcol = 0, startrow = 1)
    sheet1 = results.sheets['Formats_Not_On_White_List']
    sheet1.write('A1', 'FORMATS NOT ON WHITELIST', format)
    sheet1.write_row('B1:X1',['','','','','','','','','','','','','','','','','','','','','','','','','','',''],format)

whiteListFormats()

results.save()

