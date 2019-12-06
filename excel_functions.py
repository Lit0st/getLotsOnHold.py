# List of excel functions for python macros
import os

import xlrd
import xlwt


def newWorkbook(column_headers=[], sheet_title='Sheet1'):
    """Create a new workbook and sheet with column headers."""

    wb = xlwt.Workbook() # Create a new workbook
    sheet = wb.add_sheet(sheet_title) # Add a new sheet and name it
    hdrow = sheet.row(0)

    r = len(column_headers)
    
    # Label each column with the headers and set the width
    for i in range(0, r):   
        hdrow.write(i, column_headers[i])
        hdcol = sheet.col(i)
        hdcol.width = 5000

    return wb

def getLotInfo(sheet, rx):
    """Retrieve Lot Info from the sheet and store it in a dictionary."""
    LotInfo = {
        'LotNum': sheet.cell(rx, colx=0).value,
        'BPN': sheet.cell(rx, colx=8).value,
        'ToolGroup': sheet.cell(rx, colx=7).value,
        'Queue': sheet.cell(rx, colx=6).value,
        'HoldReason': sheet.cell(rx, colx=2).value,
        'SPCDetail': sheet.cell(rx, colx=21).value,
        'DMRComment': sheet.cell(rx, colx=18).value,
        'ENGComment': sheet.cell(rx, colx=21).value,
        'HoldSetBy': sheet.cell(rx, colx=15).value,
        }
    return LotInfo


def updateLOH(filename, LotInfo, LOHwb, LOHsheet):
    """Add lot info to the new LOH sheet."""
    
    nextRow = xlrd.open_workbook(filename).sheet_by_name('LOH').nrows
    
    for key in LotInfo.keys():
        if key == 'LotNum':
            c = 0
        elif key == 'ToolGroup':
            c = 1
        elif key == 'BPN':
            c = 2
        elif key == 'Queue':
            c = 3
        elif key == 'HoldSetBy':
            c = 4
        elif key == 'SPCDetail':
            c = 7
        elif key == 'HoldReason':
            c = 8
        elif key == 'ENGComment':
            c = 9
        elif key == 'DMRComment':
            c = 10
        LOHsheet.write(nextRow, c, LotInfo[key])
    
    LOHwb.save(filename)
    return LOHwb

def assertCellValueInKeywordList(cellValue, keywordList):
    """Check to see if the cell value is in the keyword list."""
    l = [i.lower() for i in keywordList]
    if cellValue.lower() in l:
        return True
    else:
        return False
    
