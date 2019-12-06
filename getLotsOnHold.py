""" 
    This program uses the Excel sheet _2_LOH_REPORT.xls in the public H:
    drive to compile a list of lots currently on hold that include 
    specified keywords in their information and have been on hold less 
    than 90 days. Current keyword lists include the last tool the lot 
    ran on, hold operator name, and hold type (e.g. SPC hold or STL 
    holds [note: scheduled holds and dead stock holds are currently 
    excluded]). 
    
    The list of lots and their information is then compiled into a new
    Excel file titled LOH.xls and saved to the current working directory.
    
    Uses functions contained in excel_functions.py to manipulate excel 
    docs.
    
"""

import os
from excel_functions import *

os.chdir('H:\\PROD\\ML\\Public\\Automated Trackers')  # Change directory to Trackers
wb = xlrd.open_workbook('_2_LOH_REPORT.xls')  # Open the LOH Report Workbook.
DMR_sheet = wb.sheet_by_name('DMR') # Find the 'DMR' sheet. Contains all held lots.

os.chdir('C:\\Users\\rrumph\\Desktop\\python_work') # New file save location
f = 'LOH.xls'  # New file name

# Keyword lists.
tool_kw = ['LPC', 'LPH', 'LPN', 'LPP', 'LPT', 'MTE', 'APT', 'RVP', 
           'LMK']

holdop_kw = ['regan', 'rumph', 'natalia', 'olifer', 
             'rupinder rai', 'caryl', 'cayanan']
             
holdReason_kw = ['AbnEqHld(AbnEQEnd)', 'AbnEqHld(AbnInProcess)',
                'AbnEqHld(AbnQTY)', 'AbnPrdHld(AbnTestReslt)',
                'AbnPrdHld(SPC)', 'IMDHld', 'ImmediateHld',
                'EqGrpSTLHld', 'ExprMaintHld', 'MinrStpSTLHld']

# Sheet headers
LOH_headers = ['Lot', 'Tool Group', 'BPN', 'Queue', 'Hold Set By',
               'Notes', 'Hold Reason', 'Disposition', 'SPC Detail',
               'Eng Comment', 'DMR Comment']

LOH_wb = newWorkbook(LOH_headers, sheet_title='LOH')
LOH_sheet = LOH_wb.get_sheet('LOH')
LOH_wb.save(f)

lots = []

# Find lots that meet the tool name, hold operator name, and hold type criteria.
for rx in range(DMR_sheet.nrows):
    
    try:
        cx = 6  # Hold length (hours)
        val = DMR_sheet.cell(rx, cx).value
        if int(val) < 2160: # Excludes lots on hold >90 days
            cx = 7  # Tool column
            val = DMR_sheet.cell(rx, cx).value
            if assertCellValueInKeywordList(val, tool_kw) == True:
                cx = 2 # Hold reason column
                val = DMR_sheet.cell(rx, cx).value
                if assertCellValueInKeywordList(val, holdReason_kw) == True:
                    try:
                        LotInfo = getLotInfo(DMR_sheet, rx)
                        if LotInfo['LotNum'] not in lots:
                            updateLOH(f, LotInfo, LOH_wb, LOH_sheet)
                            lots.append(LotInfo['LotNum'])
                    except AttributeError:
                        continue

            cx = 15 # Hold Operator column
            val = DMR_sheet.cell(rx, cx).value
            if assertCellValueInKeywordList(val, holdop_kw) == True:
                cx = 2 # Hold reason column
                val = DMR_sheet.cell(rx, cx).value
                if assertCellValueInKeywordList(val, holdReason_kw) == True:
                    try:
                        LotInfo = getLotInfo(DMR_sheet, rx)
                        if LotInfo['LotNum'] not in lots:
                            updateLOH(f, LotInfo, LOH_wb, LOH_sheet)
                            lots.append(LotInfo['LotNum'])
                    except AttributeError:
                        continue
    except ValueError:
        continue
