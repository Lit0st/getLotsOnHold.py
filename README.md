# getLotsOnHold.py

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
