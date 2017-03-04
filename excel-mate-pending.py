# -*- coding: utf-8 -*-
"""
JOBS
columns (zero indexed), for reference
0  = A: "Proposal Number"
1  = B: "Mo"
2  = C: "Day"
3  = D: "Year"
4  = E: "Address"
5  = F: "Job Description"
6  = G: "Job?"
7  = H: "End Date"
8  = I: "Client"
9  = J: "Team"
10 = K: "Fee"
11 = L: "NB?"
12 = M: "NC?"
13 = N: "NJ?"
14 = O: "Acct #"
15 = P: "Scope"
16 = Q: "Prep"
17 = R: "Budget"
18 = S: "W"
19 = T: "Notes"
20 = U: "RFP Date"
21 = V: "Due Date"
22 = W: "# of Days"

PENDING JOBS

"""

import xlrd

# set the file location; use double \\
# home
pathHome = "Z:\\Documents\\dev\\RAND\\xlrd-test\\prolog.xls"
# RAND test
pathWorkTEST = "\\\\data\\userdata$\\cfitzgerald\\Desktop\\dev\\xlrd-test\\prolog.xls"
# RAND
pathWork = "\\\\data\\Proposal Log\\Proposal Log 1-1-14 - present.xls"
# open the Excel file (.xls only)
workbook = xlrd.open_workbook(pathHome, formatting_info=True)
# set sheetMain to "In-Office Proposal Log" tab
sheetJobs = workbook.sheet_by_name("Jobs")
# set sheetMain to "In-Office Proposal Log" tab
sheetPending = workbook.sheet_by_name("Pending Jobs")

# define a function that creates an array of valid entries, per column
# takes the column index as an argument
# excludes bold fonts; empty (type 0) & blank (type 6) cells
# includes (type 1) cells
# returns an array of valid entries
def validEntriesByColumn ( columnIndex ):
    
    # Jobs
    # create an empty array for valid entries
    validEntriesJobs = []

    # iterate through all rows, starting from the third row (2)
    for row in range(2, sheetJobs.nrows):
        # assign font_list; seems to be necessary, just go with it for now...        
        workbookFonts = workbook.font_list
        # assign cell value        
        cellValue = sheetJobs.cell_value(row, columnIndex)        
        # assign cell type
        cellType = sheetJobs.cell_type(row, columnIndex)
        # assign cell font
        cellFont = workbook.xf_list[sheetJobs.cell_xf_index(row, columnIndex)]
        # if cell IS valid type (1) and NOT italicized...
        # add cell value to validEntriesMain array
        if cellType == 1 and not workbookFonts[cellFont.font_index].bold:
            validEntriesJobs.append(cellValue)
            
    # Pending Jobs
    validEntriesPending = []

    for row in range(2, sheetPending.nrows):
        workbookFonts = workbook.font_list
        cellValue = sheetPending.cell_value(row, columnIndex)        
        cellType = sheetPending.cell_type(row, columnIndex)
        cellFont = workbook.xf_list[sheetPending.cell_xf_index(row, columnIndex)]
        if cellType == 1 and not workbookFonts[cellFont.font_index].bold:
            validEntriesPending.append(cellValue)
    
    # return a tuple of the arrays      
    return(validEntriesJobs, validEntriesPending)

# unpack/assign valid entries array to column-specific variable    
validEntriesJobsColumnA, validEntriesPendingColumnA = validEntriesByColumn(0)

# assign any proposal number value that appears in both arrays (or nothing)
duplicateProposalNumber = set(validEntriesJobsColumnA) & set(validEntriesPendingColumnA)

# if there is a duplicate entry, print the proposal number (using str to remove '{ and }')
# else print a message that there are no duplicate entries
if duplicateProposalNumber:
    print("Remove from 'Pending Jobs':", str(duplicateProposalNumber)[2:-2])
else:
    print("'Pending Jobs' has no 'Job' entries.")
