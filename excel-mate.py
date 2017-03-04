# -*- coding: utf-8 -*-
"""

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

"""
#====================

import xlrd
import regex

# set the file location; use double \\
# home
path_home = "Z:\\Documents\\dev\\RAND\\xlrd-test\\prolog.xls"
# RAND test
path_work_test = "\\\\data\\userdata$\\cfitzgerald\\Desktop\\dev\\xlrd-test\\prolog.xls"
# RAND
path_work = "\\\\data\\Proposal Log\\Proposal Log 1-1-14 - present.xls"
# open the Excel file (.xls only)
workbook = xlrd.open_workbook(path_work, formatting_info=True)
# set sheetMain to "In-Office Proposal Log" tab
sheet_main = workbook.sheet_by_name("In-Office Proposal Log")
# set sheetMain to "In-Office Proposal Log" tab
sheet_energy = workbook.sheet_by_name("In-Office Energy Log")

#-----------------------------
# create a variable for string formatting methods
# for one digit to the right of the decimal point, with % symbol
format_percentage = '{0:.1f}%'
# e.g. 05 instead of 5
format_double_digit = '{0:02d}'

#-----------------------------
# define a function that creates an array of valid entries, per column
# takes the column index as an argument
# excludes italic fonts; empty (type 0) & blank (type 6) cells
# includes (type 1) cells
# returns an array of valid entries
def find_valid_entries ( column, sheet ):
    
    # create an empty array for valid entries
    ve = []

    # iterate through all rows, starting from the third row (2)
    for row in range(2, sheet.nrows):
        # assign font_list; seems to be necessary, just go with it for now...        
        workbook_fonts = workbook.font_list
        # assign cell value        
        cell_value = sheet.cell_value(row, column)        
        # assign cell type
        cell_type = sheet.cell_type(row, column)
        # assign cell font
        cell_font = workbook.xf_list[sheet.cell_xf_index(row, column)]
        # if cell IS valid type (1) and NOT italicized...
        # add cell value to ve_main array
        if cell_type == 1 and not workbook_fonts[cell_font.font_index].italic:
            ve.append(cell_value)
    
    # return the array      
    return(ve)

#-----------------------------
# assign valid entries array to column/sheet-specific variable
# for CS
ve_column_q_main = find_valid_entries(16, sheet_main)
ve_column_q_energy = find_valid_entries(16, sheet_energy)
# for Team
ve_column_j_main = find_valid_entries(9, sheet_main)
ve_column_j_energy = find_valid_entries(9, sheet_energy)
# for notes
ve_column_t_main = find_valid_entries(19, sheet_main)
ve_column_t_energy = find_valid_entries(19, sheet_energy)

# assign number of valid entries
# for CS (main & energy)
ve_count_column_q_main = len(ve_column_q_main)
ve_count_column_q_energy = len(ve_column_q_energy)
# for Team (main)
ve_count_column_j_main = len(ve_column_j_main)
ve_count_column_j_energy = len(ve_column_j_energy)
# for notes (main & energy)
ve_count_column_t_main = len(ve_column_t_main)
ve_count_column_t_energy = len(ve_column_t_energy)

#-----------------------------
# create regex variables using regex.compile
# multiple entries must be iterated   
draft = regex.compile(r'.*(to draft).*')
need = [regex.compile(status) for status in [r'.*(to advise).*',
                                             r'.*(to provide).*',
                                             r'.*(Client).*',
                                             r'.*(PEV).*',
                                             r'.*(Zob).*',
                                             r'.*(to review).*',
                                             r'.*(to price).*',
                                             r'.*(visit).*']]
final = regex.compile(r'.*(SV has).*')

#---------------------
# class for Status entries
class Status:
    
    def __init__( self, column ):
        self.column = column

        self.array_index_draft = []
        self.count_draft = 0
        
        self.array_index_need = []
        self.count_need = 0
        
        self.array_index_final = []
        self.count_final = 0
        
        # call draft/need/final methods
        self.ve_status_draft()
        self.ve_status_need()
        self.ve_status_final()

        # create array variable for the combined category indices (AFTER their methods have been called)       
        # merge the arrays for draft, need, and final and assign
        # needed to find misc entries
        self.array_index_categorized = sorted(self.array_index_draft + 
                                              self.array_index_need + 
                                              self.array_index_final)
                                                
        # not actually finding "misc" value/location
        # will need to determine index through values not appearing in other categories    
        self.array_index_misc = []
        self.count_misc = 0

        # call misc method
        self.ve_status_misc()
        
    # determine status for DRAFT category; mutate count and array index
    def ve_status_draft( self ):
        # iterate through each entry in ve_column_t_main
        # index_main (starting from 0) is incremented at the end
        # this allows the "index" of each regex match to be appended to the arrayIndex    
        for index, entry in enumerate(self.column):
            if regex.search(draft, entry):
                self.count_draft += 1
                self.array_index_draft.append(index)
                
    # determine status for need category; mutate count and array index   
    def ve_status_need( self ):
        for index, entry in enumerate(self.column):
            for status in need:
                if regex.search(status, entry):
                    self.count_need += 1
                    self.array_index_need.append(index)
                    break
            
    # determine status for FINAL category; mutate count and array index
    def ve_status_final( self ):
        for index, entry in enumerate(self.column):
            if regex.search(final, entry):
                self.count_final += 1
                self.array_index_final.append(index)
            
    # determine status for MISC category; mutate array index
    def ve_status_misc( self ):
        # use enumerate to iterate through the index numbers of valid entries array
        # if the index number is NOT in the categorized index array
        # append the index number to the misc index array
        for index, entry in enumerate(self.column):
            if index not in self.array_index_categorized:
                self.array_index_misc.append(index)
                
#-----------------------------
# create instances of the STATUS class
main = Status(ve_column_t_main)
energy = Status(ve_column_t_energy)

# get the counts for main and energy instances
count_draft = main.count_draft + energy.count_draft
count_need = main.count_need + energy.count_need
count_final = main.count_final + energy.count_final
# get the count for entries not in the "categories"
count_misc = (ve_count_column_t_main + ve_count_column_t_energy) - (count_draft + count_need + count_final)

# percentages
def status_percentage( status ):
        return(format_percentage.format((status / (ve_count_column_j_main + ve_count_column_j_energy))*100))

count_draft_percentage = status_percentage(count_draft)
count_need_percentage = status_percentage(count_need)
count_final_percentage = status_percentage(count_final)
count_misc_percentage = status_percentage(count_misc)

#---------------------
# class for CS team members
class CS:
    
    def __init__( self, name, label ):
        self.name = name
        self.label = label
        
    def cs_entries( self ):
        self.cs_count = 0
        self.cs_entries_index_main = []
        self.cs_entries_index_energy = []
        index_main = 0
        index_energy = 0
        
        # In-Office PROPOSAL Log
        # increment cs_count if a match
        # append the array index into a new array, for use in comparing with the note entries array
        for entry in ve_column_q_main:
            if entry == self.label:
                self.cs_count += 1
                self.cs_entries_index_main.append(index_main)
            index_main += 1

        # In-Office ENERGY Log
        for entry in ve_column_q_energy:
            if entry == self.label:
                self.cs_count += 1
                self.cs_entries_index_energy.append(index_energy)
            index_energy += 1
                            
        return(self.cs_count)
    
    def cs_percentage( self ):
        return(format_percentage.format((self.cs_entries() / (ve_count_column_q_main + ve_count_column_q_energy))*100))
        
    # compare two array indices
    # index for matching cs_entries?
    # index for validEntriesColumnT
    def cs_entry_notes( self ):
        self.cs_entry_notes = []
                
        for index in self.cs_entries_index_main:
            self.cs_entry_notes.append(ve_column_t_main[index])
            #print(index)
                
        return(self.cs_entry_notes)

    # DRAFT status/percentage methods     
    def cs_entry_status_draft( self ):
        self.count_draft = 0  
        
        for index in self.cs_entries_index_main:
            if index in main.array_index_draft:
                self.count_draft += 1

        for index in self.cs_entries_index_energy:
            if index in energy.array_index_draft:
                self.count_draft += 1
                
        return(self.count_draft)
        
    def cs_entry_percentage_draft( self ):
        return(format_percentage.format((self.cs_entry_status_draft() / self.cs_entries())*100))

    # need status/percentage methods 
    def cs_entry_status_need( self ):
        self.count_need = 0  
        
        for index in self.cs_entries_index_main:
            if index in main.array_index_need:
                self.count_need += 1
        
        for index in self.cs_entries_index_energy:
            if index in energy.array_index_need:
                self.count_need += 1        
        
        return(self.count_need)

    def cs_entry_percentage_need( self ):
        return(format_percentage.format((self.cs_entry_status_need() / self.cs_entries())*100))
    
    # FINAL status/percentage methods    
    def cs_entry_status_final( self ):
        self.count_final = 0    
        
        for index in self.cs_entries_index_main:
            if index in main.array_index_final:
                self.count_final += 1
        
        for index in self.cs_entries_index_energy:
            if index in energy.array_index_final:
                self.count_final += 1        
        
        return(self.count_final)
    
    def cs_entry_percentage_final( self ):
        return(format_percentage.format((self.cs_entry_status_final() / self.cs_entries())*100))

    # MISC status/percentage methods         
    def cs_entry_status_misc( self ):
        self.count_misc = 0        
        
        for index in self.cs_entries_index_main:
            if index in main.array_index_misc:
                self.count_misc += 1
        
        for index in self.cs_entries_index_energy:
            if index in energy.array_index_misc:
                self.count_misc += 1

        return(self.count_misc)

    def cs_entry_percentage_misc( self ):
        return(format_percentage.format((self.cs_entry_status_misc() / self.cs_entries())*100))

#---------------------        
# class for RAND teams
class Team:
    
    def __init__( self, label ):
        self.label = label
        
    def rand_team_entries( self ):
        self.count_team = 0
    
        # In-Office PROPOSAL Log
        for entry in ve_column_j_main:
            if entry == self.label:
                self.count_team += 1

        # In-Office ENERGY Log
        for entry in ve_column_j_energy:
            if entry == self.label:
                self.count_team += 1
    
        return(self.count_team)
    
    def team_percentage( self ):
        return(format_percentage.format((self.rand_team_entries() / (ve_count_column_j_main + ve_count_column_j_energy))*100))

#-----------------------------
# tallies

drafted_text = "\tto be drafted"
needing_input_text = "\tneeding input"
finalized_text = "\tto be finalized"
uncategorized_text = "\tuncategorized"

print("================================") 
print("TOTAL:\t",
      ve_count_column_q_main + ve_count_column_q_energy)

print("--------------------------------")

print("Proposal Log:\t",
      ve_count_column_q_main)
print("Energy Log:\t",
      ve_count_column_q_energy)

print("--------------------------------")

print(format_double_digit.format(count_draft),
      "/",
      count_draft_percentage,
      drafted_text)
print(format_double_digit.format(count_need),
      "/",
      count_need_percentage,
      needing_input_text)
print(format_double_digit.format(count_final),
      "/",
      count_final_percentage,
      finalized_text)
print(format_double_digit.format(count_misc),
      "/",
      count_misc_percentage,
      uncategorized_text)

#-----------------------------
# create instances of the RAND team class
arch = Team("Architectural")
energy = Team("Energy")
code = Team("Code")
facade = Team("Facade")
forensics = Team("Forensics")
mep = Team("MEP")
structural = Team("Structural")

# create an array for all Team instances
teams_array = [arch,
               energy,
               code,
               facade,
               forensics,
               mep,
               structural]

# print the randTeam tallies in a single function
def team_tally( rand_team ):
    
    print(format_double_digit.format(rand_team.rand_team_entries()),
          "/",
          rand_team.team_percentage(),
          "\tfor",
          rand_team.label)

# call the tally function for each instance in the array
print("================================") 
for teams in teams_array:
    team_tally(teams)

#-----------------------------
# create instances of the CS team member classs
csCF = CS("CF", "FitzGerald")
csNA = CS("NA", "Abreu")
csBF = CS("BF", "Feldman")
csCC = CS("CC", "Carr")

# create an array for all CS instances
csr_array = [csCF,
             csNA,
             csBF,
             csCC]

# print the CS tallies in a single function
def cs_tally( csr ):
    
    print("================================") 
    print(csr.cs_entries(),
          "/",
          csr.cs_percentage(),
          "\tfor",
          csr.name)
    #print(csr.cs_entry_notes())
    print("--------------------------------")
    print(format_double_digit.format(csr.cs_entry_status_draft()),
          "/",
          csr.cs_entry_percentage_draft(),
          drafted_text)
    print(format_double_digit.format(csr.cs_entry_status_need()),
          "/",
          csr.cs_entry_percentage_need(),
          needing_input_text)
    print(format_double_digit.format(csr.cs_entry_status_final()),
          "/",
          csr.cs_entry_percentage_final(),
          finalized_text)
    print(format_double_digit.format(csr.cs_entry_status_misc()),
          "/",
          csr.cs_entry_percentage_misc(),
          uncategorized_text)

# call the tally function for each instance in the array
for csr in csr_array:
    cs_tally(csr)

print("--------------------------------")