#! python3
# This will pull text files from the database system
import glob
import openpyxl
import os
os.chdir('S:\CSR\Contract Renewal Text Files')
sourceBook = openpyxl.load_workbook(glob.glob('*.xlsx'))
sourceSheet = sourceBook.sheetnames[0]

# See if there is a way to track what text files have already been pulled.
# I may have to have the macro call this program after the source sheet is
# formatted and then add "pulled" to a cell through this program as I pull
# text files.

# Maybe also ask the user for the first contract they would like to pull,
# or give the user the option of inputting a list of contracts they would
# like to pull (i.e. ones that came out blank) or leaving the input empty
# to pull all files. Could also nest these options.

# Go through, pulling text files and saving them under the contract name.
# Make sure this also pulls the non-contract items if that applies
