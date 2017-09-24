#! python3
# This will pull text files from the database system
from datetime.datetime import now
import glob
import openpyxl
import os
os.chdir('S:\CSR\Contract Renewal Text Files')
sourceBook = openpyxl.load_workbook(glob.glob('*.xlsx'))
sourceSheet = sourceBook.sheetnames[0]
emptyCount = 0
listOfFiles = glob.glob('*.txt')

with open(os.path.join('DataFiles', 'CcList.txt'),
          encoding='utf-8') as file:
    listOfCCRecipients = [line.rstrip() for line in file]

with open(os.path.join('DataFiles', 'NonConList.txt'),
          encoding='utf-8') as file:
    listOfNonConReps = [line.rstrip() for line in file]

with open(os.path.join('DataFiles', 'RepsThatNeedCc.txt'),
          encoding='utf-8') as file:
    listOfRepsNeedingCc = [line.rstrip() for line in file]

with open(os.path.join('DataFiles', 'NoProcess.txt'),
          encoding='utf-8') as file:
    listToAvoid = [line.rstrip() for line in file]

for i in range(2, sourceSheet.max_row):  # Skips header row
    # Check what column should be examined
    if sourceSheet.cell(row=i, column=5) in listOfFiles:
        sourceSheet.cell(row=i, column=9).value = \
            now().strftime('%m-%d-%Y %H:%M:%S')
    elif sourceSheet.cell(row=i, column=8) in listToAvoid:
        sourceSheet.cell(row=i, column=9).value = 'IGNORED'

# See if there is a way to track what text files have already been pulled.
# I may have to have the macro call this program after the source sheet is
# formatted and then add "pulled" to a cell through this program as I pull
# text files.

# Maybe also ask the user for the first contract they would like to pull,
# or give the user the option of inputting a list of contracts they would
# like to pull (i.e. ones that came out blank) or leaving the input empty
# to pull all files. Could also nest these options.

# Go through, pulling text files and saving them under the contract name.
# Make sure this also pulls the non-contract items if that applies.
