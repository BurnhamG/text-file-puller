#! python3
# This will pull text files from the database system
from datetime.datetime import now
import glob
import openpyxl
import os
import pyautogui
from pywinauto import application
import time

os.chdir('S:\CSR\Contract Renewal Text Files')
sourceBook = openpyxl.load_workbook(glob.glob('*.xlsx'))
sourceSheet = sourceBook.sheetnames[0]
emptyCount = 0
listOfFiles = glob.glob('*.txt')
listOfFiles_dict = {}

print('Gathering a list of files that already exist...')

for i in listOfFiles:
    listOfFiles_dict[i] = time.localtime(os.stat(i).st_mtime)

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

# Checks for files that have already been pulled
for i in range(2, sourceSheet.max_row):  # Skips header row

    # Check what column should be examined
    if sourceSheet.cell(row=i, column=5) + '.txt' in listOfFiles:
        modified_date = \
            ' '.join('-'.join(listOfFiles_dict[
                              sourceSheet.cell(row=i, column=5).value].tm_mon,
                              listOfFiles_dict[
                              sourceSheet.cell(row=i, column=5).value].tm_mday,
                              listOfFiles_dict[
                              sourceSheet.cell(row=i, column=5).value].tm_year),
                     ':'.join(listOfFiles_dict[
                         sourceSheet.cell(row=i, column=5).value].tm_hour,
                         listOfFiles_dict[
                         sourceSheet.cell(row=i, column=5).value].tm_min,
                         listOfFiles_dict[
                         sourceSheet.cell(row=i, column=5).value].tm_sec))
        sourceSheet.cell(row=i, column=9).value = modified_date

    # Checks for text files to ignore
    if sourceSheet.cell(row=i, column=8) in listToAvoid:
        sourceSheet.cell(row=i, column=9).value = 'IGNORED'

# Ask the user for the first contract they would like to pull, or give the user
# the option of inputting a list of contracts they would like to pull
# (i.e. ones that came out blank) or leaving the input empty to pull all files.

print('What is the first contract you would like to pull? Leave this empty to \
        pull all text files.')
startContract = input()

if startContract.strip():
    for i in range(2, sourceSheet.max_row):
        if sourceSheet.cell(row=i, column=5) == startContract:
            # Here is where I need to set focus on the window and then input

# Setting current time = now().strftime('%m-%d-%Y %H:%M:%S')

# Go through, pulling text files and saving them under the contract name.
# Make sure this also pulls the non-contract items if that applies.
