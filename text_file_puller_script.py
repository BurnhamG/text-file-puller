#! python3
"""This will pull text files from the system.

Much of this is based on code from automatetheboringstuff.com, particularly
regarding Excel spreadsheets.
"""
# Clean these up
from ctypes import windll
from datetime import datetime
import glob
import openpyxl
import os
import psutil
import pyautogui
from pywinauto import application
import re
import time
import win32gui
import win32process


def window_callback(winHandle, pid):
    """List the handle of the selected process based on process ID.

    This is the callback function for win32gui.EnumWindows. winHandle is
    the hwnd, the handle to a window.
    """
    tid, current_pid = win32process.GetWindowThreadProcessId(winHandle)
    if pid == current_pid and win32gui.IsWindowVisible(winHandle):
        appWindows.append(winHandle)


def listEmailGroups():
    """Identify the groups of representatives to avoid certain files."""
    print('Identifying groups for emailing...')


    names_list = []

    with open(os.path.join('DataFiles', 'CcList.txt'),
              encoding='utf-8') as file:
        names_list.append([line.rstrip() for line in file])

    with open(os.path.join('DataFiles', 'NonConList.txt'),
              encoding='utf-8') as file:
        names_list.append([line.rstrip() for line in file])

    with open(os.path.join('DataFiles', 'RepsThatNeedCc.txt'),
              encoding='utf-8') as file:
        names_list.append([line.rstrip() for line in file])

    with open(os.path.join('DataFiles', 'NoProcess.txt'),
              encoding='utf-8') as file:
        names_list.append([line.rstrip() for line in file])

    return names_list


def stepRecognize(inputImage):
    """Check for text on screen. If none matches, stop."""
    readyStatus = pyautogui.locateOnScreen(os.path.join(
                                           'DataFiles/Images', inputImage),
                                           minSearchTime=.5)
    if not readyStatus:
        print('Image not found.')
        return False
    else:
        return True


def getKeystrokes():
    """Get an array of the keystrokes used when pulling text files."""
    print('Obtaining keystrokes...')

    keystrokes_list = []

    with open(os.path.join('DataFiles', 'Keystrokes.txt'),
              encoding='utf-8') as file:
        keystrokes_list = file.readlines()
        # Removes comments
        keystrokes_list = [[x.split('|') for x in [x.replace(', ', '|')
                            for x in keystrokes_list
                            if not x.startswith('#')]]]

    return keystrokes_list


def getWindow(inAppWindow):
    """Set focus on the window and begin input."""
    for proc in psutil.process_iter():
        procName = proc.name()
        if re.match('*mvbase*', procName.lower()):
            processID = proc.pid
            processName = proc.name()
            bringToFront = windll.user32.SetWindowPos
    win32gui.EnumWindows(window_callback, processID)
    windowLoc = win32gui.GetWindowRect(inAppWindow[0])
    return windowLoc


def findAlreadyPulled():
    """Mark files that have already been pulled with the pull time."""
    if sourceSheet.cell(row=i, column=5) + '.txt' in listOfFiles:
        modified_date = ' '.join([
                                '-'.join([
                                        str(listOfFiles_dict[
                                            sourceSheet.cell(row=i, column=5)
                                                .value]
                                            .tm_mon).strformat(%m),
                                        str(listOfFiles_dict[
                                            sourceSheet.cell(row=i, column=5)
                                                .value]
                                            .tm_mday).strformat(%d),
                                        str(listOfFiles_dict[
                                            sourceSheet.cell(row=i, column=5)
                                                .value]
                                            .tm_year).strformat(%y)
                                        ]),
                                ':'.join([
                                        str(listOfFiles_dict[
                                            sourceSheet.cell(row=i, column=5)
                                                .value]
                                            .tm_hour).strformat(%H),
                                        str(listOfFiles_dict[
                                            sourceSheet.cell(row=i, column=5)
                                                .value]
                                            .tm_min).strformat(%M),
                                        str(listOfFiles_dict[
                                            sourceSheet.cell(row=i, column=5)
                                                .value]
                                            .tm_sec).strformat(%S)
                                        ])
                                ])
    sourceSheet.cell(row=i, column=9).value = modified_date


def getContractInfo():
    contractInfo = {}
    for row in range(2, sourceSheet.max_row + 1):
        contractNo = sourceSheet['A' + str(row)]
        companyNo = sourceSheet['E' + str(row)]
        companyRep = sourceSheet['G' + str(row)]
        contractRep = sourceSheet['H' + str(row)]

        # Ensure key for contract exists
        contractInfo.setdefault(contractNo, {'Companies': [companyNo],
                                            'CompanyRep': companyRep,
                                            'ContractRep': contractRep}
                                )
        if sourceSheet['A' + str(row + 1)].value == None:
            contractInfo[contractNo]['Companies'].append(sourceSheet['E' +
                                                                    str(i + 1)
                                                                    ]
                                                        )





os.chdir('S:\CSR\Contract Renewal Text Files')
sourceBook = openpyxl.load_workbook(glob.glob('*.xlsx'))
sourceSheet = sourceBook.sheetnames[0]
emptyCount = 0
listOfFiles = glob.glob('*.txt')
listOfFiles_dict = {}
# Modification time of the files
for i in listOfFiles:
    listOfFiles_dict[i] = time.localtime(os.stat(i).st_mtime)

appWindows = []
allNames = listEmailGroups()
keystrokes = getKeystrokes()

listToAvoid = allNames[3]
# Checks for files that have already been pulled
for i in range(2, sourceSheet.max_row + 1):  # Skips header row
    # TODO: Check what column should be examined
    findAlreadyPulled()
    # Checks for text files to ignore
    if sourceSheet.cell(row=i, column=8) in listToAvoid:
        sourceSheet.cell(row=i, column=9).value = 'IGNORED'

# Ask the user for the first contract they would like to pull, or give the user
# the option of inputting a list of contracts they would like to pull
# (i.e. ones that came out blank) or leaving the input empty to pull all files.
print('What contracts would you like to pull?')
print('If entering a range, separate the numbers with a hyphen, or')
print('leave this empty to pull all text files.')
print('Type "end" to exit.')
pullContracts = input()

if pullContracts.strip() == 'end':
    return SystemExit
elif pullContracts.strip() != '':
    currentcontractNumber = startContract = pullContracts.split('-')[0]
    endContract = pullContracts.split('-')[1]
    windowCoords = getWindow(appWindows)
    winX = windowLoc[0]
    winY = windowLoc[1]
    winWidth = windowLoc[2] - winX
    winHeight = windowLoc[3] - winY

    # Here the -1 effectively locks the window on top. Make sure to
    # change this back!
    bringToFront(appWindows[0], -1, winX, winY, 0, 0, 0x0001)

    # Here we actually start interacting with the database
    pyautogui.click(winX + (winWidth / 2), winY + (winHeight / 2))
    pyautogui.typewrite(keystrokes[0] * 5)
    pyautogui.typewrite(keystrokes[1])
    for i in range(2, sourceSheet.max_row + 1):

        if sourceSheet.cell(row=i, column=1).value == currentContractNumber:

            # Check for image
            if not stepRecognize('Step1.png'):
                raise SystemExit
            if stepRecognize('Step1.png'):
                pyautogui.typewrite([datetime.now().month, '.01.',
                                    datetime.now().year - 1, 'enter'])
                pyautogui.typewrite([datetime.now().month, datetime.now().day,
                                    datetime.now().year], 'enter'])
                pyautogui.typewrite([sourceSheet.cell(row=i, column=5).value])
                pyautogui.typewrite([sourceSheet.cell(row=i, column=1).value])
                # Check for multiple companies on same contract
                if sourceSheet.cell(row= i + 1, column=1).value == None:

# Setting current time = datetime.now().strftime('%m-%d-%Y %H:%M:%S')

# Go through, pulling text files and saving them under the contract name.
# Make sure this also pulls the non-contract items if that applies.
