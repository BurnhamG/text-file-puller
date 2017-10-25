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

    with open(os.path.join('DataFiles', 'NonConList.txt'),
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
    # TODO: Reintegrate this with main code
    for proc in psutil.process_iter():
        procName = proc.name()
        if re.match('*mvbase*', procName.lower()):
            processID = proc.pid
    win32gui.EnumWindows(window_callback, processID)
    windowLoc = win32gui.GetWindowRect(inAppWindow[0])

    return windowLoc


def findAlreadyPulled():
    """Mark files that have already been pulled with the pull time."""
    if sourceSheet.cell(row=i, column=5) + '.txt' in listOfFiles:
        # TODO: Simplify this with a variable
        modified_date = ' '.join([
                                 '-'.join([
                                          str(listOfFiles_dict[
                                              sourceSheet.cell(row=i, column=5)
                                              .value]
                                              .tm_mon).strformat('%m'),
                                          str(listOfFiles_dict[
                                              sourceSheet.cell(row=i, column=5)
                                              .value]
                                              .tm_mday).strformat('%d'),
                                          str(listOfFiles_dict[
                                              sourceSheet.cell(row=i, column=5)
                                              .value]
                                              .tm_year).strformat('%y')]),
                                 ':'.join([
                                          str(listOfFiles_dict[
                                              sourceSheet.cell(row=i, column=5)
                                              .value]
                                              .tm_hour).strformat('%H'),
                                          str(listOfFiles_dict[
                                              sourceSheet.cell(row=i, column=5)
                                              .value]
                                              .tm_min).strformat('%M'),
                                          str(listOfFiles_dict[
                                              sourceSheet.cell(row=i, column=5)
                                              .value]
                                              .tm_sec).strformat('%S')])])
    sourceSheet.cell(row=i, column=9).value = modified_date


def getContractInfo():
    """Put contracts into a dictionary for future reference."""
    contractInfo = {}
    count = 1
    for row in range(2, sourceSheet.max_row + 1):
        contractNo = sourceSheet['A' + str(row)].value
        companyNo = sourceSheet['E' + str(row)].value
        companyRep = sourceSheet['G' + str(row)].value
        contractRep = sourceSheet['H' + str(row)].value

        # Ensure key for contract exists
        contractInfo.setdefault(contractNo, {'Companies': [companyNo],
                                             'CompanyRep': companyRep,
                                             'ContractRep': contractRep,
                                             'CompanyCount': 1}
                                )
        # Check for contract with multiple companies
        while sourceSheet['A' + str(row + count)].value is None:
            contractInfo[contractNo]['Companies'].append(sourceSheet[
                'E' + str(row + count)])
            count += 1
            contractInfo[contractNo]['CompanyCount'] += 1
    return contractInfo


def saveContractFiles(allContracts, contract, keystrokes, nonConReps, non=0):
    """Pull the files from the database."""
    count = 0
    txtPath = os.path.join('H:', os.sep, 'CONTXTFILES', contract)

    pyautogui.typewrite(Now.month)
    pyautogui.typewrite('.01.')
    pyautogui.typewrite(Now.year - 1)
    pyautogui.typewrite(['enter'])
    pyautogui.typewrite(Now.month)
    pyautogui.typewrite(Now.day)
    pyautogui.typewrite(Now.year)
    pyautogui.typewrite(['enter'])
    for items in allContracts[contract]['Companies']:
        pyautogui.typewrite(allContracts[contract]['Companies'][items - 1])
        pyautogui.typewrite(['enter'])
        pyautogui.typewrite(contract)
        pyautogui.typewrite(['enter'])
        count += 1
    pyautogui.typewrite(['enter'])
    if count > 1:
        pyautogui.typewrite(keystrokes[4])
    if non != 0:
        pyautogui.typewrite(keystrokes[6])
    else:
        pyautogui.typewrite(keystrokes[5])
    pyautogui.typewrite(['enter'])
    pyautogui.typewrite(['enter'])
    pyautogui.typewrite(txtPath)
    pyautogui.typewrite(['enter'] * 2)
    time.sleep(10)
    if allContracts[contract]['CompanyRep'] in nonConReps and non == 0:
        saveContractFiles(allContracts,
                          contract,
                          keystrokes,
                          nonConReps,
                          non=1
                          )


def menuSetup(winX, winWidth, winY, winHeight):
    """Prepare the database screen for the input."""
    pyautogui.click(winX + (winWidth / 2), winY + (winHeight / 2))
    pyautogui.typewrite(['enter'] * 5)
    for i in range(0, 4):
        pyautogui.typewrite(keystrokes[i])
        pyautogui.typewrite(['enter'])


def exitStrategy():
    """Exit the program."""
    print('All files pulled!')
    input()


"""This is the main body of the program."""
os.chdir(os.path.join('S:', os.sep, 'CSR', 'Contract Renewal Text Files'))
sourceBook = openpyxl.load_workbook(glob.glob('*.xlsx'))
sourceSheet = sourceBook.sheetnames[0]
emptyCount = 0
listOfFiles = glob.glob('*.txt')
listOfFiles_dict = {}
Now = datetime.now()

# Modification time of the files
for i in listOfFiles:
    listOfFiles_dict[i] = time.localtime(os.stat(i).st_mtime)

appWindows = []
allNames = listEmailGroups()
keystrokes = getKeystrokes()

nonConReps = allNames[0]
listToAvoid = allNames[1]
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

allContracts = getContractInfo()
# Setting current time = datetime.now().strftime('%m-%d-%Y %H:%M:%S')

# Go through, pulling text files and saving them under the contract name.
# Make sure this also pulls the non-contract items if that applies.

windowCoords = getWindow(appWindows)
winX = windowCoords[0]
winY = windowCoords[1]
winWidth = windowCoords[2] - winX
winHeight = windowCoords[3] - winY
setWinPos = windll.user32.SetWindowPos
# The -1 locks the window on top.
setWinPos(appWindows[0], -1, winX, winY, 0, 0, 0x0001)

if pullContracts.strip() == 'end':
    setWinPos(appWindows[0], 1, winX, winY, 0, 0, 0x0001)
    raise SystemExit
elif pullContracts.strip() != '':
    startContract = pullContracts.split('-')[0]
    endContract = pullContracts.split('-')[1]

    # Start interacting with the database
    menuSetup(winX, winWidth, winY, winHeight)
    for contract in range(startContract, endContract + 1):
        if contract in allContracts:
            saveContractFiles(allContracts, contract, keystrokes, nonConReps)
else:
    menuSetup(winX, winWidth, winY, winHeight)
    for contract in sorted(allContracts):
        saveContractFiles(allContracts, contract, keystrokes, nonConReps)

setWinPos(appWindows[0], 1, winX, winY, 0, 0, 0x0001)
exitStrategy()
