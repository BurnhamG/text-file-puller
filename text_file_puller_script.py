#! python3
"""This will pull text files from the system."""
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

appWindows = []


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

    for i in listOfFiles:
        listOfFiles_dict[i] = time.localtime(os.stat(i).st_mtime)

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


os.chdir('S:\CSR\Contract Renewal Text Files')
sourceBook = openpyxl.load_workbook(glob.glob('*.xlsx'))
sourceSheet = sourceBook.sheetnames[0]
emptyCount = 0
listOfFiles = glob.glob('*.txt')
listOfFiles_dict = {}

allNames = listEmailGroups()
keystrokes = getKeystrokes()

listToAvoid = allNames[3]
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
                              sourceSheet.cell(row=i, column=5).value]
                              .tm_year),
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
            # Set focus on the window and begin input
            for proc in psutil.process_iter():
                procName = proc.name()
                if re.match('*mvbase*', procName.lower()):
                    processID = proc.pid
                    processName = proc.name()
                    bringToFront = windll.user32.SetWindowPos
            win32gui.EnumWindows(window_callback, processID)
            windowLoc = win32gui.GetWindowRect(appWindows[0])
            winX = appWindows[0]
            winY = appWindows[1]
            winWidth = appWindows[2] - winX
            winHeight = appWindows[3] - winY
            # Here the -1 effectively locks the window on top. Make sure to
            # change this back!
            bringToFront(appWindows[0], -1, winX, winY, 0, 0, 0x0001)

            # Here we actually start interacting with the database
            pyautogui.click(winX + (winWidth / 2), winY + (winHeight / 2))
            # Press enter 5 times, to get back to the main menu from anywhere
            pyautogui.typewrite(['enter'] * 5)
            pyautogui.typewrite([3, 'enter', 10, 'enter', 32, 'enter', 3,
                                'enter'])
            # Check for image
            if stepRecognize('Step1.png'):
                pyautogui.typewrite([datetime.now().strftime('%m'),
                                    datetime.now.strftime('%d'),
                                    datetime.now().strftime('%y')])
# Setting current time = datetime.now().strftime('%m-%d-%Y %H:%M:%S')

# Go through, pulling text files and saving them under the contract name.
# Make sure this also pulls the non-contract items if that applies.
