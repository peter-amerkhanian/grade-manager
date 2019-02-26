#!/usr/bin/env python3

import time
import pyautogui
import pyperclip as pyperclip
import xlwings as xw


def page_turn(reverse):
    if reverse:
        pyautogui.press('up')
        time.sleep(.02)
    else:
        pyautogui.press('down')
        time.sleep(.02)


def copy_paste(data, student_names, student_rejects, reverse=False):
    """
    data - a list
    reverse - boolean
    """
    count = 0
    for index, name in enumerate(student_names):
        if name in student_rejects:
            page_turn(reverse)
        else:
            pyautogui.PAUSE = 0.1
            for _ in range(6):
                pyautogui.press('backspace')
            pyautogui.PAUSE = 0.3
            print(name, data[count])
            pyautogui.typewrite(data[count])
            count += 1
            time.sleep(.02)
            if index == len(student_names)-1:
                pyautogui.press('right')
                time.sleep(.02)
            else:
                page_turn(reverse)
                time.sleep(.02)


def name_check(last_row_with_names):
    input("...\n...\nMove your mouse to the first student's name and press Enter")
    excel_names = [str(score).strip() for score in xw.Range(f'B11:B{last_row_with_names}').value]
    student_names = []
    student_rejects = []
    pyautogui.click(pyautogui.position())
    while True:
        pyautogui.hotkey('ctrl', 'c')
        time.sleep(.02)
        name_text = str(pyperclip.paste()).strip()
        if len(student_names) > 0:
            if name_text == student_names[-1]:
                break
        if sorted(name_text.split()) not in [sorted(n.split()) for n in excel_names]:
            student_rejects.append(name_text)
        student_names.append(name_text)
        pyautogui.press('down')
    print(student_names)
    excel_rejects = [name for name in excel_names
                     if sorted(name.split()) not in [sorted(n.split()) for n in student_names]]
    excel_reject_indices = [excel_names.index(name) for name in excel_rejects]
    excel_reject_indices_reverse = [excel_names[::-1].index(name) for name in excel_rejects]
    print(excel_reject_indices)
    return student_names, student_rejects, excel_reject_indices, excel_reject_indices_reverse


if __name__ == '__main__':
    pyautogui.FAILSAFE = True
    print("...\n...\nThis program will be extracting grades from an Excel file.")
    print("Make sure that you have your desired Excel file open, and the correct sheet selected,")
    input("then press 'Enter'")
    last_row = input("...\n...\nWhat is the last row occupied by students? ")
    print("Checking that names line up")
    names, rejects, xl_rejects, xl_rejects_reverse = name_check(last_row)

    # begin entering data
    while True:

        # check for quimestre scores
        if 'QUIM' in xw.sheets.active.name:
            columns = ['C']
        else:
            columns = ['X', 'Y', 'Z', 'AA', 'AB']

        # start data entry
        input("...\n...\nMove your mouse to the first cell then press 'Enter'"
              " \n ***Leave the computer alone after you press Enter!*** "
              " \n ***Move your mouse to the upper left hand\n corner of the screen to stop the process!***"
              "...\n...\n")

        pyautogui.click(pyautogui.position())
        pyautogui.hotkey('ctrl', 'a')
        for ind, col in enumerate(columns):
            grade_dict = {}
            if ind % 2 == 0:
                assignment = [str(score) for score in xw.Range(f'{col}11:{col}{last_row}').value]
                for index in xl_rejects:
                    assignment.pop(index)
                grade_dict['column'] = col
                grade_dict['scores'] = assignment
                copy_paste(grade_dict['scores'], names, rejects)
            else:
                assignment = [str(score) for score in list(xw.Range(f'{col}11:{col}{last_row}').value)[::-1]]
                for index in xl_rejects_reverse:
                    assignment.pop(index)
                grade_dict['column'] = col
                grade_dict['scores'] = assignment
                copy_paste(grade_dict['scores'], names[::-1], rejects, reverse=True)
        input("Grade entry complete, press enter and open up the next sheet.")


