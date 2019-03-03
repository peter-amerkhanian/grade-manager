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
    input("...\n...\nMove your mouse to the first student's name and press Enter\n")
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
        # if sorted(name_text.split()) not in [sorted(n.split()) for n in excel_names]:
        if max([len(set(name_text.split()) & set(n.split())) for n in excel_names]) < 3:
            student_rejects.append(name_text)
            print("Name in the Ministerio file and not the Excel: ", name_text)
        student_names.append(name_text)
        pyautogui.press('down')
    # excel_rejects = [name for name in excel_names
    #                  if sorted(name.split()) not in [sorted(n.split()) for n in student_names]]
    excel_rejects = [name for name in excel_names if max([len(set(name.split()) & set(n.split())) for n in student_names]) < 3]
    if len(excel_rejects) > 0:
        print("The following names are in the Excel file and not in the Ministerio file:", "\n".join(excel_rejects))
    excel_reject_indices = [excel_names.index(name) for name in excel_rejects]
    excel_reject_indices_reverse = [excel_names[::-1].index(name) for name in excel_rejects]
    # print(excel_reject_indices)
    return student_names, student_rejects, excel_reject_indices, excel_reject_indices_reverse


def main():
    pyautogui.FAILSAFE = True
    while True:
        last_row = input("What is the last row in the excel sheet occupied by students? ")
        try:
            int(last_row)
            break
        except ValueError:
            print("Not a valid row number, please try again ")
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
                # changed rounding
                assignment = [str(format(score, '.2f')) for score in xw.Range(f'{col}11:{col}{last_row}').value]
                for index in xl_rejects:
                    assignment.pop(index)
                grade_dict['column'] = col
                grade_dict['scores'] = assignment
                copy_paste(grade_dict['scores'], names, rejects)
            else:
                assignment = [str(format(score, '.2f')) for score in list(xw.Range(f'{col}11:{col}{last_row}').value)[::-1]]
                for index in xl_rejects_reverse:
                    assignment.pop(index)
                grade_dict['column'] = col
                grade_dict['scores'] = assignment
                copy_paste(grade_dict['scores'], names[::-1], rejects, reverse=True)
        input("\n\nGrade entry complete, press enter and open up the next sheet.\n")


if __name__ == '__main__':
    main()

