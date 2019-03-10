#!/usr/bin/env python3

import pyautogui
import xlwings as xw
from data_entry import copy_paste
from name_check import name_check


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
                xl_range = list(xw.Range(f'{col}11:{col}{last_row}').value).copy()
                xl_range_no_null = [0.00 if x == '' or x is None else x for x in xl_range]
                scores = [str(format(score, '.2f')) for score in xl_range_no_null]
                for index in xl_rejects:
                    scores.pop(index)
                copy_paste(scores, names, rejects)
            else:
                xl_range = list(xw.Range(f'{col}11:{col}{last_row}').value)[::-1].copy()
                xl_range_no_null = [0.00 if x == '' or x is None else x for x in xl_range]
                scores = [str(format(score, '.2f')) for score in xl_range_no_null]
                for index in xl_rejects_reverse:
                    scores.pop(index)
                copy_paste(scores, names[::-1], rejects, reverse=True)
        input("\n\nGrade entry complete, press enter and open up the next sheet.\n")


if __name__ == '__main__':
    main()

