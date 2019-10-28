import pyautogui
import xlwings as xw
from data_entry import copy_paste
from name_check import name_check
from typing import List, Optional
import collections


def main() -> None:
    pyautogui.FAILSAFE: bool = True
    print(f"Entering grades for {xw.books.active.name}")
    excel_names: List[Optional[str, int, None]] = xw.Range("B10:B100").value.copy()
    last_row: None = None
    for x in range(len(excel_names)):
        if excel_names[x] == 0.0 and excel_names[x - 1] == 0.0:
            last_row: int = x - 2 + 10
            break
    if not last_row:
        while True:
            last_row: str = input("Please manually enter the last row in the excel sheet occupied by students ")
            try:
                last_row: int = int(last_row)
                break
            except ValueError:
                print("Not a valid row number, please try again ")
    names, rejects, xl_reject_indices, xl_rejects_reverse = name_check(last_row)
    excel_names: List[str] = [str(name).strip().lower() for name in excel_names[:last_row - 9]]
    for index in sorted(xl_reject_indices, reverse=True):
        del excel_names[index]
    excel_names_reverse: List[str] = excel_names[::-1]
    # begin entering data
    columns = ['E', 'F', 'G']
    # start data entry
    input("\n\nMove your mouse to the first cell then press 'Enter'"
          " \n *Leave the computer alone after you press Enter!* "
          " \n *(Move your mouse to the upperleft hand side to cancel)*"
          "\n\n")
    input()
    pyautogui.click(pyautogui.position())
    pyautogui.hotkey('ctrl', 'a')
    for ind, col in enumerate(columns):
        if ind % 2 == 0:
            # changed rounding
            xl_range = list(xw.Range(f'{col}10:{col}{last_row}').value).copy()
            xl_range_no_null = [0.00 if x == '' or x is None else x for x in xl_range]
            scores_forward = [str(format(score, '.2f')) for score in xl_range_no_null]
            for index in sorted(xl_reject_indices, reverse=True):
                del scores_forward[index]
            grade_dict = collections.OrderedDict(zip(excel_names, scores_forward))
            copy_paste(grade_dict, names, rejects)
        else:
            xl_range = list(xw.Range(f'{col}10:{col}{last_row}').value)[::-1].copy()
            xl_range_no_null = [0.00 if x == '' or x is None else x for x in xl_range]
            scores_backward = [str(format(score, '.2f')) for score in xl_range_no_null]
            for index in sorted(xl_rejects_reverse, reverse=True):
                del scores_backward[index]
            grade_dict = collections.OrderedDict(zip(excel_names_reverse, scores_backward))
            copy_paste(grade_dict, names[::-1], rejects, reverse=True)
    test_column = "I"
    pyautogui.press('right')
    pyautogui.press('right')
    xl_range = list(xw.Range(f'{test_column}10:{test_column}{last_row}').value)[::-1].copy()
    xl_range_no_null = [0.00 if x == '' or x is None else x for x in xl_range]
    scores_backward = [str(format(score, '.2f')) for score in xl_range_no_null]
    excel_names_reverse: List[str] = excel_names[::-1]
    for index in sorted(xl_rejects_reverse, reverse=True):
        del scores_backward[index]
    grade_dict = collections.OrderedDict(zip(excel_names_reverse, scores_backward))
    copy_paste(grade_dict, names[::-1], rejects, reverse=True)
    print("\n\nGrade entry complete.\n")


if __name__ == '__main__':
    try:
        main()
    except pyautogui.FailSafeException:
        print("\nProcess quit.")

