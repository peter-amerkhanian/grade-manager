import pyautogui
import xlwings as xw
from data_entry import copy_paste
from name_check import name_check


def main():
    pyautogui.FAILSAFE: bool = True
    while True:
        last_row: str = input("What is the last row in the excel sheet occupied by students? ")
        try:
            int(last_row)
            break
        except ValueError:
            print("Not a valid row number, please try again ")
    names, rejects, xl_reject_indices, xl_rejects_reverse = name_check(last_row)

    # begin entering data
    while True:
        columns = ['E', 'F', 'G']
        # start data entry
        input("\n\nMove your mouse to the first cell then press 'Enter'"
              " \n *Leave the computer alone after you press Enter!* "
              " \n *(Move your mouse to the upperleft hand side to cancel)*"
              "\n\n")
        pyautogui.click(pyautogui.position())
        pyautogui.hotkey('ctrl', 'a')
        for ind, col in enumerate(columns):
            # grade_dict = {}
            if ind % 2 == 0:
                # changed rounding
                xl_range = list(xw.Range(f'{col}10:{col}{last_row}').value).copy()
                xl_range_no_null = [0.00 if x == '' or x is None else x for x in xl_range]
                scores_forward = [str(format(score, '.2f')) for score in xl_range_no_null]
                for index in sorted(xl_reject_indices, reverse=True):
                    del scores_forward[index]
                copy_paste(scores_forward, names, rejects)
            else:
                xl_range = list(xw.Range(f'{col}10:{col}{last_row}').value)[::-1].copy()
                xl_range_no_null = [0.00 if x == '' or x is None else x for x in xl_range]
                scores_backward = [str(format(score, '.2f')) for score in xl_range_no_null]
                for index in sorted(xl_rejects_reverse, reverse=True):
                    del scores_backward[index]
                copy_paste(scores_backward, names[::-1], rejects, reverse=True)
        input("\n\nGrade entry complete, press enter and open up the next sheet.\n")


if __name__ == '__main__':
    try:
        main()
    except pyautogui.FailSafeException:
        print("\nProcess quit.")

