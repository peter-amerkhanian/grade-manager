import pyautogui
import pyperclip as pyperclip
import time
import xlwings as xw
from typing import List, Tuple


def compare_to_all_names(name: str, all_names: List[str]) -> List[int]:
    words_in_common_with_each_name: List[int] = [len(set(name.split()) & set(n.split())) for n in all_names]
    return words_in_common_with_each_name


def name_check(last_row_with_names: str) -> Tuple[List[str], List[str], List[int], List[int]]:
    input("\nMove your mouse to the first student's name and press Enter.")
    names: List[str] = xw.Range(f'B10:B{last_row_with_names}').value
    excel_names: List[str] = [str(name).strip() for name in names].copy()
    all_student_names: List[str] = []
    minister_rejects: List[str] = []
    pyautogui.click(pyautogui.position())
    while True:
        pyautogui.hotkey('ctrl', 'c')
        time.sleep(.02)
        minister_text: str = str(pyperclip.paste()).strip()
        if len(all_student_names) > 0:
            if minister_text == all_student_names[-1]:
                break
        if max([len(set(minister_text.split()) & set(excel_name.split())) for excel_name in excel_names]) < 2:
            minister_rejects.append(minister_text)
        all_student_names.append(minister_text)
        pyautogui.press('down')
    excel_rejects: List[str] = [name for name in excel_names if max(compare_to_all_names(name, all_student_names)) < 2]
    print("Done\n")
    if len(minister_rejects) > 0:
        print("The following names are in the Ministerio app and not in the Excel file:\n", "\n".join(minister_rejects))
    if len(excel_rejects) > 0:
        print("The following names are in the Excel file and not in the Ministerio app:\n", "\n".join(excel_rejects))
    excel_reject_indices: List[int] = [excel_names.index(name) for name in excel_rejects]
    print(excel_reject_indices)
    excel_reject_indices_reverse: List[int] = [excel_names[::-1].index(name) for name in excel_rejects]
    return all_student_names, minister_rejects, excel_reject_indices, excel_reject_indices_reverse

