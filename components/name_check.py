import pyautogui
import pyperclip as pyperclip
import time
import xlwings as xw
from typing import List, Tuple


def compare_to_minister_names(excel_name: str, minister_names: List[str]) -> List[int]:
    """Compares one student name from the excel list to every
    other student name in the minister list. Every int in the
    list represents how many words the input name has in common
     with a given name."""
    words_in_common_with_each_name: List[int] = [len(set(excel_name.split()) & set(n.split())) for n in minister_names]
    return words_in_common_with_each_name


def name_check(last_row_with_names: str) -> Tuple[List[str], List[str], List[int], List[int]]:
    input("\nPASO 1: Mueva el ratón al nombre del primer alumno y pulse Intro/Enter.")
    names: List[str] = xw.Range(f'B10:B{last_row_with_names}').value
    excel_names: List[str] = [str(name).strip() for name in names].copy()
    minister_names: List[str] = []
    minister_rejects: List[str] = []
    pyautogui.click(pyautogui.position())
    while True:
        pyautogui.hotkey('ctrl', 'c')
        time.sleep(.02)
        minister_text: str = str(pyperclip.paste()).strip()
        if len(minister_names) > 0:
            if minister_text == minister_names[-1]:
                break
        if max([len(set(minister_text.split()) & set(excel_name.split())) for excel_name in excel_names]) < 2:
            minister_rejects.append(minister_text)
        minister_names.append(minister_text)
        pyautogui.press('down')
    excel_rejects: List[str] = [name for name in excel_names if max(compare_to_minister_names(name, minister_names)) < 2]
    print("Listo.")
    if len(minister_rejects) > 0:
        print("Los siguientes nombres se encuentran en la aplicación Ministerio y no en el archivo de Excel:\n", "\n".join(minister_rejects))
    if len(excel_rejects) > 0:
        print("Los siguientes nombres se encuentran en el archivo de Excel y no en la aplicación Ministerio:\n", "\n".join(excel_rejects))
    excel_reject_indices: List[int] = [excel_names.index(name) for name in excel_rejects]
    excel_reject_indices_reverse: List[int] = [excel_names[::-1].index(name) for name in excel_rejects]
    return minister_names, minister_rejects, excel_reject_indices, excel_reject_indices_reverse

