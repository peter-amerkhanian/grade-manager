import pyautogui
import time
from typing import List
from collections import OrderedDict


def page_turn(reverse: bool) -> None:
    """determines whether to continue up or down the column"""
    if reverse:
        pyautogui.press('up')
        time.sleep(.02)
    else:
        pyautogui.press('down')
        time.sleep(.02)


def copy_paste(data: OrderedDict, student_names: list, student_rejects: list, reverse: bool = False) -> None:
    """Enters one column of grades"""
    count: int = 0
    index: int
    name: str
    student_names: List[str] = [name.strip().lower() for name in student_names]
    student_rejects: List[str] = [name.strip().lower() for name in student_rejects]
    data_list: List[str] = list(data.values())
    for index, name in enumerate(student_names):
        if name in student_rejects:
            page_turn(reverse)
        else:
            pyautogui.PAUSE = 0.1
            for _ in range(6):
                pyautogui.press('backspace')
            pyautogui.PAUSE = 0.3
            value: str = data.get(name)
            if not value:
                value: str = data_list[count]
                print(f"PLEASE REVIEW: {name.upper()} - {value}")
            # print(name, value)
            pyautogui.typewrite(value)
            count += 1
            time.sleep(.02)
            page_turn(reverse)
            time.sleep(.02)
        if index == len(student_names)-1:
            pyautogui.press('right')
            time.sleep(.02)
