import pyautogui
import time


def page_turn(reverse: bool) -> None:
    """determines whether to continue up or down the column"""
    if reverse:
        pyautogui.press('up')
        time.sleep(.02)
    else:
        pyautogui.press('down')
        time.sleep(.02)


def copy_paste(data: list, student_names: list, student_rejects: list, reverse: bool = False):
    """Enters one column of grades"""
    count: int = 0
    print(student_names)
    print(len(data))
    index: int
    name: str
    for index, name in enumerate(student_names):
        print(index)
        print(len(student_names))
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
            page_turn(reverse)
            time.sleep(.02)
        if index == len(student_names)-1:
            pyautogui.press('right')
            time.sleep(.02)
