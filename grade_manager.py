import time
import pyautogui
import pyperclip as pyperclip
import xlwings as xw


def copy_paste(data, reverse=False, blank=True):
    """
    data - a list
    reverse - boolean
    """
    row_count = 0
    for score in data[:-1]:
        if not blank:
            pyautogui.hotkey('ctrl', 'c')
            time.sleep(.02)
            text = str(pyperclip.paste())
            while True:
                if "0.01" in text:
                    if reverse:
                        pyautogui.press('up')
                    else:
                        pyautogui.press('down')
                        row_count += 1
                    time.sleep(.02)
                    pyautogui.hotkey('ctrl', 'c')
                    text = str(pyperclip.paste())
                else:
                    break
            pyautogui.PAUSE = 0.1
            for _ in range(len(text)):
                pyautogui.press('backspace')
        pyautogui.PAUSE = 0.3
        pyautogui.typewrite(score)
        time.sleep(.02)
        if reverse:
            pyautogui.press('up')
            time.sleep(.02)
        else:
            pyautogui.press('down')
            row_count += 1
            time.sleep(.02)
    pyautogui.PAUSE = 0.1
    for _ in range(len(text)):
        pyautogui.press('backspace')
    pyautogui.PAUSE = 0.3
    pyautogui.typewrite(data[-1])
    time.sleep(.02)
    pyautogui.press('right')
    return row_count


def cleanup():
    pyautogui.PAUSE = 0.3
    pyautogui.hotkey('ctrl', 'c')
    time.sleep(.02)
    text = str(pyperclip.paste())
    if "0.01" in text:
        for _ in range(5):
            pyautogui.PAUSE = 0.1
            for _ in range(4):
                pyautogui.press('backspace')
            pyautogui.press('left')
            time.sleep(.02)
        pyautogui.PAUSE = 0.4
        for _ in range(5):
            pyautogui.press('right')
            time.sleep(.02)
    pyautogui.press('up')


def name_check(last_row_with_names):
    input("...\n...\nMove your mouse to the first student's name and press Enter")
    excel_names = [str(score).strip() for score in xw.Range(f'B11:B{last_row_with_names}').value]
    names = []
    rejects = []
    pyautogui.click(pyautogui.position())
    while True:
        pyautogui.hotkey('ctrl', 'c')
        time.sleep(.02)
        name_text = str(pyperclip.paste()).strip()
        if len(names) > 0:
            if name_text == names[-1]:
                break
        if name_text not in excel_names:
            rejects.append(name_text)
        names.append(name_text)
        pyautogui.press('down')
        return names, rejects

# To Do!!!
# - Finish and test name name_check()
# - Fix the cleanup() function - it needs to move left
# and then run up the rows


if __name__ == '__main__':
    # init pyautogui fail safe
    pyautogui.FAILSAFE = True
    # start program
    print("...\n...\nThis program will be extracting grades from an Excel file.")
    print("Make sure that you have your desired Excel file open, and the correct sheet selected,")
    input("then press 'Enter'")
    last_row = input("...\n...\nWhat is the last row occupied by students? ")
    print("Checking that names line up")
    names, rejects = name_check(last_row)
    while True:
        blank = input("...\n...\nIs the spreadsheet you will be copying the information into currently blank? (Y/N)"
                      "\n(If you are copying into a spreadsheet app that supports overwriting cells, just select Y) ")
        if blank.lower() == "y" or blank.lower() == "n":
            break
        else:
            print('Answer not recognized. Try again')
    # check for quimestre scores
    if 'QUIM' in xw.sheets.active.name:
        columns = ['C']
    else:
        columns = ['X', 'Y', 'Z', 'AA', 'AB']
    # begin entering data
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
            grade_dict['column'] = col
            grade_dict['scores'] = assignment
            if blank == 'y':
                count = copy_paste(grade_dict['scores'])
            else:
                count = copy_paste(grade_dict['scores'], blank=False)
        else:
            assignment = [str(score) for score in list(xw.Range(f'{col}11:{col}{last_row}').value)[::-1]]
            grade_dict['column'] = col
            grade_dict['scores'] = assignment
            if blank == 'y':
                copy_paste(grade_dict['scores'], reverse=True)
            else:
                copy_paste(grade_dict['scores'], reverse=True, blank=False)
    time.sleep(.02)
    pyautogui.press('left')
    for _ in range(count):
        cleanup()


