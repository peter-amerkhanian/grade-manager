import time
import pyautogui
import pyperclip as pyperclip
import xlwings as xw


def copy_paste(data, reverse=False, blank=True):
    """
    data - a list
    reverse - boolean
    """
    for score in data[:-1]:
        if not blank:
            pyautogui.hotkey('ctrl', 'c')
            time.sleep(.01)
            text = str(pyperclip.paste())
            print(text)
            while True:
                if "0.01" in text:
                    print("!!!!\n")
                    if reverse:
                        pyautogui.press('up')
                    else:
                        pyautogui.press('down')
                    pyautogui.hotkey('ctrl', 'c')
                    time.sleep(.01)
                    text = str(pyperclip.paste())
                else:
                    break
            pyautogui.PAUSE = 0.1
            for _ in range(len(text)):
                pyautogui.press('backspace')
        pyautogui.PAUSE = 0.3
        pyautogui.typewrite(score)
        time.sleep(.01)
        if reverse:
            pyautogui.press('up')
            time.sleep(.01)
        else:
            pyautogui.press('down')
            time.sleep(.01)
    pyautogui.PAUSE = 0.1
    for _ in range(len(text)):
        pyautogui.press('backspace')
    pyautogui.PAUSE = 0.3
    pyautogui.typewrite(data[-1])
    print("***", data[-1])
    time.sleep(.02)
    pyautogui.press('right')


if __name__ == '__main__':
    # init pyautogui fail safe
    pyautogui.FAILSAFE = True
    # start program
    print("...\n...\nThis program will be extracting grades from an Excel file.")
    print("Make sure that you have your desired Excel file open, and the correct sheet selected,")
    input("then press 'Enter'")
    last_row = input("...\n...\nWhat is the last row occupied by students? ")
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
                copy_paste(grade_dict['scores'])
            else:
                copy_paste(grade_dict['scores'], blank=False)
        else:
            assignment = [str(score) for score in list(xw.Range(f'{col}11:{col}{last_row}').value)[::-1]]
            grade_dict['column'] = col
            grade_dict['scores'] = assignment
            if blank == 'y':
                copy_paste(grade_dict['scores'], reverse=True)
            else:
                copy_paste(grade_dict['scores'], reverse=True, blank=False)
