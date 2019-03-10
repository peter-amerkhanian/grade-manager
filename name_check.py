import pyautogui
import pyperclip as pyperclip
import time


def name_check(last_row_with_names):
    input("...\n...\nMove your mouse to the first student's name and press Enter\n")
    excel_names = [str(score).strip() for score in xw.Range(f'B11:B{last_row_with_names}').value].copy()
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
        if max([len(set(name_text.split()) & set(n.split())) for n in excel_names]) < 3:
            student_rejects.append(name_text)
            print("Name in the Ministerio file and not the Excel: ", name_text)
        student_names.append(name_text)
        pyautogui.press('down')
    excel_rejects = [name for name in excel_names if max([len(set(name.split()) & set(n.split())) for n in student_names]) < 3]
    if len(excel_rejects) > 0:
        print("The following names are in the Excel file and not in the Ministerio file:", "\n".join(excel_rejects))
    excel_reject_indices = [excel_names.index(name) for name in excel_rejects]
    excel_reject_indices_reverse = [excel_names[::-1].index(name) for name in excel_rejects]
    return student_names, student_rejects, excel_reject_indices, excel_reject_indices_reverse