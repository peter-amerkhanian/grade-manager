import pyautogui
import pyperclip as pyperclip
import time
import xlwings as xw


def compare_to_all_names(name, all_names):
    words_in_common_with_each_name = [len(set(name.split()) & set(n.split())) for n in all_names]
    return words_in_common_with_each_name


def name_check(last_row_with_names):
    input("\nMove your mouse to the first student's name and press Enter.")
    names = xw.Range(f'B10:B{last_row_with_names}').value
    excel_names = [str(name).strip() for name in names].copy()
    all_student_names = []
    ministerio_rejects = []
    pyautogui.click(pyautogui.position())
    while True:
        pyautogui.hotkey('ctrl', 'c')
        time.sleep(.02)
        ministerio_text = str(pyperclip.paste()).strip()
        if len(all_student_names) > 0:
            if ministerio_text == all_student_names[-1]:
                break
        if max([len(set(ministerio_text.split()) & set(excel_name.split())) for excel_name in excel_names]) < 2:
            ministerio_rejects.append(ministerio_text)
        all_student_names.append(ministerio_text)
        pyautogui.press('down')
    excel_rejects = [name for name in excel_names if max(compare_to_all_names(name, all_student_names)) < 2]
    print("Done\n")
    if len(ministerio_rejects) > 0:
        print("The following names are in the Ministerio app and not in the Excel file:\n", "\n".join(ministerio_rejects))
    if len(excel_rejects) > 0:
        print("The following names are in the Excel file and not in the Ministerio app:\n", "\n".join(excel_rejects))
    excel_reject_indices = [excel_names.index(name) for name in excel_rejects]
    print(excel_reject_indices)
    excel_reject_indices_reverse = [excel_names[::-1].index(name) for name in excel_rejects]
    return all_student_names, ministerio_rejects, excel_reject_indices, excel_reject_indices_reverse
