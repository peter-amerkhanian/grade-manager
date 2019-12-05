import pyautogui
import xlwings as xw
from components.data_entry import copy_paste
from components.name_check import name_check
from typing import List, Optional
import collections


def main() -> None:
    pyautogui.FAILSAFE: bool = True
    # Check for excel being open
    try:
        print(f"Hola, actualmente está ingresando las calificaciones de {xw.books.active.name}.")
    except AttributeError:
        print('ERROR: No hay ningún archivo de Excel abierto actualmente. ' +
              'Abra un archivo de Excel y vaya a una página de "actas" antes de ejecutar el programa')
        return None
    excel_names: List[Optional[str, int, None]] = xw.Range("B10:B100").value.copy()
    last_row: None = None
    for x in range(len(excel_names)):
        if excel_names[x] == 0.0 and excel_names[x - 1] == 0.0:
            last_row: int = x - 2 + 10
            break
    if not last_row:
        while True:
            last_row: str = input(
                "Por favor, introduzca manualmente la última fila en la hoja de Excel ocupada por los estudiantes.")
            try:
                last_row: int = int(last_row)
                break
            except ValueError:
                print("ERROR: No es un número de fila válido, por favor inténtelo de nuevo.")
    names, rejects, xl_reject_indices, xl_rejects_reverse = name_check(last_row)
    excel_names: List[str] = [str(name).strip().lower() for name in excel_names[:last_row - 9]]
    for index in sorted(xl_reject_indices, reverse=True):
        del excel_names[index]
    excel_names_reverse: List[str] = excel_names[::-1]
    # Get parcial column
    columns_to_parciales = {'1': 'E', '2': 'F', '3': 'G', 'examen': 'I'}
    while True:
        parcial = input("\nPASO 2: Cuál parcial va a entrar: 1, 2, 3 o examen? ")
        try:
            columns = [columns_to_parciales[parcial]]
            break
        except KeyError:
            print("ERROR: Ingrese 1, 2, 3 o examen")
    # begin entering data
    columns_to_parciales = {1:'E', 2:'F', 3:'G'}
    # start data entry
    input("\nPASO 3:\nMueva el ratón a la primera celda y luego presione 'Intro/Enter'"
          " \nDeje la computadora en paz después de pulsar Intro/Enter! "
          " \n(Mueva el ratón hacia el lado superior izquierdo para cancelar)")
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
    print("\nIngreso de parcial está completo.")


if __name__ == '__main__':
    try:
        main()
    except pyautogui.FailSafeException:
        print("\nProceso parado.")
    input("\nPulse cualquier tecla para salir.")

