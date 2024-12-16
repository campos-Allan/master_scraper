import datetime
from datetime import timedelta
import time
from tkinter import messagebox
import pyautogui

def sap(cod, feriado, dias=0):
    now = datetime.datetime.now()
    if now.today().weekday() == 0 or feriado:
        atras = 'yes'
        if now.today().weekday() != 0:
            voltando = now - timedelta(days=1+int(dias))
            ano_atras = str(voltando.year)
            dia_atras = str(voltando.strftime('%d'))
            mes_atras = str(voltando.strftime('%m'))
        elif now.today().weekday() == 0 and feriado:
            dias = int(input('Quantos dias: '))
            voltando = now - timedelta(days=3+dias)
            ano_atras = str(voltando.year)
            dia_atras = str(voltando.strftime('%d'))
            mes_atras = str(voltando.strftime('%m'))
        else:
            voltando = now - timedelta(days=3)
            ano_atras = str(voltando.year)
            dia_atras = str(voltando.strftime('%d'))
            mes_atras = str(voltando.strftime('%m'))
    else:
        atras = 'no'
        dias = 0
        ano_atras = 0
        dia_atras = 0
        mes_atras = 0

    ontem = now - timedelta(days=1)
    ano = str(ontem.year)
    dia = str(ontem.strftime('%d'))
    mes = str(ontem.strftime('%m'))
    pyautogui.click(730, 1050)
    time.sleep(7)
    pyautogui.click(50, 135)
    pyautogui.click(50, 135)
    messagebox.showinfo(
        "Continuar", "Se tiver aparecido um aviso, feche e aperte em ok")
    time.sleep(3)
    pyautogui.click(270, 80)
    pyautogui.click(270, 105)
    pyautogui.click(270, 80)
    pyautogui.click(50, 80)
    pyautogui.click(50, 80)
    time.sleep(2)
    pyautogui.typewrite(cod)
    pyautogui.click(700, 700)
    if atras == 'yes':
        pyautogui.click(500, 300)
        time.sleep(1)
        pyautogui.typewrite(f'{dia_atras}.{mes_atras}.{ano_atras}')
        pyautogui.click(800, 300)
        time.sleep(1)
        pyautogui.typewrite(f'{dia}.{mes}.{ano}')
    else:
        pyautogui.click(500, 300)
        time.sleep(1)
        pyautogui.typewrite(f'{dia}.{mes}.{ano}')
        pyautogui.click(800, 300)
        time.sleep(1)
        pyautogui.typewrite(f'{dia}.{mes}.{ano}')

    pyautogui.click(630, 480)
    pyautogui.press('backspace', presses=12)
    pyautogui.click(37, 185)
    time.sleep(9)
    pyautogui.click(150, 185)
    time.sleep(1)
    pyautogui.click(60, 35)

    pyautogui.click(100, 130)
    time.sleep(1)
    pyautogui.click(500, 170)
    time.sleep(1)
    pyautogui.click(400, 550)
    time.sleep(2)
    pyautogui.moveTo(400, 755, duration=1)
    pyautogui.moveTo(405, 755, duration=1)
    time.sleep(1)
    pyautogui.click(400, 755)
    time.sleep(2)
    pyautogui.click(680, 720)
    time.sleep(3)
    pyautogui.click(900, 875)
    time.sleep(4)
    pyautogui.click(310, 585)
    time.sleep(3)
    pyautogui.click(760, 685)
    time.sleep(3)
    pyautogui.click(535, 505)
    time.sleep(5)
    pyautogui.getWindowsWithTitle("Planilha em Basis (1)")[0].maximize()
    time.sleep(1)
    if cod == '1120':
        pyautogui.click(270, 100)
        time.sleep(1)
        pyautogui.click(200, 170)
        time.sleep(1)
        pyautogui.click(1000, 450)
        time.sleep(1)
        pyautogui.click(1390, 960)
        time.sleep(3)
        pyautogui.moveTo(1470, 600, duration=1)
        pyautogui.mouseDown(button='right')
        pyautogui.mouseUp(button='right')
        time.sleep(1)
        pyautogui.click(1490, 620)
        time.sleep(1)
        pyautogui.click(250, 370)
        pyautogui.typewrite('DINAMICA TERMINAIS RIO VERDE S/A')
        pyautogui.hotkey('enter')
        messagebox.showinfo(
            "Sucesso", "Salve o arquivo na mesma pasta desse programa")
    else:
        pyautogui.click(270, 100)
        time.sleep(1)
        pyautogui.click(200, 170)
        time.sleep(1)
        pyautogui.click(1000, 650)
        time.sleep(1)
        pyautogui.click(1390, 960)
        time.sleep(3)
        pyautogui.click(1430, 890)
        time.sleep(1)
        pyautogui.click(1430, 660)
        time.sleep(1)
        pyautogui.moveTo(1425, 520,duration=1)
        pyautogui.click()
        time.sleep(1)
        pyautogui.click(250, 370)
        pyautogui.typewrite('SINOP')
        pyautogui.hotkey('enter')
        messagebox.showinfo(
            "Sucesso", "Salve o arquivo na mesma pasta desse programa")
polo=input('RV ou SINOP:')
dias=input('Dias de feriado: ')
if dias=='0':
    feriado=False
else:
    feriado=True
if polo == 'SINOP':
    sap('1109',feriado,dias)
elif polo == 'RV':
    sap('1120',feriado,dias)
