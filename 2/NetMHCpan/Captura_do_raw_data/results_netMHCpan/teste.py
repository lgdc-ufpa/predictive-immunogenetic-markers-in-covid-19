from openpyxl import load_workbook
import pyautogui
from pathlib import Path
from time import sleep
import os

dir_downloads = str(Path.home()) + '\Downloads'
print(dir_downloads)

pyautogui.hotkey('win', 'e')

for i in range(7):
    pyautogui.press('tab')
    sleep(0.5)

for i in range(3):
    pyautogui.press('right')
    sleep(0.5)
pyautogui.press('enter')

for i in range(9):
    pyautogui.press('tab')
    sleep(0.5)

pyautogui.press('down')
pyautogui.press('up')
pyautogui.press('enter')

sleep(3)

pyautogui.press('left')
pyautogui.press('enter')

keyboard = ['alt', 'a', 'a', 'y', '4', 'home', 'enter', 'right', 'enter']
for elem in keyboard:
    pyautogui.press(elem)
    sleep(0.5)

