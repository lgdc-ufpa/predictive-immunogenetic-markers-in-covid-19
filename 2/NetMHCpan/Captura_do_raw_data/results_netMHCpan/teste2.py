import pyautogui
from pathlib import Path
from time import sleep
import clipboard

import os

dir_downloads = str(Path.home()) + '\Downloads'
print(dir_downloads)

pyautogui.hotkey('win', 'e')
sleep(5)
for i in range(6):
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
nomes = []

n = 1

for i in range(n):
    # coloca o nome do arquivo no clipboard
    pyautogui.press('f2')
    sleep(1)
    pyautogui.hotkey('ctrl' + 'c')

    # pega do clipboard
    clipboard.copy()
    print(type(clipboard.paste), clipboard.paste)
    nomes.append('{}'.format(clipboard.paste))
    # win32clipboard.CloseClipboard()

print(nomes)

# import clipboard
# import pyautogui
# import clipboard
# # clipboard.copy("this text is now in the clipboard")
# """Oi 2"""
# pyautogui.hotkey('ctrl', 'c')
# print(clipboard.paste())
