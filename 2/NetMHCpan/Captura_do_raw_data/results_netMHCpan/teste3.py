import os
import pyautogui
from time import sleep
lista = os.listdir('C:\\Users\\mauro\\Downloads')

for elem in lista:

    sleep(2)
    pyautogui.hotkey('win')
    sleep(1)
    pyautogui.typewrite(elem)
    sleep(1)
    pyautogui.press('enter')