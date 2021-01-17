import pyautogui
from pathlib import Path
from time import sleep
import clipboard
import os
from os import listdir


class Conversao():
    def __init__(self):
        self._dirDownload = str(Path.home()) + '\Downloads'
        self._xls_names = []

    @property
    def dirDownloads(self):
        return self._dirDownload

    @dirDownloads.setter
    def dirDownloads(self, value):
        self._dirDownloads = value

    @property
    def xls_names(self):
        return self._dirDownloads

    @xls_names.setter
    def xls_names(self, value):
        self._xls_names = value

    def xls_open(self, xls):

        sleep(5)
        pyautogui.press('win')
        sleep(5)
        pyautogui.typewrite(xls)
        sleep(8)
        pyautogui.press('enter')
        sleep(3)


    def xls_convert(self):
        pyautogui.press('left')
        pyautogui.press('enter')
        sleep(8)

        keyboard = ['alt', 'a', 'a', 'y', '4', 'home', 'enter', 'right', 'enter']
        for elem in keyboard:
            pyautogui.press(elem)
            sleep(2)

        sleep(6)
        pyautogui.hotkey('alt', 'f4')



