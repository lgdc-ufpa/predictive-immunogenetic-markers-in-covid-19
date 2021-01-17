import pyautogui
import os
from os import listdir
from time import sleep
from NetMHCpan.Captura_do_raw_data.results_netMHCpan.Conversao import Conversao

lista = os.listdir('C:\\Users\\mauro\\Downloads')
xlsx = []

conversao = Conversao()

for elem in lista:
    if 'xls.xlsx' in elem:
        xlsx.append(elem[:-5])

for elem in lista:
    if not (elem in xlsx or '.xlsx' in elem):
        conversao.xls_open(xls='C:\\Users\\mauro\\Downloads\\' + elem)
        conversao.xls_convert()