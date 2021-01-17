from openpyxl import load_workbook
# from netMHCpan.Modules.modules import getMinimoRnaking_sb_wb as getMinimoRnaking_sb_wb
# from netMHCpan.Modules.modules import getHLAs as getHLAs

from NetMHCpan.Captura_do_raw_data.Modules.modules import getHLAs as getHLAs
from NetMHCpan.Captura_do_raw_data.Modules.modules import getMinimoRnaking_sb_wb as getMinimoRnaking_sb_wb
import pyautogui

# global epitopo
epitopo = str
# excel_path = "C:\\Users\\Bruno Conde\\Documents\\LGHM\\Automação\\SCRIPTS\\netMHCpan\\results_netMHCpan\\"
excel_path = "C:\\Users\\mauro\\Desktop\\Bruno\\Automação\\SCRIPTS\\NetMHCpan\\Captura_do_raw_data\\results_netMHCpan\\"
hla_list = []
results = [] # [hla, rankingMinimo, binders, strongBinders, weakBinders]

abriu = False
while not abriu:
    try:
        n_excelFiles = int(input("Entre com o numero de arquivos excel:" + '\n'))
        epitopo = input("Entre com o nome do epitopo (sem a extensao .txt):" + '\n')
        abriu = True
        print("n_excelFiles: ", n_excelFiles)
    except:
        print("Entre com um número inteiro")

for index in range(n_excelFiles):
    print("Abrindo o arquivo " + str(index + 1))
    book = load_workbook(excel_path + str(index + 1) + ".xlsx")
    sheet = book.active
    hla_list = getHLAs(sheet)
    print("Obtendo o resultdado dos hlaS: ", hla_list)
    getMinimoRnaking_sb_wb(sheet, hla_list, epitopo, excel_path)



