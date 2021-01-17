from selenium import webdriver
from openpyxl import load_workbook
from NetMHCpan.Modules import SalvarEstado
import os
import pyautogui

diretorio = "C:\\Users\\mauro\\Desktop\\Bruno\\Automação\\SCRIPTS\\NetMHCpan\\"
driver_path = "C:\\Users\\mauro\\Desktop\\Bruno\\Automação\\configuring the envyronmet\\chromedriver\\chromedriver_win32\\chromedriver.exe"

#variaveis
ultima_analise = 1
numero_analises = 0
n_alelos_remanescentes = 0
alelos = []
linha = int
final = bool
tamanho = 9

alelos_path = "C:\\Users\\mauro\\OneDrive\\Doutorado\\alelostestefinal.xlsx"
book_alelos = load_workbook(alelos_path)
sheet = book_alelos.active

epitopos_path = str
epitopos_txt = str
epitopos_path = "C:\\Users\\mauro\\OneDrive\\Doutorado\\Epitopos\\"

#Tenta abrir o arquivo dos epitopos
abrir = False
while not abrir:
    try:
        nome_epitopos = str(input("Entre com o nome do epitopo sem a extensão .txt:"))
        epitopos_txt = open(epitopos_path + nome_epitopos+'.txt', "r")
        epitopos = (epitopos_txt.readlines())
        abrir = True
    except:
        print("Nome incorreto. Insira o nome sem a extensão .txt")

# checa se há registros
try:
    file = open(diretorio + nome_epitopos + '_estado.txt', 'r')
    estados = file.readlines()
    print(estados)
    file.close() #TODO: testar estados[0], [1] e [2] e os tipos
    ultima_analise = int(estados[0])
    numero_analises = int(estados[1])
    n_alelos_remanescentes = int(estados[2]) #TODO: testar essas 3 variaveis do estado
    print("Existe um estado")
    print("ultima analise: ", ultima_analise, type(ultima_analise))
    print("numero_analises: ", numero_analises, type(numero_analises))
    print("n_alelos_remanescentes: ", n_alelos_remanescentes, type(n_alelos_remanescentes))

    # checa se a ultima_analise feita é a análise final
    if ultima_analise == int(estados[1]):
        print("Analise Final")
        final = True
    else:
        final = False

except:
    print("Não há estados")
    final = False
    linha = 0
    while (not final):
        linha = linha + 1
        print(sheet["A"+str(linha)].value)
        if str(sheet["A" + str(linha)].value) == "None":
            print("Acabou")
            final = True
            linha = linha - 1
    final = False
    print("numero de alelos: ", linha)
    print("numero de analises de 10: ", int(linha / 10))
    print("numero de alelos da ultima analise: ", linha % 10)

    #checa se todas as análises serão de 10 alelos
    if (linha % 10 == 0):
        numero_analises = int(linha / 10)

    #acrescenta uma analise a mais para os alelos remanescentes
    else:
        numero_analises = int(linha / 10) + 1
        n_alelos_remanescentes = linha % 10

    #cria o estado
    # ultima analise e numeroamalises estao com valores trocados
    file = open(diretorio + nome_epitopos + '_estado.txt', 'w')
    file.write(str(ultima_analise) + '\n' + str(numero_analises) + '\n' + str(n_alelos_remanescentes))
    file.close()

    # checa se há registros
    file = open(diretorio + nome_epitopos + '_estado.txt', 'r')
    estados = file.readlines()
    print(estados)
    estado = file.readline()
    print(estado)
    file.close()

while(ultima_analise != numero_analises):

    #netMHCpan
    print("Entra no netMHCIIpan e coloca os epitopos")
    driver = webdriver.Chrome(driver_path)
    driver.get("http://www.cbs.dtu.dk/services/NetMHCIIpan/")
    textarea = driver.find_element_by_tag_name('textarea')

    #É aqui
    #textarea.send_keys(epitopos)

    driver.find_element_by_name("length").click()
    pyautogui.press('backspace', presses=2)
    pyautogui.typewrite('9')
