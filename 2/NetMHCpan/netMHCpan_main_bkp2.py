from selenium import webdriver
from openpyxl import load_workbook
driver_path = "C:\\Users\\mauro\\Desktop\\Bruno\\Automação\\configuring the envyronmet\\chromedriver\\chromedriver_win32\\chromedriver.exe"

#estado
epitopos_path = str
nome_epitopos_txt = str
alelos_path = "C:\\Users\\mauro\\OneDrive\\Doutorado\\alelostestefinal.xlsx"
epitopos_path = "C:\\Users\\mauro\\OneDrive\\Doutorado\\Epitopos\\"
nome_peptideos_txt = open(epitopos_path + "VARV epitopes"+'.txt', "r")

epitopos = (nome_peptideos_txt.readlines())
driver = webdriver.Chrome(driver_path)

driver.get("http://www.cbs.dtu.dk/services/NetMHCpan/")
textarea = driver.find_element_by_tag_name('textarea')
textarea.send_keys(epitopos)

alelos_path = "C:\\Users\\mauro\\OneDrive\\Doutorado\\alelostestefinal.xlsx"
book_alelos = load_workbook(alelos_path)
sheet = book_alelos.active

length = 10
alelos = []
for i in range(length):
    print(i+1, sheet["A"+str(i+1)].value)
    if not i == 9:
        alelos.append(sheet["A"+str(i+1)].value + ",")
    else:
        alelos.append(sheet["A"+str(i+1)].value)
print(alelos)
alele_area = driver.find_element_by_name("allele")
alele_area.send_keys(alelos)

driver.find_element_by_name("BApred").click()
driver.find_element_by_name("sort").click()
driver.find_element_by_name("xlsdump").click()

#Submit
inputList = driver.find_elements_by_tag_name("input")
inputList[-2].click()

#Output
link_list = driver.find_elements_by_tag_name("a")
link_list[-3].click()









