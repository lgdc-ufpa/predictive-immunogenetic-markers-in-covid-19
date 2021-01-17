from selenium import webdriver
driver_path = "C:\\Users\\mauro\\Desktop\\Bruno\\Automação\\configuring the envyronmet\\chromedriver\\chromedriver_win32\\chromedriver.exe"
driver = webdriver.Chrome(driver_path)
driver.get("http://www.cbs.dtu.dk/cgi-bin/webface2.fcgi?jobid=5CE7EC9D00007551A294FB5C&wait=20")

print("nethMHCpan: gera o excel de saida")
# Output
driver.find_elements_by_tag_name("a")[-3].click()