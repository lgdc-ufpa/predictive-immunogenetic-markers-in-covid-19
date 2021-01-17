from selenium import webdriver
driver_path = "C:\\Users\\mauro\\Desktop\\Bruno\\Automação\\configuring the envyronmet\\chromedriver\\chromedriver_win32\\chromedriver.exe"

driver = webdriver.Chrome(driver_path)
driver.get("http://www.cbs.dtu.dk/services/NetMHCpan/")

Xpath = '//option[@value="9"]'
option_9mer = driver.find_element_by_xpath(Xpath)
print(option_9mer.text)
option_9mer.click()
print()


