from selenium import webdriver
from time import sleep
driver = webdriver.Firefox()
driver.get('www.google.com')
sleep(10)