from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
import pyautogui, sys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import NoSuchElementException
import os
import time

#browser = webdriver.Safari()
def webpage_initiate():
    global browser
    browser = webdriver.Chrome(ChromeDriverManager().install())
    browser.get("#")
    browser.maximize_window()
    WebDriverWait(browser, 15).until(EC.presence_of_element_located((By.ID, "wpName1")))
    browser.find_element_by_id("wpName1").send_keys("#")
    browser.find_element_by_id("wpPassword1").send_keys("#")
    browser.find_element_by_id("wpLoginAttempt").click()


def error_Handler():
        global browser
        try:
            browser.find_element_by_class_name("mw-editsection").click()
        except:
            try:
                browser.find_element_by_id("ca-view").click()
            except:
                browser.get("http://statwiki.kolobkreations.com/index.php?title=Special:UserLogin&returnto=Special:UnwatchedPages")
                browser.find_element_by_xpath("//*[@id='mw-content-text']/div/ol/li[10]/a[2]").click()
                browser.get("http://statwiki.kolobkreations.com/index.php?title=Special:UserLogin&returnto=Special:UnwatchedPages")
                WebDriverWait(browser, 15).until(EC.presence_of_element_located((By.CLASS_NAME, "mw-numlink")))
                browser.find_element_by_xpath("//*[@id='mw-content-text']/div/ol/li[10]/a[1]").click()
                error_Handler()
def link_destroy(): 
    global iCount               
    while(iCount < 500000):
        try:
            WebDriverWait(browser, 15).until(EC.presence_of_element_located((By.CLASS_NAME, "mw-numlink")))
            browser.find_element_by_xpath("//*[@id='mw-content-text']/div/ol/li[10]/a[1]").click()
            error_Handler()
        except:
            browser.close()
            break
        
        try:
            WebDriverWait(browser, 15).until(EC.presence_of_element_located((By.ID, "ca-view")))
            browser.find_element_by_id("p-cactions-label").click()
            browser.find_element_by_xpath("//*[@id='ca-delete']/a").click()
        except:
            browser.close()
            break

        try:  
            WebDriverWait(browser, 15).until(EC.presence_of_element_located((By.ID, "wpConfirmB")))
            browser.find_element_by_id("wpConfirmB").click()
        except:
            browser.close()
            break
        try:
            WebDriverWait(browser, 15).until(EC.presence_of_element_located((By.ID, "mw-returnto")))
            browser.get("http://statwiki.kolobkreations.com/index.php?title=Special:UserLogin&returnto=Special:UnwatchedPages")
        except:
            browser.close()
            break
        if(iCount % 10 == 0):
            browser.refresh()
        iCount = iCount + 1
myMax = 0
while(myMax < 1):
    iCount = 0
    time.sleep(3)
    webpage_initiate()
    link_destroy()
    
    



