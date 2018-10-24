# -*- coding: utf-8 -*-
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
import time
import logging as log

# prepare the parameters
driverAddress = "E:/SoftwareExtral/chromedriver_win32/chromedriver.exe"
email = "xuewen.chang@iba-group.com"
passwd = "Changwen22"
siteTag = "Guangzhou"
siteItem = "02 - Guangzhou - Test Matrix"
browser = webdriver.Chrome(driverAddress)


def isElementExist_tag(element, tag):
    '''
    :param element:
    :return:
    '''
    flag = True
    try:
        el = element.find_element_by_tag_name(tag)
        return flag, el
    except:
        el = None
        flag = False
        print("No Element :" + tag)
        return flag, el

try:
    # initialise work
    browser.get("https://app.smartsheet.com/b/home")
    input = browser.find_element_by_id("loginEmail")
    input.send_keys(email)
    input.send_keys(Keys.ENTER)
    # time.sleep(10)
    wait1 = WebDriverWait(browser, 10)
    wait1.until(EC.presence_of_element_located((By.ID, "loginPassword")))
    passfild = browser.find_element_by_id("loginPassword")
    passfild.send_keys(passwd)
    passfild.send_keys(Keys.ENTER)
    wait = WebDriverWait(browser, 10)
    wait.until(EC.presence_of_element_located((By.ID, "desktopHome")))

    time.sleep(5)
    # goto the Home tab
    browser.find_element_by_class_name("clsDesktopTabCaption").click()
    print "click the item := Home"
    time.sleep(5)
    # after that, goto the site
    siteTags = browser.find_elements_by_class_name("clsTreeNodeTitle")
    for t in siteTags:
        if str(t.text).startswith(siteTag):
            print "click the item := " + t.text
            t.click()
            break
    time.sleep(5)
    # then we goto the Test Matrix
    siteItems = browser.find_elements_by_class_name("clsCellWrapper")
    for i in siteItems:
        if str(i.text) == siteItem:
            print "click the item := " + i.text
            i.click()
            break
    browser.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    elems = browser.find_elements_by_css_selector("div>a")
    for elem in elems:
        print elem.get_attribute("href")



    time.sleep(10)
finally:
    browser.close()



