# -*- coding: utf-8 -*-
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
import time
import logging as log

# prepare the parameters
driverAddress = "C:/Users/xwchang/Desktop/SmartSheet/Tools/chromedriver_win32/chromedriver.exe"
phantomjs = "C:/Users/xwchang/Desktop/SmartSheet/phantomjs-2.1.1-windows/phantomjs-2.1.1-windows/bin/phantomjs"
email = "xuewen.chang@iba-group.com"
passwd = "Changwen22"
siteTag = "Guangzhou"
siteItem = "02 - Guangzhou - Test Matrix"
browser = webdriver.Chrome(driverAddress)


def getHTMLText(url):
    driver = webdriver.PhantomJS(executable_path=phantomjs)
    time.sleep(2)
    driver.get(url)
    time.sleep(2)
    return driver.page_source



def find_switch_method(holder, ty, par):
    '''
    :param holder: the element
    :param ty: type of the target element
    :param par: target element name
    :return:
    '''
    if ty == "tag":
        return holder.find_element_by_tag_name(par)
    elif ty == "class":
        return holder.find_element_by_class_name(par)
    elif ty == "css":
        return holder.find_element_by_css_selector(par)
    else:
        return None


def isElementExist(element, typ, para):
    '''
    Enhancement Exist estimate method
    :param element: the holder item
    :param tag: the item inside you want to check
    :return: boolean value
    '''
    flag = True
    try:
        el = find_switch_method(element, typ, para)
        return flag, el
    except:
        el = None
        flag = False
        print("No Element :" + para)
        return flag, el

try:
    # initialise work
    browser.get("https://app.smartsheet.com/b/home")
    input = browser.find_element_by_id("loginEmail")
    input.send_keys(email)
    input.send_keys(Keys.ENTER)
    # time.sleep(10)
    wait1 = WebDriverWait(browser, 15)
    wait1.until(EC.presence_of_element_located((By.ID, "loginPassword")))
    passfild = browser.find_element_by_id("loginPassword")
    passfild.send_keys(passwd)
    passfild.send_keys(Keys.ENTER)
    wait = WebDriverWait(browser, 15)
    wait.until(EC.presence_of_element_located((By.ID, "desktopHome")))

    time.sleep(10)
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
    # find all href
    # wait for the table list to load
    tableWait = WebDriverWait(browser, 40)
    tableWait.until(EC.presence_of_element_located((By.CLASS_NAME, "clsGridTable")))
    time.sleep(15)
    tables = browser.find_elements_by_class_name(".clsGridCursor")
    key_els = ""
    for table in tables:
        flag, tb = isElementExist(table, "class", "clsGridTable")
        if flag:
            print "got you! Item ======================= " + str(tb)
            key_els = tb.find_elements_by_tag_name("a")
    for key in key_els:
        print key.text

    time.sleep(5)

finally:
    browser.close()



