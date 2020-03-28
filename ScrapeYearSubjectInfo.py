##What this program does:
#1. Logs into Scopus through Western account
#2. Performs an 'advanced query search' for a particular year between 1861-2020
#3. Records the number of conference papers and articles that were published in that year
#4. 


import os
import os.path
from os import path
import time
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from pynput.mouse import Button, Controller
import pandas as pd
from pathlib import Path
import shutil
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
import csv
global driver

def get_random_ua():
    random_ua = ''
    ua_file = r"C:\Users\trevo\Desktop\Work Life\Non-Academic Projects\LinkedInInnovationArticle\ua_list.txt"
    try:
        with open(ua_file) as f:
            lines = f.readlines()
        if len(lines) > 0:
            prng = np.random.RandomState()
            index = prng.permutation(len(lines) - 1)
            idx = np.asarray(index, dtype=np.integer)[0]
            random_proxy = lines[int(idx)]
    except Exception as ex:
        print('Exception in random_ua')
        print(str(ex))
    finally:
        return random_ua

chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument('--user-agent=' + str(get_random_ua()))
driver = webdriver.Chrome("/Users/trevo/downloads/chromedriver", chrome_options=chrome_options)


def SignIntoScopus():
    url = 'https://www.scopus.com/home.uri'
    global driver
    driver.get(url)

    driver.get('https://www.scopus.com/signin.uri')

    email = 'tsmit256@uwo.ca'
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.NAME, 'pf.username')))
    driver.find_element_by_name('pf.username').send_keys(email)
    driver.find_element_by_name('action').click()
    
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.NAME, 'action')))
    driver.find_element_by_name('action').click()
    try:
        driver.find_element_by_name('action').click()
        driver.find_element_by_name('action').click()
        driver.find_element_by_name('action').click()
    except:
        pass
    #on western website
    userID = 'tsmit256'
    password = ''
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.NAME, 'j_username')))
    driver.find_element_by_name('j_username').send_keys(userID)
    driver.find_element_by_name('j_password').send_keys(password)
    
    driver.find_element_by_name('j_password').send_keys(Keys.ENTER)

    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.ID, 'advSearchLink')))
    driver.find_element_by_id('advSearchLink').click()

    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.ID, 'searchfield')))    
    driver.find_element_by_id('searchfield').clear()
    driver.find_element_by_id('searchfield').send_keys("DOCTYPE ( ar )  OR  DOCTYPE ( cp )")

    el = driver.find_element_by_xpath('//*[@id="advSearch"]/span[1]')
    action = webdriver.common.action_chains.ActionChains(driver)
    action.move_to_element_with_offset(el, 5, 5)
    action.click()
    action.perform()

def SearchScopus():
    global driver
    global df
    global rowCounter
    global Counter
    global year

    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.ID, 'editAuthSearch')))
    driver.find_element_by_id('editAuthSearch').click()

    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.ID, 'searchfield')))
    driver.find_element_by_id('searchfield').clear()
    driver.find_element_by_id('searchfield').send_keys("DOCTYPE ( ar )  OR  DOCTYPE ( cp )  AND  ( LIMIT-TO ( PUBYEAR ,  " + str(year[Counter][0]) + " ) )")
    
    el = driver.find_element_by_xpath('//*[@id="advSearch"]/span[1]')
    action = webdriver.common.action_chains.ActionChains(driver)
    action.move_to_element_with_offset(el, 5, 5)
    action.click()
    action.perform()

    try:
        WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.ID, 'viewAllLink_SUBJAREA')))
    except:
        e = driver.find_element_by_xpath('/html/body/div[1]/div/div[1]/header/div[2]/a')
        action = webdriver.common.action_chains.ActionChains(driver)
        action.move_to_element_with_offset(e, 5, 5)
        action.click()
        action.perform()

        WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.ID, 'advSearchLink')))
        driver.find_element_by_id('advSearchLink').click()

        WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.ID, 'searchfield')))
        driver.find_element_by_id('searchfield').clear()
        driver.find_element_by_id('searchfield').send_keys("DOCTYPE ( ar )  OR  DOCTYPE ( cp )  AND  ( LIMIT-TO ( PUBYEAR ,  " + str(year) + " ) )")

        el = driver.find_element_by_xpath('//*[@id="advSearch"]/span[1]')
        action = webdriver.common.action_chains.ActionChains(driver)
        action.move_to_element_with_offset(el, 5, 5)
        action.click()
        action.perform()
        
    driver.find_element_by_id('viewAllLink_SUBJAREA').click()
    x = year[Counter]
    df.loc[x[1], "Year"] = str(year[Counter][0])

    ##Record numbers for each subject
    ul = 1
    li = 1
    doubleBreak = False
    while ul < 4:
        while li < 11:
            try:
                WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div/div[1]/div[1]/div/div[3]/form/div[4]/div[1]/div/div/div/div[3]/div/div[2]/div[2]/div[1]/div[3]/div[4]/div[4]/div/div[3]/div/div/div/div[2]/div/ul[' + str(ul) + ']/li[' + str(li) + ']/label/span')))
            except:
                driver.get(driver.current_url)
                WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.ID, 'viewAllLink_SUBJAREA')))
                driver.find_element_by_id('viewAllLink_SUBJAREA').click()
                WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div/div[1]/div[1]/div/div[3]/form/div[4]/div[1]/div/div/div/div[3]/div/div[2]/div[2]/div[1]/div[3]/div[4]/div[4]/div/div[3]/div/div/div/div[2]/div/ul[' + str(ul) + ']/li[' + str(li) + ']/label/span')))

            subj = driver.find_element_by_xpath('/html/body/div[1]/div/div[1]/div[1]/div/div[3]/form/div[4]/div[1]/div/div/div/div[3]/div/div[2]/div[2]/div[1]/div[3]/div[4]/div[4]/div/div[3]/div/div/div/div[2]/div/ul[' + str(ul) + ']/li[' + str(li) + ']/label/span').text
            subjNUMBER = driver.find_element_by_xpath('/html/body/div[1]/div/div[1]/div[1]/div/div[3]/form/div[4]/div[1]/div/div/div/div[3]/div/div[2]/div[2]/div[1]/div[3]/div[4]/div[4]/div/div[3]/div/div/div/div[2]/div/ul[' + str(ul) + ']/li[' + str(li) + ']/button/span[1]/span[2]').text
            
            df.loc[x[1], subj] = subjNUMBER

            try:
                if(li < 10):
                    driver.find_element_by_xpath('/html/body/div[1]/div/div[1]/div[1]/div/div[3]/form/div[4]/div[1]/div/div/div/div[3]/div/div[2]/div[2]/div[1]/div[3]/div[4]/div[4]/div/div[3]/div/div/div/div[2]/div/ul[' + str(ul) + ']/li[' + str(li + 1) + ']/label/span').text
                else:
                    driver.find_element_by_xpath('/html/body/div[1]/div/div[1]/div[1]/div/div[3]/form/div[4]/div[1]/div/div/div/div[3]/div/div[2]/div[2]/div[1]/div[3]/div[4]/div[4]/div/div[3]/div/div/div/div[2]/div/ul[' + str(ul + 1) + ']/li[' + str(1) + ']/label/span').text
            except:
                doubleBreak = True
                break
            else:
                li += 1

        if doubleBreak:
            break
        else:
            ul += 1
            li = 1


def Main():
    global rowCounter

    
    global year
    global df
    global driver
    global Counter
    year = [ [1905,115]]
    Counter = 0

    df = pd.read_excel(r"C:\Users\trevo\Desktop\Work Life\Non-Academic Projects\LinkedInInnovationArticle\YearSubjectPublishings.xlsx")

    SignIntoScopus()
    
    while Counter < 1:
        SearchScopus()
        #click the esc button to get off subject pop-up page
        webdriver.ActionChains(driver).send_keys(Keys.ESCAPE).perform()
        Counter +=1
            
    df.to_excel(r'\Users\trevo\Desktop\Work Life\Non-Academic Projects\LinkedInInnovationArticle\YearSubjectPublishings.xlsx', index=False, engine='xlsxwriter')

Main()
