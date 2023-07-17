import os
import zipfile
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import NoAlertPresentException
import shutil
import pandas as pd
import numpy as np
import requests
from bs4 import BeautifulSoup
from soup2dict import convert
import datetime
import re
import time




def get_chromedriver(chromedriver_path=None, use_proxy=False, user_agent=None,
                    PROXY_HOST=None, PROXY_PORT=None, PROXY_USER=None, PROXY_PASS=None, download_folder=None,
                    return_options_only=False):

    manifest_json = """
    {
        "version": "1.0.0",
        "manifest_version": 2,
        "name": "Chrome Proxy",
        "permissions": [
            "proxy",
            "tabs",
            "unlimitedStorage",
            "storage",
            "<all_urls>",
            "webRequest",
            "webRequestBlocking"
        ],
        "background": {
            "scripts": ["background.js"]
        },
        "minimum_chrome_version":"22.0.0"
    }
    """

    background_js = """
    var config = {
            mode: "fixed_servers",
            rules: {
            singleProxy: {
                scheme: "http",
                host: "%s",
                port: parseInt(%s)
            },
            bypassList: ["localhost"]
            }
        };

    chrome.proxy.settings.set({value: config, scope: "regular"}, function() {});

    function callbackFn(details) {
        return {
            authCredentials: {
                username: "%s",
                password: "%s"
            }
        };
    }

    chrome.webRequest.onAuthRequired.addListener(
                callbackFn,
                {urls: ["<all_urls>"]},
                ['blocking']
    );
    """ % (PROXY_HOST, PROXY_PORT, PROXY_USER, PROXY_PASS)

    chrome_options = webdriver.ChromeOptions()
    # allow multiple download
    prefs_experim = {'profile.default_content_setting_values.automatic_downloads': 1}
    
    if use_proxy:
        pluginfile = 'proxy_auth_plugin.zip'

        with zipfile.ZipFile(pluginfile, 'w') as zp:
            zp.writestr("manifest.json", manifest_json)
            zp.writestr("background.js", background_js)
        chrome_options.add_extension(pluginfile)
    if user_agent:
        chrome_options.add_argument('--user-agent=%s' % user_agent)
    if download_folder:
        prefs_experim["download.default_directory"] = download_folder

    chrome_options.add_experimental_option("prefs", prefs_experim)

    if return_options_only:
        return chrome_options
    else:
        driver = webdriver.Chrome(
            service=Service(chromedriver_path), # executable_path=chromedriver_path,
            options=chrome_options)
        
        return driver

    
def get_seduta(CHROMEDRIVER_PATH, GRADUATION_DATE, ESSE3_URL, ROOT_URL, config):
    
    driver=get_chromedriver(chromedriver_path=CHROMEDRIVER_PATH)
    driver.get(ESSE3_URL)
    driver.maximize_window()
    time.sleep(2)

    wait=WebDriverWait(driver, 10)
    
    # close cookies
    try:
        cookie=wait.until(
            EC.presence_of_element_located((By.XPATH, '/html/body/div[4]/div/div[1]/div/div[2]/button[2]')))
        cookie.click()
    except:
        print('\n## Error for cookies')
        raise
    # login button
    try:
        login=wait.until(
            EC.presence_of_element_located((By.XPATH, "/html/body/div[2]/div/div/main/div[2]/div[2]/a[4]")))
        login.click()
    except:
        print('\n## Error for login button')
        raise
    # enter credentials
    try:
        cf=wait.until(
            EC.presence_of_element_located((By.XPATH, "/html/body/div/div[2]/div[2]/div[1]/div/div/div[2]/form/div[1]/input")))
        cf.send_keys(config.Codice_Fiscale)
        pwd=wait.until(
            EC.presence_of_element_located((By.XPATH, "/html/body/div/div[2]/div[2]/div[1]/div/div/div[2]/form/div[2]/input")))
        pwd.send_keys(config.Password)
        time.sleep(1)
        login=wait.until(
            EC.presence_of_element_located((By.XPATH, "/html/body/div/div[2]/div[2]/div[1]/div/div/div[2]/form/div[3]/button")))
        login.click()
    except:
        print('\n## Error for credentials')
        raise
    # select Docente
    try:
        time.sleep(2)
        docente=wait.until(
            EC.presence_of_element_located((By.XPATH, "//a[contains(text(),'Accedi come DOCENTE')]")))
        docente.click()
    except:
        print('\n## Error for Docente')
        raise
    # open seduta
    try:
        menu=wait.until(
            EC.presence_of_element_located((By.XPATH, "/html/body/header/div/div/a[2]")))
        menu.click()
        time.sleep(1)
        commissioni=wait.until(
                EC.presence_of_element_located((By.XPATH, "/html/body/header/div/nav/div/div[2]/ul/li[4]")))
        commissioni.click()
        time.sleep(1)
        titolo=wait.until(
                EC.presence_of_element_located((By.XPATH, "/html/body/header/div/nav/div/div[2]/ul/li[4]/ul/li[1]")))
        titolo.click()
        time.sleep(1)
    #     seduta=wait.until(
    #             EC.presence_of_element_located((By.XPATH, "/html/body/div[2]/div/div/main/div[3]/div/div/div[2]/table/tbody/tr/td[4]/a/img")))
    #     seduta.click()
        seduta=wait.until(
                        EC.presence_of_element_located((By.XPATH, "/html/body/div[2]/div/div/main/div[3]/div/div/div[2]/table")))
        seduta=convert(BeautifulSoup(seduta.get_attribute('outerHTML'), 'html.parser'))
        for sed in seduta['table'][0]['tbody']:
            if GRADUATION_DATE in sed['#text']:
                link=seduta['table'][0]['tbody'][0]['tr'][0]['td'][3]['a'][0]['@href']
                driver.get(ROOT_URL + link)
    except:
        print('\n## Error for seduta')
        raise
        
    return driver