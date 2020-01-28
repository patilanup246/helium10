import re
import os
import csv
import pandas as pd
import glob
import datetime
import time

wait_time = 60

dir_path = os.path.dirname(os.path.realpath(__file__))
df = pd.read_excel(dir_path + '/selection.xlsx', sheet_name='test5')

from selenium import webdriver
from selenium.webdriver.chrome.options import Options

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

import pyautogui
import time

def helium_next_login(driver, username, password):
    driver.get('http://members.helium10.com')
    driver.find_element_by_css_selector('#loginform-email').send_keys(username)
    driver.find_element_by_css_selector('#loginform-password').send_keys(password)
    driver.find_element_by_css_selector('.btn-success').click()


def get_next_login(current_username=None, current_password=None):
    credentials_df = pd.read_csv(dir_path + '/credentials.csv')
    if current_username is not None and current_password is not None:
        credentials_df[
            (credentials_df.username == current_username) &
            (credentials_df.password == current_password)
        ] = [current_username, current_password, 100, time.time()]
    credentials_df.to_csv(dir_path + '/credentials.csv', index=False)
    credentials_df.uses = credentials_df.uses.fillna(0)
    credentials_df['time_delta'] = time.time() - credentials_df.iloc[:, -1].fillna(0)
    valid_credentials_df = credentials_df[
        (credentials_df.uses < 100) | (credentials_df.time_delta > 43200)
    ]
    try:
        return valid_credentials_df.iloc[0][['username', 'password']].tolist()
    except:
        print("ran out of logins")
        exit()

options = webdriver.ChromeOptions()
options.add_extension(dir_path + '/helium10.crx')
options.add_argument("--window-size=1400,1050")

prefs = {"download.default_directory" : dir_path}
options.add_experimental_option("prefs",prefs)

def start_driver(username, password):
    driver = webdriver.Chrome(options=options, executable_path=dir_path + '/chromedriver')
    # driver = webdriver.Chrome(options=options, executable_path=dir_path + '/windows_chromedriver.exe')
    while len(driver.window_handles) == 1:
        pass
    driver.close()
    driver.switch_to_window(driver.window_handles[0])

    helium_next_login(driver, username, password)

    driver.get('https://www.amazon.com/ref=nav_logo')
    driver.get('https://www.amazon.com/gp/site-directory?ref_=nav_shopall_btn')
    return driver

current_credentials = get_next_login()
driver = start_driver(*current_credentials)

df = df[df['select']]

to_save = []
data = [] 

output_file = '/' + datetime.datetime.now().strftime('%Y%m%d')
output_file += '_final.csv'

pages = 0
for index, row in df.iterrows():
    row_pending = True
    while row_pending:
        driver.get(row['url'])

        # check if the zip code is correct
        if "90712" not in driver.find_element_by_css_selector('#glow-ingress-line2').text:
            driver.find_element_by_css_selector('#glow-ingress-line2').click()
            time.sleep(3)
            driver.find_element_by_css_selector('#GLUXZipUpdateInput').send_keys("90712\n")
            driver.get(row['url'])

        results = driver.find_elements_by_css_selector('.a-spacing-top-small > span:nth-child(1)')
        # skip if the page is blank
        if len(results) == 0:
            row_pending = False
            continue
        try:
            items = driver.find_element_by_css_selector('.a-spacing-top-small > span:nth-child(1)').text
            items = re.search('of ([0-9]+) results', items)
            if items is not None:
                items = items.group(1)
                items = int(items)
        except:
            items = 24
        if type(items) is not int:
            items = 24
        
        # click Helium10 extension button
        pyautogui.click(dir_path + '/icons/1.png')

        # click X-Ray function
        max_retries = 5
        for n_retries in range(max_retries):
            try:
                x, y = pyautogui.locateCenterOnScreen(dir_path + '/icons/2.png', confidence=0.9)
                pyautogui.click(x, y)
                break
            except:
                time.sleep(3)
        time.sleep(5)
        if driver.find_element_by_css_selector("i[class='fa fa-spinner fa-spin']").get_attribute('style') == '':
            current_credentials = get_next_login(*current_credentials)
            driver.quit()
            driver = start_driver(*current_credentials)
            continue
        else:
            row_pending = False


        # wait until sales number loads
        start_wait_time = time.time()
        current_wait_time = time.time() - start_wait_time
        while '$' not in driver.find_element_by_css_selector('#h10-bb-sales-number').text and current_wait_time <= wait_time:
            current_wait_time = time.time() - start_wait_time
            time.sleep(1.5)
            pass
            
        # load more if there are more than 24 items
        if items > 24:
            try:
                time.sleep(1.5)
                x, y = pyautogui.locateCenterOnScreen(dir_path + '/icons/more.png', confidence=0.9)
                current_total_revenue = driver.find_element_by_css_selector('#h10-bb-sales-number').text
                pyautogui.click(x, y)
                time.sleep(1.5)
                while driver.find_element_by_css_selector("i[class='fa fa-spinner fa-spin']").get_attribute('style') != 'display: none;':
                    # print("waiting to load more")
                    time.sleep(3)
                    if current_total_revenue != driver.find_element_by_css_selector('#h10-bb-sales-number').text:
                        break
            except:
                pass

        # download the items
        try:
            x, y = pyautogui.locateCenterOnScreen(dir_path + '/icons/3.png', confidence=0.9)
            pyautogui.click(x, y)
        except:
            continue
        
        start_wait_time = time.time()
        current_wait_time = time.time() - start_wait_time
        while current_wait_time <= wait_time:
            current_wait_time = time.time() - start_wait_time
            time.sleep(1.5)
            if len(glob.glob(dir_path + '/Helium*csv')) > 0:
                break

        df_tosave = pd.read_csv(glob.glob(dir_path + '/Helium*csv')[-1])

        for helium_file in glob.glob(dir_path + '/Helium*csv'):
            os.remove(helium_file)

        if os.path.exists(dir_path + output_file):
            df_final = pd.read_csv(dir_path + output_file)
            df_tosave = df_tosave.assign(**row)
            df_final = pd.concat([df_final, df_tosave], axis=0)
            df_final.to_csv(dir_path + output_file, index=False)
        else:
            df_tosave.assign(**row).to_csv(dir_path + output_file, index=False)
driver.quit()
