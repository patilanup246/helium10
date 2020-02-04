import re
import os
import csv
import pandas as pd
import glob
import datetime
import time

wait_time = 60

from selenium import webdriver
from selenium.webdriver.chrome.options import Options

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

import pyautogui
import time

dir_path = os.path.dirname(os.path.realpath(__file__))
options = webdriver.ChromeOptions()
options.add_extension(dir_path + '/helium10.crx')
options.add_argument("--window-size=1400,1050")

prefs = {"download.default_directory": dir_path}
options.add_experimental_option("prefs", prefs)


def helium_next_login(driver, username, password):
    driver.get('http://members.helium10.com')
    driver.find_element_by_css_selector('#loginform-email').send_keys(username)
    driver.find_element_by_css_selector('#loginform-password').send_keys(password)
    driver.find_element_by_css_selector('.btn-success').click()


def start_driver(username, password):
    # driver = webdriver.Chrome(options=options, executable_path=dir_path + '/chromedriver')
    driver = webdriver.Chrome(options=options, executable_path=dir_path + '/chromedriver.exe')
    while len(driver.window_handles) == 1:
        pass
    driver.close()
    driver.switch_to_window(driver.window_handles[0])

    helium_next_login(driver, username, password)

    driver.get('https://www.amazon.com/ref=nav_logo')
    driver.get('https://www.amazon.com/gp/site-directory?ref_=nav_shopall_btn')
    return driver


def main():
    df = pd.read_csv(dir_path + '/selection.csv', encoding="ISO-8859-1")
    df = df[df['select']]
    output_file = '/' + datetime.datetime.now().strftime("%m/%d/%Y%H:%M:%S")
    output_file += '_final.csv'
    credentials_df = pd.read_csv(dir_path + '/credentials.csv')
    for credentials_dfindex, credentials_dfrow in credentials_df.iterrows():
        perlogin = 0
        for index, row in df.iterrows():
            try:
                time.sleep(5)
                if row['select'] == True:
                    
                    if perlogin == 0:
                        driver = start_driver(credentials_dfrow.username, credentials_dfrow.password)
                    elif perlogin > 100:
                        perlogin = 0
                        break
                    row_pending = True
                    while row_pending:
                        perlogin= df.loc[credentials_df.username == credentials_dfrow.username, 'uses']
                        perlogin = perlogin + 1
                        credentials_df.loc[(credentials_df.username == credentials_dfrow.username), 'uses'] = perlogin
                        credentials_df.to_csv(dir_path + '/credentials.csv', index=False)
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
                            df.loc[(df.url == row['url']), 'error'] = "Error : No record found"
                            df.to_csv(dir_path + '/selection.csv', index=False)
                            break
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
                        try:
                            image = pyautogui.locateOnScreen(dir_path + '/icons/1.png')  # Searches for the image
                            if image == None:
                                image = pyautogui.locateOnScreen(dir_path + '/icons/1.png')
                                if image == None:
                                    image = pyautogui.locateOnScreen(dir_path + '/icons/4.png')
                                    if image == None:
                                        image = pyautogui.locateOnScreen(dir_path + '/icons/8.png')
                                        if image == None:
                                            df.loc[(df.url == row[
                                                'url']), 'error'] = "Error in Chrome Extention button click"
                                            df.to_csv(dir_path + '/selection.csv', index=False)
                                            break

                            buttonhel = pyautogui.center(image)
                            buttonhelx, buttonhely = buttonhel
                            pyautogui.click(buttonhelx, buttonhely)
                            time.sleep(10)
                        except Exception as e:
                            df.loc[(df.url == row['url']), 'error'] = "Error in Chrome Extention button click"
                            df.to_csv(dir_path + '/selection.csv', index=False)
                            time.sleep(5)
                            break

                        # click X-Ray function
                        max_retries = 5
                        for n_retries in range(max_retries):
                            try:
                                limitover = None
                                imageopen = pyautogui.locateOnScreen(
                                    dir_path + '/icons/2.png')  # Searches for the image
                                if imageopen == None:
                                    imageopen = pyautogui.locateOnScreen(dir_path + '/icons/2.png')
                                    if imageopen == None:
                                        imageopen = pyautogui.locateOnScreen(dir_path + '/icons/5.png')
                                        if imageopen == None:
                                            imageopen = pyautogui.locateOnScreen(dir_path + '/icons/7.png')
                                            if imageopen == None:
                                                imageopen = pyautogui.locateOnScreen(dir_path + '/icons/x1.png')
                                                if imageopen == None:
                                                    imageopen = pyautogui.locateOnScreen(dir_path + '/icons/x2.png')
                                                    if imageopen == None:
                                                        imageopen = pyautogui.locateOnScreen(dir_path + '/icons/x3.png')
                                                        if imageopen == None:
                                                            imageopen = pyautogui.locateOnScreen(
                                                                dir_path + '/icons/x4.png')
                                                            if imageopen == None:
                                                                imageopen = pyautogui.locateOnScreen(
                                                                    dir_path + '/icons/x4.png')
                                                                if imageopen == None:
                                                                    break

                                buttonopen = pyautogui.center(imageopen)
                                buttonopenx, buttonopeny = buttonopen
                                pyautogui.click(buttonopenx, buttonopeny)
                                time.sleep(10)
                                break
                            except:
                                df.loc[(df.url == row[
                                    'url']), 'error'] = "Error in Xray- Amazon button click"
                                df.to_csv(dir_path + '/selection.csv', index=False)
                                time.sleep(3)
                                break
                        if imageopen == None:
                            df.loc[(df.url == row[
                                'url']), 'error'] = "Error in Xray- Amazon button click"
                            df.to_csv(dir_path + '/selection.csv', index=False)
                            break

                        time.sleep(70)

                        # load more if there are more than 24 items
                        if items > 24:
                            try:
                                time.sleep(1.5)
                                try:
                                    imagemore = pyautogui.locateOnScreen(
                                        dir_path + '/icons/more.png')  # Searches for the image
                                    if imagemore == None:
                                        imagemore = pyautogui.locateOnScreen(dir_path + '/icons/more.png')
                                        if imagemore == None:
                                            imagemore = pyautogui.locateOnScreen(dir_path + '/icons/more1.png')

                                    buttonmore = pyautogui.center(imagemore)
                                    buttonmorex, buttonmorey = buttonmore
                                    pyautogui.click(buttonmorex, buttonmorey)
                                    current_total_revenue = driver.find_element_by_css_selector(
                                        '#h10-bb-sales-number').text
                                except:
                                    time.sleep(3)
                                    pass

                                time.sleep(1.5)
                                while driver.find_element_by_css_selector(
                                        "i[class='fa fa-spinner fa-spin']").get_attribute(
                                    'style') != 'display: none;':
                                    # print("waiting to load more")
                                    time.sleep(3)
                                    if current_total_revenue != driver.find_element_by_css_selector(
                                            '#h10-bb-sales-number').text:
                                        break
                            except:
                                pass

                        # download the items
                        try:
                            imagedown = pyautogui.locateOnScreen(dir_path + '/icons/3.png')  # Searches for the image
                            if imagedown == None:
                                imagedown = pyautogui.locateOnScreen(dir_path + '/icons/d1.png')
                                if imagedown == None:
                                    imagedown = pyautogui.locateOnScreen(dir_path + '/icons/6.png')
                                    if imagedown == None:
                                        imagedown = pyautogui.locateOnScreen(dir_path + '/icons/d2.png')
                                        if imagedown == None:
                                            imagedown = pyautogui.locateOnScreen(dir_path + '/icons/d3.png')
                                            if imagedown == None:
                                                imagedown = pyautogui.locateOnScreen(dir_path + '/icons/d4.png')
                                                if imagedown == None:
                                                    imagedown = pyautogui.locateOnScreen(dir_path + '/icons/d5.png')
                                                    if imagedown == None:
                                                        df.loc[(df.url == row[
                                                            'url']), 'error'] = "Error Download button click"
                                                        df.to_csv(dir_path + '/selection.csv', index=False)
                                                        break
                            buttondown = pyautogui.center(imagedown)
                            buttondownx, buttondowny = buttondown
                            pyautogui.click(buttondownx, buttondowny)
                        except:
                            df.loc[(df.url == row['url']), 'error'] = "Error Download button click"
                            df.to_csv(dir_path + '/selection.csv', index=False)
                            time.sleep(3)
                            break

                        start_wait_time = time.time()
                        current_wait_time = time.time() - start_wait_time
                        while current_wait_time <= wait_time:
                            current_wait_time = time.time() - start_wait_time
                            time.sleep(1.5)
                            if len(glob.glob(dir_path + '/Helium*csv')) > 0:
                                break
                        try:
                            df_tosave = pd.read_csv(glob.glob(dir_path + '/Helium*csv')[-1])
                            df_tosave['DownloadTime'] = datetime.datetime.now()
                            for helium_file in glob.glob(dir_path + '/Helium*csv'):
                                os.remove(helium_file)

                            if os.path.exists(dir_path + output_file):
                                df_final = pd.read_csv(dir_path + output_file, encoding="ISO-8859-1")
                                df_tosave = df_tosave.assign(**row)
                                df_final = pd.concat([df_final, df_tosave], axis=0)
                                df_final.to_csv(dir_path + output_file, encoding="utf-8", index=False)
                            else:
                                df_tosave.assign(**row).to_csv(dir_path + output_file, encoding="utf-8", index=False)
                        except Exception as e:
                            try:
                                if os.path.exists(dir_path + output_file):
                                    df_final = pd.read_csv(dir_path + output_file, encoding="utf-8")
                                    df_tosave = df_tosave.assign(**row)
                                    df_final = pd.concat([df_final, df_tosave], axis=0)
                                    df_final.to_csv(dir_path + output_file, encoding='utf-8', index=False)
                                else:
                                    df_tosave.assign(**row).to_csv(dir_path + output_file,
                                                                   encoding='utf-8', index=False)
                            except Exception as e:
                                try:
                                    if os.path.exists(dir_path + output_file):
                                        df_final = pd.read_csv(dir_path + output_file, encoding="latin-1")
                                        df_tosave = df_tosave.assign(**row)
                                        df_final = pd.concat([df_final, df_tosave], axis=0)
                                        df_final.to_csv(dir_path + output_file, index=False, encoding='latin-1')
                                    else:
                                        df_tosave.assign(**row).to_csv(dir_path + output_file, index=False,
                                                                       encoding='latin-1')
                                except Exception as e:
                                    df.loc[(df.url == row['url']), 'error'] = "Error :" + str(e)
                                    df.to_csv(dir_path + '/selection.csv', index=False)
                                    break
                        row_pending = False
                        df.loc[(df.url == row['url']), 'select'] = False
                        df.to_csv(dir_path + '/selection.csv', index=False)

                        df.loc[(df.url == row['url']), 'error'] = ""
                        df.to_csv(dir_path + '/selection.csv', index=False)
            except Exception as e:
                df.loc[(df.url == row['url']), 'error'] = "Error :" + str(e)
                df.to_csv(dir_path + '/selection.csv', index=False)
                pass


if __name__ == "__main__":
    print("script start...!!")
    try:
        first = True
        a = 1
        while a == 1:
            now = datetime.datetime.now()
            if first == True:
                main()
                first = False
            if now.hour == 7 and now.minute > 10 and now.minute < 20:
                main()
    except Exception as e:
        print("Error " + str(e))
