import time
import os
import re
import warnings
import traceback
import datetime
import shutil
from math import ceil
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.keys import Keys #pip install selenium 
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.select import Select
warnings.filterwarnings('ignore')

import undetected_chromedriver as uc #pip install undetected-chromedriver 
#from selenium.webdriver.chrome.service import Service
#driver_path = 'chromedriver.exe'
#service = Service(driver_path)






try:
    print('Reading input data       ', end = '\r')
    #pegando dados dos arquivos txt
    with open('login.txt', 'r', encoding='utf8') as file:
        login = file.read()
    with open('password.txt', 'r', encoding='utf8') as file:
        password = file.read()
    with open('timeout.txt', 'r', encoding='utf8') as file:
        timeout = file.read()
    for i in [' ', '\t', '\n']:
        login = login.replace(i, '')
        password = password.replace(i, '')
        timeout = timeout.replace(i, '')
    timeout = float(timeout)


    tipo_pje = '1'#######################################################################input('Extrair pje1 (1) ou pje2 (2)? ')




    #lendo XLSX de inputs
    for diretorio, subpastas, arquivos in os.walk('Excel'):
        for file in arquivos:
            if file.count('.xls') != 0 or file.count('.xlsx') != 0:
                arquivo_XLSX = 'Excel/' + file
    df_inputs = pd.read_excel(arquivo_XLSX)
    n_rows = len(df_inputs[df_inputs.columns[0]])
    df = df_inputs


    
    #opening the web site
    print('Opening the website       ', end = '\r')
    options = Options()   
    options.add_argument("--start-maximized")
    #options.add_argument("--incognito")
    options.add_argument("--disable-popup-blocking")

    driver = uc.Chrome(options=options)#para acessar no modo undetected
    
    driver.maximize_window()
    wait = WebDriverWait(driver, 60)
    wait_faster = WebDriverWait(driver, 5)
    wait_fast = WebDriverWait(driver, 0.1)
    

    #logging in
    driver.get('https://www.amazon.com/ap/signin?openid.pape.max_auth_age=0&openid.return_to=https%3A%2F%2Fwww.amazon.com%2F-%2Fpt%2Fcustomer-preferences%2Fedit%3Fie%3DUTF8%26preferencesReturnUrl%3D%252F-%252Fpt%252Fcustomer-preferences%252Fedit%253Fie%253DUTF8%2526preferencesReturnUrl%253D%25252F%2526ref_%253Dtopnav_lang_ais%26ref_%3Dnav_signin%26language%3Den_US&openid.identity=http%3A%2F%2Fspecs.openid.net%2Fauth%2F2.0%2Fidentifier_select&openid.assoc_handle=usflex&openid.mode=checkid_setup&openid.claimed_id=http%3A%2F%2Fspecs.openid.net%2Fauth%2F2.0%2Fidentifier_select&openid.ns=http%3A%2F%2Fspecs.openid.net%2Fauth%2F2.0')

    field = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="ap_email"]'))).send_keys(login)
    time.sleep(0.30)

    field = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="continue"]'))).click()
    time.sleep(0.30)

    field = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="ap_password"]'))).send_keys(password)
    time.sleep(0.30)

    field = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="signInSubmit"]'))).click()
    time.sleep(0.30)



    espanhol = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="icp-language-settings"]/div[3]/div/label/input')))
    driver.execute_script("arguments[0].click();", espanhol)#click falso
    time.sleep(0.30)

    espanhol = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="icp-save-button"]/span/input')))
    driver.execute_script("arguments[0].click();", espanhol)#click falso
    time.sleep(0.30)

    

    
    

    for i in range(0, n_rows):
        n_pedido = str(df_inputs[df_inputs.columns[6]][i])
        n_pedido = str(df_inputs[df_inputs.columns[6]][i])
        TRACKING = str(df_inputs[df_inputs.columns[7]][i])
        fecha = str(df_inputs[df_inputs.columns[12]][i])
        precio = str(df_inputs[df_inputs.columns[8]][i])
        Status = str(df_inputs[df_inputs.columns[20]][i])
        print(f'Processing line {i + 1}/{n_rows} - Pedido {n_pedido}')
        
        
        if Status == 'OK1':
            print('\tYa OK')
        else:
            try:
                #precio
                driver.get(f'https://www.amazon.com/gp/your-account/order-history/ref=ppx_yo2ov_dt_b_search?opt=ab&search={n_pedido}')
                time.sleep(timeout)

                try:
                    precio = wait_faster.until(EC.presence_of_element_located((By.XPATH, '//*[@id="ordersContainer"]/div[1]/div[1]/div/div/div/div[1]/div/div[2]/div[2]/span'))).text
                    precio = precio.replace(' ', '').replace('$', '').replace('.', ',')
                except:
                    precio = ''
                df.iat[i, 8] = precio
                print(f'\tPrice: {precio}')
                

                #TRACKING
                driver.get(f'https://www.amazon.com/progress-tracker/package/ref=ppx_yo_dt_b_track_package?_encoding=UTF8&itemId=rglktrjmnmnsqn&orderId={n_pedido}')
                time.sleep(timeout)

                try:
                    TRACKING = wait_faster.until(EC.presence_of_element_located((By.XPATH, '//*[@id="pt-page-container-inner"]/div[3]/div[2]/div[5]'))).text
                    TRACKING = TRACKING.split('ID de rastreo: ')[1].split('\n')[0]
                    TRACKING = TRACKING.replace(' ', '')
                except:
                    try:
                        TRACKING = wait_faster.until(EC.presence_of_element_located((By.XPATH, '//*[@id="carrierRelatedInfo-container"]/div/div[2]/div/h4'))).text
                        TRACKING = TRACKING.split('ID de rastreo: ')[1].split('\n')[0]
                        TRACKING = TRACKING.replace(' ', '')
                    except:
                        TRACKING = ''
                df.iat[i, 7] = TRACKING
                print(f'\tTracking: {TRACKING}')




                #fecha     //*[@id="pt-page-container-inner"]/div[3]/div[2]/div[5]/div[2]/section/h1
                try:                 
                    fecha = wait_faster.until(EC.presence_of_element_located((By.XPATH, '//*[@id="pt-page-container-inner"]/div[3]/div[2]/div[5]/div[1]/section/h1'))).text
                except:
                    fecha = ''
                if fecha == '':
                    try:                 
                        fecha = wait_faster.until(EC.presence_of_element_located((By.XPATH, '//*[@id="pt-page-container-inner"]/div[3]/div[2]/div[5]/div[2]/section/h1'))).text
                    except:
                        fecha = ''
                if fecha == '':
                    try:                 
                        fecha = wait_faster.until(EC.presence_of_element_located((By.XPATH, '//*[@id="tracking-events-container"]/div/div[3]/div[1]/span'))).get_attribute('innerHTML')
                    except:
                        fecha = ''
                if fecha == '':
                    try:                 
                        fecha = wait_faster.until(EC.presence_of_element_located((By.XPATH, '//*[@id="primaryStatus"]'))).text#Ahora previsto para el 2 de enero
                    except:
                        fecha = ''
                df.iat[i, 12] = fecha
                print(f'\tFecha: {fecha}')


                

                Stat = 'OK1'
            except:
                Stat = 'error'
                print(traceback.format_exc())########
                time.sleep(2)




            #results
            df.iat[i, 20] = Stat
            df.to_excel(arquivo_XLSX, 'Sheet1', index=False)    
except:
    print('\n\tSome error occurred... ')
    print(traceback.format_exc())

end = input('\n\nProgram finished! Press ENTER to close')
