
# coding: UTF-8

# open_jahu.py

import os
import time
import pyautogui
import openpyxl
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from modules.read_ctms import read_ctms
import pyperclip

URL = 'https://jau.portalservicos.app.br:2053/consulta-debito?menu=4'

def update_status_ctm_xlsx (ctm, status, caminho_excel):
    
    if os.path.exists(caminho_excel):
        
        df = pd.read_excel(caminho_excel)
        
    else:
        
        df = pd.DataFrame (columns=['CÓDIGO CTM', 'STATUS'])

    if ctm in df['CÓDIGO CTM'].values:
        
        df.loc[df['CÓDIGO CTM'] == ctm, 'STATUS'] = status
        
    else:
        
        novo_registro = pd.DataFrame({'CÓDIGO CTM': [ctm], 'STATUS': [status]})
        df = pd.concat([df, novo_registro], ignore_index=True)
    
    df.to_excel (caminho_excel, index=False)

def open_jahu ():

    SCRIPT_DIR = os.path.dirname (os.path.abspath(__file__))
    PROJECT_ROOT = os.path.abspath (os.path.join(SCRIPT_DIR, '..'))
    JAHU_PDFS_DIR = os.path.join (PROJECT_ROOT, 'JAHU', 'PDFs')
    STATUS_EXCEL = os.path.join (PROJECT_ROOT, 'JAHU', "CTM'S_STATUS.xlsx")
    
    if not os.path.exists (os.path.join(PROJECT_ROOT, 'JAHU')):
        
        os.makedirs (os.path.join(PROJECT_ROOT, 'JAHU'), exist_ok=True)

    ctm_lines = read_ctms ()

    for ctm_line in ctm_lines:
        
        if os.path.exists (f'{JAHU_PDFS_DIR}\\{ctm_line}.pdf'):
            
            update_status_ctm_xlsx(ctm_line, 'CONCLUÍDO', STATUS_EXCEL)
            
            continue
        
        while True:
        
            driver = None
            
            try:
                
                chrome_options = Options ()
                chrome_options.add_argument (f'--app={URL}')
                chrome_options.add_argument ('--disable-prompt-on-repost')
                chrome_options.add_argument ('--incognito')
                chrome_options.add_argument ('--disable-geolocation')
                chrome_options.add_argument ('--log-level=3')
                chrome_options.add_experimental_option ('excludeSwitches', ['enable-logging', 'enable-automation'])
                chrome_options.add_experimental_option ('useAutomationExtension', False)

                service = Service (log_output='chromedriver.log')

                driver = webdriver.Chrome (service=service, options=chrome_options)
                driver.maximize_window ()

                time.sleep (5.5)

                element = driver.find_element (By.XPATH, '//*[@id="text-opcao-pesquisa"]/div/div[1]/input')
                
                if not element:
                    
                    raise RuntimeError ('CAMPO XPATH DE PESQUISA NÃO LOCALIZADO!')
                    
                element.click ()
                time.sleep (1.5)
                pyautogui.press ('down')
                time.sleep (1.5)
                pyautogui.press ('enter')
                time.sleep (1.5)
                pyautogui.press ('tab')
                time.sleep (1.5)
                pyautogui.write (f'{ctm_line}')
                time.sleep (1.5)
                pyautogui.press ('tab')
                time.sleep (1.5)
                pyautogui.press ('enter')

                time.sleep (15.5)

                element = driver.find_element (By.XPATH, '/html/body/srv-root/layout/srv-layout-menu/div/div/srv-consulta-debitos/srv-basic-layout-base/div/div/div[2]/srv-shortcut/dx-toolbar/div/div[3]/div[2]/div/div/div/span')
                
                if not element:
                    
                    raise RuntimeError ('CAMPO XPATH DE IMPRIMIR NÃO LOCALIZADO!')
                    
                element.click ()

                time.sleep (15.5)

                pyautogui.hotkey ('ctrl', 's')
                time.sleep (5.5)

                if not os.path.exists (JAHU_PDFS_DIR):
                    
                    os.makedirs (JAHU_PDFS_DIR, exist_ok=True)

                full_path = os.path.join (JAHU_PDFS_DIR, f'{ctm_line}.pdf')

                pyperclip.copy (full_path)
                pyautogui.hotkey ('ctrl', 'v')
                time.sleep (4.5)
                pyautogui.press ('enter')

                time.sleep (15.5)
                
                pdf_path = f'{JAHU_PDFS_DIR}\\{ctm_line}.pdf'
                
                if os.path.exists (pdf_path):
                    
                    file_size = os.path.getsize (pdf_path)
                    
                    if file_size > 1024:
                        
                        try:
                            
                            with open (pdf_path, 'rb') as f:
                                
                                header = f.read (4)
                                
                                if header == b'%PDF':
                                    
                                    update_status_ctm_xlsx (ctm_line, 'CONCLUÍDO', STATUS_EXCEL)
                                    
                                else:
                                    
                                    os.remove (pdf_path)
                                    update_status_ctm_xlsx (ctm_line, 'NÃO CONCLUÍDO', STATUS_EXCEL)
                                    
                        except Exception:
                            
                            update_status_ctm_xlsx (ctm_line, 'NÃO CONCLUÍDO', STATUS_EXCEL)
                            
                    else:
                        
                        os.remove (pdf_path)
                        update_status_ctm_xlsx (ctm_line, 'NÃO CONCLUÍDO', STATUS_EXCEL)
                        
                else:
                    
                    update_status_ctm_xlsx (ctm_line, 'NÃO CONCLUÍDO', STATUS_EXCEL)

                try:
                    
                    driver.quit ()
                    
                except Exception:
                    
                    pass
                    
                break

            except Exception:
                
                update_status_ctm_xlsx (ctm_line, 'NÃO CONCLUÍDO', STATUS_EXCEL)

                try:
                    
                    if driver:
                        
                        driver.quit ()
                        
                except Exception:
                    
                    pass

                break

