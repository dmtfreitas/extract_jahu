
# coding: UTF-8

# taxpayer_folders_creator.py

import os
import sys
import shutil

CURRENT_DIR = os.path.dirname (os.path.abspath (__file__))
PARENT_DIR = os.path.abspath (os.path.join (CURRENT_DIR, '..'))
sys.path.append (PARENT_DIR)

import pdfplumber
import re
from openpyxl import Workbook
from modules.read_ctms import read_ctms
from datetime import datetime

get_cwd = os.getcwd ()

folder_path = os.path.dirname (os.path.abspath (__file__))
return_folder = os.path.abspath (os.path.join (folder_path, '..'))
pdf_folder = os.path.join (return_folder, 'JAHU', 'PDFs')

ctm_codes = read_ctms()

wb = Workbook()
ws = wb.active
line_sheet = 1

def log_error (message, log_file_path):
    
    try:
        
        with open (log_file_path, 'a', encoding='utf-8') as file_log:
            
            timestamp = datetime.now ().strftime ('%Y-%m-%d %H:%M:%S')
            file_log.write (f'[{timestamp}] {message}\n')
            
    except Exception:
        
        pass

def taxpayer_folders_creator_and_distribution ():

    jahu_folder = os.path.join (return_folder, 'JAHU')
    os.makedirs (jahu_folder, exist_ok=True)

    log_file = os.path.join (jahu_folder, 'errors_taxpayer_folders_creator.txt')

    for ctm_code in ctm_codes:

        try:
            pdf_path = os.path.join (pdf_folder, f'{ctm_code}.pdf')

            if not os.path.exists (pdf_path):
                
                log_error (f'CTM {ctm_code}: PDF NOT FOUND IN {pdf_path}', log_file)
                continue

            taxpayer = None

            with pdfplumber.open (pdf_path) as pdf:
                
                for page in pdf.pages:

                    text = page.extract_text ()

                    if not text:
                        
                        continue

                    lines = text.split ('\n')

                    for line in lines:
                        
                        pattern = re.search (r'Contribuinte:(.*?)CPF/CNPJ:', line)

                        if pattern:
                            
                            taxpayer = pattern.group (1).strip ()
                            print (taxpayer)
                            
                            ws.cell (row=line_sheet, column=1, value=taxpayer)
                            line_sheet += 1
                            
                            break

                    if taxpayer:
                        
                        break

            if not taxpayer:
                
                log_error (f'CTM {ctm_code}: NOT FOUND TAXPAYER IN PDF FILE!', log_file)
                continue

            taxpayer_folder = os.path.join (jahu_folder, taxpayer)
            os.makedirs (taxpayer_folder, exist_ok=True)

            excel_path_before = os.path.join (jahu_folder, f'{ctm_code}.xlsx')
            excel_path_after = os.path.join (taxpayer_folder, f'{ctm_code}.xlsx')

            if os.path.exists (excel_path_before):
                
                shutil.move (excel_path_before, excel_path_after)
                
            else:
                
                log_error (f'CTM {ctm_code}: FILE EXCEL NOT FOUND IN {excel_path_before}', log_file)

        except Exception as error:

            log_error (f'CTM {ctm_code}: ERRO INESPERADO - {repr (error)}', log_file)
            continue
