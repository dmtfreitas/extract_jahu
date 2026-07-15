
# coding: UTF-8

# read_ctms.py

import os

def read_ctms ():

    SCRIPT_DIR = os.path.dirname (os.path.abspath (__file__))

    PROJECT_ROOT = os.path.abspath (os.path.join (SCRIPT_DIR, '..'))

    ctm_codes = os.path.join (PROJECT_ROOT, 'ctm_codes.txt')

    with open (ctm_codes, 'r', encoding='utf-8') as ctm_codes_xlsx:
        
        ctm_lines = ctm_codes_xlsx.readlines ()

    ctm_lines = [line.strip () for line in ctm_lines if line.strip ()]

    return ctm_lines

