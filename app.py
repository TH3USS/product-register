# python -m venv .venv
# \.venv\Scripts\activate
# pip install openpyxl pyautogui

import openpyxl

workbook = openpyxl.load_workbook('venda_de_produtos.xlsx')
vendas_sheet = workbook['vendas']

for linha in vendas_sheet.iter_rows(min_row=2):
    # para pegar as cordenadas na tela: no prompt de comando 
    # pip install mouseinfo
    # python
    # from mouseinfo import mouseInfo
    # mouseInfo()

    pyautogui.click(1808, 452, duration=1.5)
    pyautogui.write(linha[0].value)
    
    pyautogui.click(1815, 476, duration=1.5)
    pyautogui.write(linha[1].value)
    
    pyautogui.click(1813, 497, duration=1.5)
    pyautogui.write(linha[2].value)
    
    pyautogui.click(1883, 532, duration=1.5)
    pyautogui.write(linha[3].value)
    
    pyautogui.click(1752, 549, duration=1.5)
    pyautogui.click(1256, 581, duration=1.5)
    
    