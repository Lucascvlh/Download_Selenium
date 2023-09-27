from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
from dotenv import load_dotenv

from pyautogui import press, write, click
import pandas as pd
import numpy as np
import time
import os

load_dotenv()

data_inicial = input('Data inicial: ')
data_final = input('Data Final: ')

driver = webdriver.Chrome()
wait = WebDriverWait(driver, 30)

driver.maximize_window()

driver.get("https://magalu.brainlaw.com.br/Account/Login?ReturnUrl=%2fHome")
driver.find_element(By.XPATH, '//*[@id="Email"]').send_keys(os.getenv('LOGIN'))
driver.find_element(By.XPATH, '//*[@id="Senha"]').send_keys(os.getenv('PASSWORD') + Keys.ENTER)

driver.get('https://magalu.brainlaw.com.br/reports/RelCumprimentos.aspx')
original_window = driver.current_window_handle

driver.find_element(By.XPATH, '//*[@id="TextBoxDtInicial"]').send_keys(data_inicial)
driver.find_element(By.XPATH, '//*[@id="TextBoxDtFinal"]').send_keys(data_final)

driver.find_element(By.XPATH, '//*[@id="ContentPlaceHolder1_btnPesquisar"]').click()

wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="ContentPlaceHolder1_ASPxgvPrazos_DXFREditorcol58_I"]')))

caminho = 'comprovantes.xlsx'
planilha = pd.read_excel(caminho, sheet_name='Planilha1', usecols=[0,1,2,3,4,5], engine='openpyxl')

def atualizar_plan(message):
    planilha.at[linha,'COMPROVANTE'] = message
    planilha.to_excel(caminho, sheet_name='Planilha1', index=False)

for linha in range(len(planilha)):
    if np.isnan(planilha.at[linha,'Nº PO']):
        continue
    Po = str(planilha.at[linha,'Nº PO']).split('.')[0]
    ValorPago = format(planilha.at[linha,'DÉBITO'],'.2f').replace('.',',')
    
    driver.find_element(By.XPATH, '//*[@id="ContentPlaceHolder1_ASPxgvPrazos_DXFREditorcol58_I"]').clear()
    wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="ContentPlaceHolder1_ASPxgvPrazos_DXDataRow1"]/td[1]/a')))

    po_elemento = driver.find_element(By.XPATH, '//*[@id="ContentPlaceHolder1_ASPxgvPrazos_DXFREditorcol58_I"]')
    po_elemento.send_keys(Po)
    receivedPo = driver.find_element(By.XPATH, '//*[@id="ContentPlaceHolder1_ASPxgvPrazos_DXDataRow0"]/td[49]').text
    while str(Po) != str(receivedPo):
        if str(receivedPo) != '':
            time.sleep(1)
        else:
            print(f'PO {Po} não encontrada.')
            atualizar_plan(f'PO {Po} não encontrada.')
            continue
        receivedPo = driver.find_element(By.XPATH, '//*[@id="ContentPlaceHolder1_ASPxgvPrazos_DXDataRow0"]/td[49]').text

    valor_total_elemento = driver.find_element(By.XPATH, '//*[@id="ContentPlaceHolder1_ASPxgvPrazos_DXDataRow0"]/td[17]').text
    if  valor_total_elemento == ValorPago:
        try:
            # Aguardar a visibilidade do comprovante
            wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="ContentPlaceHolder1_ASPxgvPrazos_cell0_66_tbComprovante_0"]/tbody/tr/td/a/img')))
        except Exception as e:
            print(f'Exceção capturada: {str(e)}')
            print(f'Comprovante da PO {Po} não encontrado')
            atualizar_plan(f'Comprovante da PO {Po} não encontrado')
            continue
    else:
        print('Valores divergentes, analisar.')
        atualizar_plan('Valores divergentes, analisar.')
        continue
    comprovante_elemento = driver.find_element(By.XPATH, '//*[@id="ContentPlaceHolder1_ASPxgvPrazos_cell0_66_tbComprovante_0"]/tbody/tr/td/a/img')
    comprovante_elemento.click()

    wait.until(EC.number_of_windows_to_be(2))

    for window_handle in driver.window_handles:
        if window_handle != original_window:
            driver.switch_to.window(window_handle)
            break

    wait.until(EC.url_contains('https://magalu.brainlaw.com.br/api/processo/documento'))

    driver.execute_script("window.print();")
    time.sleep(5)

    position = driver.get_window_position()
    x = position['x']
    y = position['y']
    click(x, y)
    press('Enter')
    time.sleep(5)
    write(Po)
    press('Enter')
    time.sleep(1)
    driver.close()
    driver.switch_to.window(original_window)
    print(f'Comprovante da PO {Po} pronto.')
    atualizar_plan(f'Comprovante da PO {Po} pronto.')
print("Saindo...")
time.sleep(2)
driver.quit()
