from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException

from pyautogui import press, write, click
from datetime import datetime
import time
import requests
import pandas as pd
import numpy as np

# Variáveis
data_atual = datetime.now()
data_atual_formatada = data_atual.strftime("%d/%m/%Y")
data_inicial = datetime(data_atual.year, data_atual.month, 1)
data_inicial_formatado = data_inicial.strftime("%d/%m/%Y")
tempo_maximo = 10
email = 'natalia.barbosa@magazineluiza.com.br'
password = '*natalia1904'
urlLoanding = 'https://magalu.brainlaw.com.br/DXR.axd?r=0_2658-TvT8l'

# Inicializar o driver do Selenium
driver = webdriver.Chrome()
wait = WebDriverWait(driver, 10)  # Tempo limite de espera em segundos

# Maximizar a janela do navegador
driver.maximize_window()

# Fazer login
driver.get("https://magalu.brainlaw.com.br/Account/Login?ReturnUrl=%2fHome")
driver.find_element(By.XPATH, '//*[@id="Email"]').send_keys(email)
driver.find_element(By.XPATH, '//*[@id="Senha"]').send_keys(password + Keys.ENTER)

# Acessar a página de relatórios
urlCumprimentos = 'https://magalu.brainlaw.com.br/reports/RelCumprimentos.aspx'
driver.get(urlCumprimentos)

# Preencher as datas iniciais e finais
driver.find_element(By.XPATH, '//*[@id="TextBoxDtInicial"]').send_keys(data_inicial_formatado)
driver.find_element(By.XPATH, '//*[@id="TextBoxDtFinal"]').send_keys(data_atual_formatada)

# Clicar no botão de pesquisa
driver.find_element(By.XPATH, '//*[@id="ContentPlaceHolder1_btnPesquisar"]').click()

# Aguardar a visibilidade do elemento PO
wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="ContentPlaceHolder1_ASPxgvPrazos_DXFREditorcol58_I"]')))

# Tempo inicial para controle de tempo máximo
inicio = time.time()

# Abrir a planilha comprovantes.xlsx
caminho = 'comprovantes.xlsx'
comprovantes_atualizados = []
planilha = pd.read_excel(caminho, sheet_name='Planilha1', usecols=[0,1,2,3,4,5], engine='openpyxl')

for linha in range(len(planilha)):
    # Ler a coluna E (PO) e D (Valor Pago)
    if np.isnan(planilha.at[linha,'Nº PO']):
        continue
    Po = str(planilha.at[linha,'Nº PO']).split('.')[0]
    ValorPago = format(planilha.at[linha,'DÉBITO'],'.2f').replace('.',',')

    while True:
        # Calcular o tempo decorrido
        tempo_decorrido = time.time() - inicio

        # Limpando os campos
        time.sleep(1)
        driver.find_element(By.XPATH, '//*[@id="ContentPlaceHolder1_ASPxgvPrazos_DXFREditorcol58_I"]').send_keys((Keys.BACKSPACE * 7))
        time.sleep(5)
        driver.find_element(By.XPATH, '//*[@id="ContentPlaceHolder1_ASPxgvPrazos_DXFREditorcol16_I"]').send_keys((Keys.BACKSPACE * 10))
        time.sleep(5)
        
        # Preencher a PO
        po_elemento = driver.find_element(By.XPATH, '//*[@id="ContentPlaceHolder1_ASPxgvPrazos_DXFREditorcol58_I"]')
        po_elemento.send_keys(Po)
        time.sleep(3)

        # Armazenar a janela original
        original_window = driver.current_window_handle

        # Verificar o status da URL
        response = requests.get(urlLoanding)
        if response.status_code == 200:
            time.sleep(3)

            # Obter o valor total
            valor_total_elemento = driver.find_element(By.XPATH, '//*[@id="ContentPlaceHolder1_ASPxgvPrazos_DXFooterRow"]/td[17]')
            valor_total = valor_total_elemento.text

            if valor_total == '0,00':
                print(f'PO {Po} não encontrada')
                comprovantes_atualizados.append(f'PO {Po} não encontrada')
                break
            valor_total = valor_total.replace('.','')
            if valor_total != ValorPago:
                valor_pago_elemento = driver.find_element(By.XPATH, '//*[@id="ContentPlaceHolder1_ASPxgvPrazos_DXFREditorcol16_I"]')
                valor_pago_elemento.send_keys(ValorPago)
                time.sleep(3)

                # Atualizar o valor total
                valor_total = valor_total_elemento.text

                # Verificar se o valor pago ainda é diferente do valor total
                if ValorPago != valor_total:
                    print('Valores diferentes do informado.')
                    comprovantes_atualizados.append('Valores diferentes do informado.')
                    break

            try:
                # Aguardar a visibilidade do comprovante
                wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="ContentPlaceHolder1_ASPxgvPrazos_cell0_66_tbComprovante_0"]/tbody/tr/td/a/img')))
            except NoSuchElementException:
                print(f'Comprovante da PO {Po} não encontrado')
                comprovantes_atualizados.append(f'Comprovante da PO {Po} não encontrado')
                break

            # Clicar no comprovante
            comprovante_elemento = driver.find_element(By.XPATH, '//*[@id="ContentPlaceHolder1_ASPxgvPrazos_cell0_66_tbComprovante_0"]/tbody/tr/td/a/img')
            comprovante_elemento.click()

            # Aguardar a abertura da nova janela
            wait.until(EC.number_of_windows_to_be(2))

            # Alternar para a nova janela
            for window_handle in driver.window_handles:
                if window_handle != original_window:
                    driver.switch_to.window(window_handle)
                    break

            # Aguardar a URL conter o trecho esperado
            wait.until(EC.url_contains('https://magalu.brainlaw.com.br/api/processo/documento'))

            # Executar o script para imprimir a página
            driver.execute_script("window.print();")
            time.sleep(5)

            # Obter a posição da janela
            position = driver.get_window_position()
            x = position['x']
            y = position['y']

            # Simular o pressionamento das teclas
            click(x, y)
            press('Enter')
            time.sleep(5)
            write(Po)
            press('Enter')
            time.sleep(1)
            driver.close()
            driver.switch_to.window(original_window)
            print(f'Comprovante da PO {Po} pronto.')
            comprovantes_atualizados.append(f'Comprovante da PO {Po} pronto.')
            break

        # Verificar se o tempo máximo foi excedido
        if tempo_decorrido >= tempo_maximo:
            print("Tempo limite excedido.")
            break

# Adicionar os resultados atualizados ao DataFrame
planilha['COMPROVANTE'] = comprovantes_atualizados

# Salvar a nova planilha
planilha.to_excel('comprovantes_atualizados.xlsx', index=False)

print("Saindo...")
time.sleep(10)
driver.quit()