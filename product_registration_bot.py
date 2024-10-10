import pandas as pd
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select

# Lê a planilha Excel
df = pd.read_excel('caminho_para_sua_planilha.xlsx') 

# Baixando sempre a versão mais atual do ChromeDriver
servico = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=servico)

# Acessa o sistema
driver.get('https://staging.connectere.agr.br/produtos/new')

# Localiza o campo de email e preenche
username = driver.find_element(By.XPATH, '/html/body/div[2]/div/div/div/div/form/div[1]/div[1]/div/input')
username.send_keys('yuri@connectere.agr.br')

# Localiza o campo de senha e preenche
password = driver.find_element(By.XPATH, '/html/body/div[2]/div/div/div/div/form/div[1]/div[2]/div/input')
password.send_keys('123456')

# Clica no botão de login
driver.find_element(By.XPATH, '/html/body/div[2]/div/div/div/div/form/div[2]/input').click()

# Preenche os campos do formulário com base nos dados da planilha
for index, row in df.iterrows():
    # Preenche o campo fabricante
    fabricante_select = Select(driver.find_element(By.XPATH, '/html/body/div[1]/section[2]/article/div[2]/div/form/div[2]/div/div[1]/div[1]/div/span/span[1]/span'))
    fabricante_select.select_by_visible_text(row['fabricante'])

    # Preenche o campo família
    familia_select = Select(driver.find_element(By.XPATH, '/html/body/div[1]/section[2]/article/div[2]/div/form/div[2]/div/div[1]/div[2]/div/span/span[1]/span'))
    familia_select.select_by_visible_text(row['familia'])

    # Preenche o campo nome
    nome_input = driver.find_element(By.XPATH, '/html/body/div[1]/section[2]/article/div[2]/div/form/div[2]/div/div[2]/div[1]/div/input')
    nome_input.send_keys(row['nome'])

    # Preenche o campo unidade de medida
    unidade_select = Select(driver.find_element(By.XPATH, '/html/body/div[1]/section[2]/article/div[2]/div/form/div[2]/div/div[3]/div[1]/div/span/span[1]/span'))
    unidade_select.select_by_visible_text(row['unidade_medida'])

    # Opcional: Clica no botão para cadastrar o produto
    driver.find_element(By.XPATH, '/html/body/div[1]/section[2]/article/div[2]/div/form/div[4]/button').click()

input("Pressione Enter para fechar o navegador...")
driver.quit()
