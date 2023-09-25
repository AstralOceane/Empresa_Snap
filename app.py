from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from time import sleep
import openpyxl
import os

# Inicializar o WebDriver do Chrome
driver = webdriver.Chrome()
driver.get('https://app.localo.com/paywall')
driver.set_window_size(1920, 1080)
sleep(15)

# Digita o email
email = 'seu-email'
campo_email = driver.find_element(By.XPATH, '//*[@id="root"]/div/div/div/div[2]/div/form/div[1]/input')
campo_email.send_keys(email)

# Digita a senha
senha = 'sua-senha'
campo_senha = driver.find_element(By.XPATH, '//*[@id="root"]/div/div/div/div[2]/div/form/div[2]/span/input')
campo_senha.send_keys(senha)

# Clicar em conectar
botao_conectar = driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div/div[2]/div/form/button')
botao_conectar.click()
sleep(10)

# Pesquisar o nome da empresa
excel_file_path = 'dados.xlsx'
workbook = openpyxl.load_workbook(excel_file_path)
worksheet = workbook.active

# Ler o nome da empresa da planilha
nome_empresa = worksheet['A2'].value  # Supondo que o nome da empresa está na célula A1
campo_pesquisa = driver.find_element(By.XPATH, '//*[@id="root"]/div/div[1]/div[2]/div/div[2]/div/div/div/div[2]/span/input')
campo_pesquisa.send_keys(nome_empresa)
campo_pesquisa.send_keys(Keys.RETURN)
sleep(10)

# Clique no resultado
resultado = driver.find_element(By.XPATH, '//*[@id="root"]/div/div[1]/div[2]/div/div[2]/div/div/div/div[2]/div/div[2]/div/div/h5')
resultado.click()
sleep(10)

# Pesquisar Palavra Chave
palavra_chave = worksheet['B2'].value
campo_palavra_chave = driver.find_element(By.XPATH, '//*[@id="root"]/div/div[1]/div[2]/div/div[2]/div/div/div/div[2]/div/div[1]/span/input')
campo_palavra_chave.send_keys(palavra_chave)
campo_palavra_chave.send_keys(Keys.RETURN)
sleep(10)

# Rolar a página para baixo
driver.execute_script("window.scrollBy(0, 100);")

# Tirar um print da página
screenshot_path = "screenshot.png"
driver.save_screenshot(screenshot_path)

# Criar uma pasta com o nome da empresa (se não existir)
if not os.path.exists(nome_empresa):
    os.mkdir(nome_empresa)

# Mover o print para a pasta
os.rename(screenshot_path, os.path.join(nome_empresa, screenshot_path))

# Renomear o print com o nome da empresa
nome_arquivo = nome_empresa + ".png"
os.rename(os.path.join(nome_empresa, screenshot_path), os.path.join(nome_empresa, nome_arquivo))

# Fechar o navegador
driver.quit()