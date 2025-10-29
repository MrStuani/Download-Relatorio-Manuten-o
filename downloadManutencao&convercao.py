from datetime import datetime , timedelta
import os
import time
import shutil
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import requests
import json
import openpyxl
import pyotp
from io import BytesIO
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from gspread_dataframe import set_with_dataframe

# Detectar o diretório de download
if os.name == 'nt':  # Windows
    download_dir = os.path.join(os.path.expanduser('~'), 'Downloads')
else:  # MacOS, Linux
    download_dir = os.path.join(os.path.expanduser('~'), 'Downloads')

# Pegar os dados necessários para login no Google Sheets
def get_credentials_from_pastebin(urls):
    for url in urls:
        response = requests.get(url)
        if response.status_code == 200:
            credentials_json = response.text
            credentials_dict = json.loads(credentials_json)
            return credentials_dict
        else:
            print(f"Falha ao acessar o Pastebin: {response.status_code} URL: {url}")
    return None

pastebin_urls = ["URL PASTEBIN COM CREDENCIAIS GOOGLE SHEETS"]

# Obter as credenciais
credentials_dict = get_credentials_from_pastebin(pastebin_urls)

# Nome da planilha
spreadsheet_name = 'NOME DA PLANILHA GOOGLE SHEETS'

# Autenticação e inicialização da planilha
def authenticate_google_sheets(credentials_dict, spreadsheet_name):
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    credentials = ServiceAccountCredentials.from_json_keyfile_dict(credentials_dict, scope)
    client = gspread.authorize(credentials)
    spreadsheet = client.open(spreadsheet_name)
    return spreadsheet

# Ler dados de um arquivo Excel
def read_excel(file_path):
    return pd.read_csv(file_path)

# Carregar dados em uma aba específica do Google Sheets
def upload_to_google_sheets(spreadsheet, sheet_name, dataframe):
    worksheet = spreadsheet.worksheet(sheet_name)
    worksheet.clear()
    set_with_dataframe(worksheet, dataframe)

# Autenticar e acessar a planilha
spreadsheet = authenticate_google_sheets(credentials_dict, spreadsheet_name)

# Configurar as opções do Chrome
chrome_options = Options()


# Seu segredo base32
secret = "SUA CHAVE TOTP PARA ACESSO A PLATAFORMA AQUI "
# Cria um objeto TOTP
totp = pyotp.TOTP(secret)
# Gera o código de 6 dígitos (válido por 30 segundos)
codigo = totp.now()
print("Código atual:", codigo)



# Inicializar o WebDriver
driver = webdriver.Chrome()

try:
    # Navegar para a página de login
    driver.get("SITE DE LOGIN DA PLATAFORMA AQUI")

    # Esperar a página carregar e fechar o primeiro modal
    WebDriverWait(driver, 20).until(
        EC.visibility_of_element_located((By.XPATH, "/html/body/div[2]/div[3]/div/div/div[2]/div/div/div/button"))
    ).click()

    # Localizar e preencher os campos de login
    WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.ID, "usuario"))).send_keys("SEU USUÁRIO AQUI")
    senha = WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.ID, "senha")))
    senha.send_keys("SUA SENHA AQUI")
    WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="autenticacao_fma"]'))).send_keys(codigo)
    time.sleep(int(2))
    senha.send_keys(Keys.RETURN)

    # Esperar e fechar o segundo modal se aparecer
    try:
        WebDriverWait(driver, 20).until(
            EC.visibility_of_element_located((By.XPATH, "/html/body/div[2]/div[7]/div/div/div[1]/button"))
        ).click()
    except:
        pass  # Caso o modal não apareça

    # Fechar qualquer modal adicional
    while True:
        try:
            WebDriverWait(driver, 5).until(
                EC.visibility_of_element_located((By.XPATH, "//div[@class='modal fade bs-example-modal-lg show']//button[contains(@class, 'close')]"))
            ).click()
        except:
            break  # Sai do loop quando não houver mais modais

    # Navegar até a seção de relatórios e clicar nos elementos apropriados
    WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/nav/div/ul/li[11]/a"))
    ).click()
    WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/nav/div/ul/li[11]/ul/li[2]/a"))
    ).click()

    # Preencher as datas do contrato
    WebDriverWait(driver, 20).until(
        EC.visibility_of_element_located(
            (By.XPATH, "/html/body/table/tbody/tr[3]/td/form/div[1]/fieldset/table/tbody/tr/td/div/div[1]/table/tbody/tr[2]/td/div/table[1]/tbody/tr[2]/td/div/table/tbody/tr[37]/td[2]/div/input[1]")
        )
    ).send_keys("01/01/2000")
    tomorrow = (datetime.today() + timedelta(days=1)).strftime('%d.%m.%Y')
    WebDriverWait(driver, 20).until(
        EC.visibility_of_element_located(
            (By.XPATH, "/html/body/table/tbody/tr[3]/td/form/div[1]/fieldset/table/tbody/tr/td/div/div[1]/table/tbody/tr[2]/td/div/table[1]/tbody/tr[2]/td/div/table/tbody/tr[37]/td[2]/div/input[2]")
        )
    ).send_keys(tomorrow)

    # Selecionar o layout do relatório
    layout_relatorio = WebDriverWait(driver, 20).until(
        EC.visibility_of_element_located(
            (By.XPATH, "/html/body/table/tbody/tr[3]/td/form/div[1]/fieldset/table/tbody/tr/td/div/span[5]/table/tbody/tr[2]/td/div/table[1]/tbody/tr/td[2]/select")
        )
    )
    layout_relatorio.send_keys("RELATÓRIO DIÁRIO DE MANUTENÇÃO")
    layout_relatorio.send_keys(Keys.RETURN)
    
    
    time.sleep(3)
    # CLICK no contrato
    WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable(
            (By.XPATH, "//*[@id='chkDataContrato']")
        )
    ).click()


    time.sleep(2)
    # Selecionar o tipo de relatório
    WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable(
            (By.XPATH, "/html/body/table/tbody/tr[3]/td/form/div[1]/fieldset/table/tbody/tr/td/div/div[8]/table/tbody/tr[2]/td/div/table/tbody/tr[4]/td[2]/div/input")
        )
    ).click()

    # Clicar no botão para gerar o relatório
    WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.XPATH, "/html/body/table/tbody/tr[3]/td/form/div[3]/input"))
    ).click()
    time.sleep(4)

    # Esperar o download completar
    today = datetime.today().strftime('%d.%m.%Y')
    old_filename = os.path.join(download_dir, 'relatorio.xls')
    new_filename = os.path.join(download_dir, f'RELATÓRIO MANUTENÇÃO {today}.xls')

    # Esperar até que o arquivo exista e renomeá-lo
    while not os.path.exists(old_filename):
        time.sleep(1)

    shutil.move(old_filename, new_filename)

finally:
    # Fechar o navegador
    driver.quit()

# Define o diretório de entrada como o diretório de download detectado
input_file = new_filename
output_dir = os.path.join(download_dir, 'relatorio_convertido')

# Cria o diretório de saída, se não existir
if not os.path.exists(output_dir):
    os.makedirs(output_dir)

# Função para verificar se o conteúdo do arquivo é HTML
def is_html(file_path):
    with open(file_path, 'rb') as f:
        header = f.read(512)
        return b'<html>' in header or b'<table>' in header

# Função para processar o arquivo
def process_file(file_path):
    global csv_file_path
    try:
        if is_html(file_path):
            # Process as HTML
            df = pd.read_html(file_path)[0]  # Assumes first table in HTML is the desired one
        elif file_path.endswith(".xlsx"):
            # Process as Excel (.xlsx)
            df = pd.read_excel(file_path, engine='openpyxl')
        elif file_path.endswith(".xls"):
            # Process as Excel (.xls)
            df = pd.read_excel(file_path, engine='xlrd')
        else:
            print(f"Unsupported file format: {file_path}")
            return
        
        # Save as CSV with UTF-8 encoding
        csv_file_name = f"{os.path.splitext(os.path.basename(file_path))[0]}.csv"
        csv_file_path = os.path.join(output_dir, csv_file_name)
        df.to_csv(csv_file_path, index=False, encoding='utf-8-sig')
        print(f"Converted {os.path.basename(file_path)} to {csv_file_name}")

    except Exception as e:
        print(f"Error processing {os.path.basename(file_path)}: {e}")

# Processar o arquivo baixado
process_file(input_file)

# Ler e carregar o arquivo processado para o Google Sheets
data = read_excel(csv_file_path)
print("Incluindo o arquivo na planilha ... ")
upload_to_google_sheets(spreadsheet, 'ABA DE INCLUSÃO AQUI', data)

print("Processo concluído.")
