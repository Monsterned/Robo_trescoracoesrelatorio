import pyautogui
import os
import tkinter as tk
from tkinter import messagebox
import sys
import time
import pandas as pd
from datetime import datetime
import numpy as np
import tkinter as tk
from tkinter import messagebox
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

pasta = os.getcwd() 

# Loop para percorrer todos os arquivos na pasta
for arquivo in os.listdir(pasta):
    # Verifica se o arquivo é uma planilha Excel
    if arquivo.endswith(".xlsx"):
        caminho_arquivo = os.path.join(pasta, arquivo)
        os.remove(caminho_arquivo)
        print(f'{arquivo} excluído com sucesso')

print('Todas as planilhas foram excluídas.')


def show_alert():
    root = tk.Tk()
    root.attributes("-topmost", True)  # Garante que a janela fique em cima das outras
    root.withdraw()  # Esconde a janela principal
    messagebox.showwarning("Aviso", "Site do GreenMile com instabilidade! Por favor tente mais tarde.")
    root.destroy()  # Fecha a janela após o aviso
    sys.exit()  # Encerra o script

def click_selenium(driver, selector, value):
    try:
        print("Clicando no botão...")
        elemento = WebDriverWait(driver, 600).until(EC.element_to_be_clickable((selector, value)))
        elemento.click()
    except Exception as e:
        print(f"Erro ao clicar: {e}")
        show_alert()

def wait_for_xlsx_download(download_dir, timeout=1200):
    """Aguarda até que um novo arquivo .xlsx seja baixado e o download esteja completo."""
    start_time = time.time()
    initial_files = set(os.listdir(download_dir))
    print('Procurando por arquivo .xlsx...')
    
    while time.time() - start_time < timeout:
        current_files = set(os.listdir(download_dir))
        new_files = current_files - initial_files

        if new_files:
            # Filtrar apenas arquivos .xlsx
            xlsx_files = [file for file in new_files if file.endswith('.xlsx')]
            if xlsx_files:
                file_path = os.path.join(download_dir, xlsx_files[0])
                print(f"Novo arquivo .xlsx detectado: {file_path}")
                
                # Verifica se o tamanho do arquivo está mudando
                while time.time() - start_time < timeout:
                    file_size = os.path.getsize(file_path)
                    time.sleep(1)
                    new_size = os.path.getsize(file_path)
                    
                    if file_size == new_size:
                        print("Download concluído.")
                        return file_path  # Retorna o caminho do arquivo baixado
                    time.sleep(1)
        
        time.sleep(1)

    print("Tempo limite de download expirado.")
    return None


# Define o caminho para o diretório de downloads
download_dir = r"C:\Users\Usuario\Documents\CC16\greenmile"

# Nome fixo do arquivo
fixed_filename = "relatorio3coracoes.xlsx"

# Configurações de opções para o Chrome
options = Options()
options.add_argument("--disable-blink-features=AutomationControlled")
options.add_experimental_option("excludeSwitches", ["enable-automation"])
options.add_experimental_option('useAutomationExtension', False)

# Configurações de preferências para downloads
prefs = {
    "download.default_directory": download_dir,  # Define o diretório de download
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True,
    "safebrowsing.disable_download_protection": True,  # Desativa a proteção contra downloads
}
options.add_experimental_option("prefs", prefs)

# Inicializa o driver do Chrome
driver = webdriver.Chrome(options=options)
driver.get("https://3coracoes.greenmile.com/#/login")
driver.maximize_window()



email = 'alessandrajetta'
senha = '3coracoes'

try:
    print("Inserir e-mail...")
    inserir_email = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.ID, 'j_username')))
    inserir_email.click()
    inserir_email.send_keys(email)
    print("E-mail inserido com sucesso...")               
except Exception as e:
    print("Erro ao inserir o e-mail:", e)

try:
    print("Inserir senha do gestor...")
    inserir_senha = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.ID, 'j_password')))
    inserir_senha.click()
    inserir_senha.send_keys(senha)
    print("senha inserida com sucesso...")               
except Exception as e:
    print("Erro ao inserir a senha:", e)

click_selenium(driver, By.XPATH, '//*[@id="LoginBox"]/div[1]/button')

pyautogui.sleep(10)

# LOGIN

driver.get("https://3coracoes.greenmile.com/#/Analytics")

pyautogui.sleep(15)

# Aguarde até o iframe estar presente e mude para o contexto do iframe
iframe = WebDriverWait(driver, 30).until(EC.frame_to_be_available_and_switch_to_it((By.XPATH, '//*[@id="dashboardAnalytics"]/div/iframe')))

pyautogui.sleep(5)

# Agora, dentro do iframe, localize e clique no elemento
try:
    elemento = WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, '//*[@id="root"]/div/div/aside/nav/div/div/div[3]/ul/div[1]/div/li/a/div[2]')))
    elemento.click()
except Exception as e:
    print(f"Erro ao encontrar ou clicar no elemento: {e}")

#pyautogui.sleep(20)
click_selenium(driver, By.XPATH, '//*[@id="root"]/div/div/main/div/div/div[2]/div[2]/div[2]/div[2]/div[2]/a/div/div/div[2]')
#pyautogui.sleep(60)
click_selenium(driver, By.XPATH, '//*[@id="root"]/div/div/main/div/div/div/div/div/div[1]/div/fieldset[2]/div/a/div')
#pyautogui.sleep(5)
print('1')
click_selenium(driver, By.XPATH, '/html/body/span/span/div/div/div/div[1]/button[1]')
print('2')
click_selenium(driver, By.XPATH, '/html/body/span/span/div/div/div/div/button[5]')
#pyautogui.sleep(20)
print('3')
click_selenium(driver, By.XPATH, '//*[@id="Dashboard-Cards-Container"]/div/div/div/div/div/div[1]/div/div[1]/div')
#pyautogui.sleep(20)
print('4')
click_selenium(driver, By.XPATH, '//*[@id="root"]/div/div/main/div/div/div[2]/main/div[3]/div/div[3]/button')
#pyautogui.sleep(10)
print('5')
click_selenium(driver, By.XPATH, '/html/body/div[5]/div/div/div/div/div/button[2]')
# pyautogui.sleep(600)


# Aguarde até que o download seja concluído
downloaded_file_path = wait_for_xlsx_download(download_dir)
if downloaded_file_path:
    print("Arquivo baixado com sucesso!")
    
    # Função para substituir o arquivo antigo pelo novo
    def replace_old_file(download_dir, old_filename, new_file_path):
        old_file_path = os.path.join(download_dir, old_filename)
        if os.path.exists(old_file_path):
            os.remove(old_file_path)
            print(f"Arquivo antigo removido: {old_filename}")
        
        new_file_path_renamed = os.path.join(download_dir, old_filename)
        os.rename(new_file_path, new_file_path_renamed)
        print(f"Novo arquivo renomeado e substituído para: {old_filename}")

    # Substitua o arquivo antigo pelo novo renomeado
    replace_old_file(download_dir, fixed_filename, downloaded_file_path)
else:
    print("Ocorreu um problema com o download.")

# Fecha o navegador
driver.quit()

def click_image(image_path, confidence=0.9):
    current_dir = os.path.dirname(__file__)  # Diretório atual do script
    caminho_imagem = pasta + r'\IMAGENS'
    image_path = os.path.join(current_dir, caminho_imagem, image_path)
    while True:
        try:
            position = pyautogui.locateOnScreen(image_path, confidence=confidence)
            if position:
                center_x = position.left + position.width // 2
                center_y = position.top + position.height // 2
                pyautogui.click(center_x, center_y)
                print("Imagem foi encontrada na tela.")
                break
        except Exception as e:
            print("Imagem não encontrada na tela. Aguardando...")
        pyautogui.sleep(1)

def click_selenium(selector, value):
    try:
        print("Clicando no botão...")
        elemento = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((selector, value)))
        elemento.click()
    except Exception as e:
        print(f"Erro ao clicar: {e}")

# Define o caminho para o diretório de downloads
download_dir = r"C:\Users\Usuario\Documents\CC16\greenmile"

# Configurações de opções para o Chrome
options = Options()

# Desabilita a detecção de automação
options.add_argument("--disable-blink-features=AutomationControlled")
options.add_experimental_option("excludeSwitches", ["enable-automation"])
options.add_experimental_option('useAutomationExtension', False)

# Configurações de preferências para downloads
prefs = {
    "download.default_directory": download_dir,  # Define o diretório de download
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True,
    "safebrowsing.disable_download_protection": True,  # Desativa a proteção contra downloads
}
options.add_experimental_option("prefs", prefs)

# # Inicializa o driver do Chrome  
#driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
driver = webdriver.Chrome(options=options)
driver.get("https://docs.google.com/spreadsheets/d/1dFE52njfx4bink43xdwccsz2l1DkultFVYl4BK3guuI/edit?gid=2076741179#gid=2076741179")
driver.maximize_window()



email = 'andrewl.12.al85@gmail.com'
senha = '99060767'

try:
    print("Inserir e-mail...")
    inserir_email = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.ID, 'identifierId')))
    inserir_email.click()
    inserir_email.send_keys(email)
    print("E-mail inserido com sucesso...")               
except Exception as e:
    print("Erro ao inserir o e-mail:", e)

click_selenium(By.XPATH, '//*[@id="identifierNext"]/div/button/span')

try:
    print("Inserir senha do gestor...")
    inserir_senha = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.NAME, 'Passwd')))
    inserir_senha.click()
    inserir_senha.send_keys(senha)
    print("senha inserida com sucesso...")               
except Exception as e:
    print("Erro ao inserir a senha:", e)

click_selenium(By.XPATH, '//*[@id="passwordNext"]/div/button/span')
click_selenium(By.XPATH, '//*[@id="yDmH0d"]/c-wiz[2]/div/div/div/div/div[2]/div[1]/button/span')
pyautogui.sleep(10)
click_selenium(By.XPATH, '//*[@id=":1f"]/div/div/div[1]/span')







# Espera até que o botão "Arquivo" esteja presente e clica nele usando o full XPath
arquivo_menu = WebDriverWait(driver, 20).until(
    EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/div[4]/div[1]/div[1]/div[1]'))
)
arquivo_menu.click()

# Espera até que a opção "Baixar" esteja visível e clica nela
baixar_opcao = WebDriverWait(driver, 20).until(
    EC.visibility_of_element_located((By.XPATH, '//div[@role="menu"]//*[text()="Download" or text()="Baixar"]'))
)
baixar_opcao.click()

time.sleep(3)

click_image('excel.png')

time.sleep(30)

# Fecha o navegador
driver.quit()

time.sleep(15)

# Carregar e formatar a planilha 'relatorio3coracoes.xlsx'
Planilha_trescoracoes = pd.read_excel("relatorio3coracoes.xlsx")

def formatar_valor(valor):
    # Remove o hífen e tudo que vem depois dele
    valor = valor.split('-')[0]
    # Remove os zeros à esquerda
    valor = valor.lstrip('0')
    return valor

# Aplicando a função na coluna 'Número da Nota Fiscal'
Planilha_trescoracoes['Número da Nota Fiscal'] = Planilha_trescoracoes['Número da Nota Fiscal'].apply(formatar_valor)
Planilha_trescoracoes['Número da Nota Fiscal'] = Planilha_trescoracoes['Número da Nota Fiscal'].astype(np.int64)

# Ler a planilha 'JETTA X 3CORAÇÕES.xlsx'
Planilha_relatorio = pd.read_excel('JETTA X 3CORAÇÕES.xlsx', sheet_name='ACOMP. DE ENTREGAS', skiprows=16)

# Iterar sobre as linhas da planilha
for i, linha in Planilha_relatorio.iterrows():
    nf = linha["NF"]
    retificacao = linha["RETIFICAÇÃO TRANSP."]

    # Verificar se retificacao é NaN
    if pd.isna(retificacao):
        #print(f'Nota: {nf} - Retificação: Não disponível (NaN)')
        
        if nf in Planilha_trescoracoes['Número da Nota Fiscal'].values:
            index = Planilha_trescoracoes[Planilha_trescoracoes['Número da Nota Fiscal'] == nf].index.values[0]
            status = Planilha_trescoracoes.at[index, 'Assinado Eletronicamente']
            
            if status == 'Assinado':
                descricao = Planilha_trescoracoes.at[index, 'Chegada no Cliente']
                descricao = str(descricao)[:10]
                
                # Converte a string para datetime e para o formato desejado
                descricao = datetime.strptime(descricao, "%Y-%m-%d").strftime("%d/%m/%Y")
                descricao = f'NF DE ENTREGUE DIA {descricao}'
                #print(descricao)
                # Atualiza a planilha com a descrição formatada
                Planilha_relatorio.at[i, "RETIFICAÇÃO TRANSP."] = descricao
            else:
                print('Não assinado')

# Salvar a planilha editada como 'BASE_DADOS.xlsx'
Planilha_relatorio.to_excel('BASE_DADOS.xlsx', index=False)

# Mostrar caixa de aviso usando tkinter
root = tk.Tk()
root.withdraw()  # Oculta a janela principal
messagebox.showinfo("Aviso", "Processo finalizado com sucesso!")