from selenium import webdriver
from selenium.webdriver.common.by import By 
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import os
import shutil

# CONFIGURAÇÃO DE LOGIN
LOGIN_URL = "http://sistema.ssw.inf.br"
DOM = 'tki'
CPF = '11744601640'
USUARIO = 'weslley'
SENHA = 'W@tki111'

# XPATHS DE LOGIN
XPATH_DOM = '/html/body/form/input[1]'
XPATH_CPF = '/html/body/form/input[2]'
XPATH_USUARIO = '/html/body/form/input[3]'
XPATH_SENHA = '/html/body/form/input[4]'
XPATH_BOTAO_LOGIN = '/html/body/form/a'

# PASTAS
download_folder = r"C:\Users\weslley\Downloads"
pasta_destino = r"C:\Users\Weslley\OneDrive\001\Bases BI"

def pegar_ultimo_arquivo(folder, extensao=None):
    arquivos = [os.path.join(folder, f) for f in os.listdir(folder)]
    if extensao:
        arquivos = [f for f in arquivos if f.lower().endswith(extensao.lower())]
    if not arquivos:
        return None
    arquivo_recente = max(arquivos, key=os.path.getmtime)
    return arquivo_recente

try:
    # INICIAR NAVEGADOR
    driver = webdriver.Chrome()
    driver.maximize_window()
    driver.get(LOGIN_URL)

    # ✅ Realiza o login
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, XPATH_DOM)))
    driver.find_element(By.XPATH, XPATH_DOM).send_keys(DOM)
    driver.find_element(By.XPATH, XPATH_CPF).send_keys(CPF)
    driver.find_element(By.XPATH, XPATH_USUARIO).send_keys(USUARIO)
    driver.find_element(By.XPATH, XPATH_SENHA).send_keys(SENHA)
    driver.find_element(By.XPATH, XPATH_BOTAO_LOGIN).click()
    print("✅ Login realizado com sucesso.")

    # ✅ Digita 047 para abrir nova guia
    WebDriverWait(driver, 10).until(EC.staleness_of(driver.find_element(By.XPATH, XPATH_CPF)))
    campo_codigo = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, '3')))
    campo_codigo.send_keys("047")
    print("✅ Digitado 047 para abrir nova guia.")
    time.sleep(3)

    # ✅ Muda para a nova aba
    driver.switch_to.window(driver.window_handles[-1])
    print("✅ Página 047 aberta com sucesso.")

    # ✅ Localiza o campo de id='13', apaga e digita 's'
    campo_13 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, '13')))
    campo_13.clear()
    campo_13.send_keys("s")
    print("✅ Campo id='13' atualizado para 's'.")

    # --- AQUI entra sua lógica para realizar o DOWNLOAD do relatório ---
    # Exemplo:
    # driver.find_element(By.ID, 'botao_download').click()

    print("⏳ Aguardando o download terminar...")
    time.sleep(90)  # ajuste conforme o tempo do seu download

    # Renomeia último arquivo baixado para 'motoristas047.csv'
    ultimo_arquivo = pegar_ultimo_arquivo(download_folder, '.sswweb')  # ajuste a extensão correta
    if ultimo_arquivo:
        arquivo_renomeado = os.path.join(download_folder, 'motoristastki100132.csv')
        if os.path.exists(arquivo_renomeado):
            os.remove(arquivo_renomeado)
        shutil.move(ultimo_arquivo, arquivo_renomeado)
        print(f"✅ Arquivo renomeado para: {arquivo_renomeado}")

        # Move para a pasta de destino
        if not os.path.exists(pasta_destino):
            os.makedirs(pasta_destino)  # cria a pasta de destino se não existir
        arquivo_destino = os.path.join(pasta_destino, 'motoristastki100132.csv')
        if os.path.exists(arquivo_destino):
            os.remove(arquivo_destino)  # remove se existir
        shutil.move(arquivo_renomeado, arquivo_destino)
        print(f"✅ Arquivo movido para: {arquivo_destino}")

    else:
        print("⚠️ Nenhum arquivo encontrado para renomear e mover.")

finally:
    input('Pressione Enter para finalizar e fechar o navegador...')
    driver.quit()
