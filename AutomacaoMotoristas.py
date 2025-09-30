from selenium import webdriver
from selenium.webdriver.common.by import By 
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import os
import shutil
import csv

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
download_folder = r"C:\Users\flavi\Downloads"
pasta_destino = r"C:\Users\flavi\OneDrive\001\Bases BI"

# ---------- FUNÇÕES AUXILIARES ----------
def esperar_novo_arquivo(pasta, prefixo="CSVmotoristas", extensao=".sswweb", timeout=180):
    """Espera até que um NOVO arquivo com prefixo e extensão apareça na pasta e finalize o download."""
    arquivos_antes = set(os.listdir(pasta))
    
    for _ in range(timeout):
        arquivos_depois = set(os.listdir(pasta))
        novos = arquivos_depois - arquivos_antes

        # filtra pelo prefixo e extensão correta
        novos_validos = [
            f for f in novos 
            if f.startswith(prefixo) and f.endswith(extensao) and not f.endswith(".crdownload")
        ]

        if novos_validos:
            caminho = os.path.join(pasta, novos_validos[0])
            # espera sumir o .crdownload
            while os.path.exists(caminho + ".crdownload"):
                time.sleep(1)
            return caminho

        time.sleep(1)
    return None

def converter_sswweb_para_csv(arquivo_origem, arquivo_destino):
    """Converte o arquivo .sswweb para .csv com encoding e separador corrigidos."""
    with open(arquivo_origem, "r", encoding="latin-1") as infile:
        linhas = infile.readlines()

    with open(arquivo_destino, "w", newline="", encoding="utf-8") as outfile:
        writer = csv.writer(outfile, delimiter=";")
        for linha in linhas:
            writer.writerow(linha.strip().split(";"))

# ---------- SCRIPT PRINCIPAL ----------
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

    # --- AQUI entra a ação que dispara o download ---
    # Exemplo:
    # driver.find_element(By.ID, 'botao_download').click()

    print("⏳ Aguardando download do arquivo CSVmotoristas...")
    ultimo_arquivo = esperar_novo_arquivo(download_folder, prefixo="CSVmotoristas", extensao=".sswweb")

    if ultimo_arquivo:
        print(f"📂 Arquivo baixado detectado: {ultimo_arquivo}")

        # Define nome final CSV
        arquivo_destino = os.path.join(pasta_destino, "motoristastki100132.csv")

        # Converte e salva direto na pasta de destino
        if not os.path.exists(pasta_destino):
            os.makedirs(pasta_destino)

        converter_sswweb_para_csv(ultimo_arquivo, arquivo_destino)
        print(f"✅ Arquivo convertido e movido para: {arquivo_destino}")

        # Remove o arquivo original .sswweb
        os.remove(ultimo_arquivo)
        print("🗑️ Arquivo .sswweb original removido.")

    else:
        print("⚠️ Nenhum arquivo CSVmotoristas encontrado dentro do tempo limite.")

finally:
    input('Pressione Enter para finalizar e fechar o navegador...')
    driver.quit()
