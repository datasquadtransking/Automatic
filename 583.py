from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import os
import shutil

# ======================
# CONFIGURAÇÃO DE LOGIN
# ======================
LOGIN_URL = "http://sistema.ssw.inf.br"
DOM = "tki"
CPF = "11744601640"
USUARIO = "weslley"
SENHA = "W@tki111"

# XPATHS DE LOGIN
XPATH_DOM = '/html/body/form/input[1]'
XPATH_CPF = '/html/body/form/input[2]'
XPATH_USUARIO = '/html/body/form/input[3]'
XPATH_SENHA = '/html/body/form/input[4]'
XPATH_BOTAO_LOGIN = '/html/body/form/a'

# ======================
# CONFIGURAÇÃO DE DOWNLOAD
# ======================
download_folder = r"C:\Users\flavi\Downloads"
destino_arquivo = r"C:\Users\flavi\OneDrive\001\Bases BI\Clientes.csv"

def pegar_arquivo_completo(folder, parte_nome, extensao="csv", timeout=180):
    """
    Espera até encontrar um arquivo no folder que:
    - contenha parte_nome no nome
    - tenha a extensão especificada
    - não esteja com extensão temporária (.crdownload)
    """
    tempo_inicial = time.time()
    while True:
        arquivos = [f for f in os.listdir(folder) if parte_nome in f and f.lower().endswith(extensao.lower())]
        arquivos_completos = [f for f in arquivos if not f.endswith(".crdownload")]
        if arquivos_completos:
            return os.path.join(folder, arquivos_completos[0])
        if time.time() - tempo_inicial > timeout:
            raise TimeoutError(f"⏰ Arquivo com '{parte_nome}' não apareceu em {timeout} segundos.")
        time.sleep(1)

# ======================
# SCRIPT PRINCIPAL
# ======================
try:
    # --- INICIAR NAVEGADOR ---
    driver = webdriver.Chrome()
    driver.maximize_window()
    wait = WebDriverWait(driver, 15)
    driver.get(LOGIN_URL)

    # --- LOGIN ---
    wait.until(EC.presence_of_element_located((By.XPATH, XPATH_DOM)))
    driver.find_element(By.XPATH, XPATH_DOM).send_keys(DOM)
    driver.find_element(By.XPATH, XPATH_CPF).send_keys(CPF)
    driver.find_element(By.XPATH, XPATH_USUARIO).send_keys(USUARIO)
    driver.find_element(By.XPATH, XPATH_SENHA).send_keys(SENHA)
    driver.find_element(By.XPATH, XPATH_BOTAO_LOGIN).click()
    print("✅ Login realizado.")

    # --- ESPERA PELA TELA PRINCIPAL ---
    WebDriverWait(driver, 10).until(EC.staleness_of(driver.find_element(By.XPATH, XPATH_CPF)))

    # --- DIGITAR 583 PARA ABRIR A TELA ---
    campo_codigo = wait.until(EC.presence_of_element_located((By.ID, "3")))
    campo_codigo.send_keys("583")
    print("✅ Código 583 digitado.")

    # --- MUDAR PARA A NOVA ABA ---
    time.sleep(2)  # espera a nova aba abrir
    driver.switch_to.window(driver.window_handles[-1])
    print("✅ Foco na aba da tela 583.")

    # --- CLICAR NO BOTÃO EXCEL ---
    botao_excel = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(
            (By.XPATH, "//a[@onclick=\"ajaxEnvia('IMP_EXC', 0);return false;\"]")
        )
    )
    driver.execute_script("arguments[0].click();", botao_excel)
    print("✅ Botão Excel clicado com sucesso.")

    # --- AGUARDAR DOWNLOAD COMPLETO ---
    print("⏳ Aguardando download do CSV...")
    caminho_csv = pegar_arquivo_completo(download_folder, "ssw0662", "csv", timeout=180)
    print(f"✅ Download completo: {caminho_csv}")

    # --- MOVER E RENOMEAR ARQUIVO ---
    os.makedirs(os.path.dirname(destino_arquivo), exist_ok=True)
    if os.path.exists(destino_arquivo):
        os.remove(destino_arquivo)
    shutil.move(caminho_csv, destino_arquivo)
    print(f"✅ Arquivo movido e renomeado para: {destino_arquivo}")

    # --- MANTER ABERTO PARA INSPEÇÃO ---
    input("Pressione ENTER para fechar o navegador...")
    driver.quit()

except Exception as e:
    print("❌ Erro:", e)
    try:
        driver.quit()
    except:
        pass
