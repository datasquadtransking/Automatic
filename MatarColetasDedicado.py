import pandas as pd
import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options

# === 1️⃣ Ler a planilha e pegar a coluna COLETA ===
caminho = r"C:\Users\flavi\Desktop\Baixar dedicado.xlsx"
df = pd.read_excel(caminho)
df = df.dropna(subset=['COLETA'])
coletas = [int(float(x)) for x in df['COLETA']]
print("✅ Coletas carregadas:", coletas)

# === 2️⃣ Dados de login e IDs ===
LOGIN_URL = "http://sistema.ssw.inf.br"
DOM = 'TKI'
CPF = '11744601640'
USUARIO = 'weslley'
SENHA = 'W@tki111'

# ⚠️ IDs dos campos
ID_DOM = '1'
ID_CPF = '2'
ID_USUARIO = '3'
ID_SENHA = '4'
ID_BOTAO_LOGIN = '5'


# === Função para login e retorno do driver ===
def login_ssw():
    chrome_options = Options()
    chrome_options.add_argument("--window-size=1920,1080")
    chrome_options.add_argument("--log-level=3")

    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)

    # <<< AQUI ele minimiza imediatamente >>>
    driver.minimize_window()

    driver.get(LOGIN_URL)

    WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.ID, ID_DOM)))
    driver.find_element(By.ID, ID_DOM).send_keys(DOM)
    driver.find_element(By.ID, ID_CPF).send_keys(CPF)
    driver.find_element(By.ID, ID_USUARIO).send_keys(USUARIO)
    driver.find_element(By.ID, ID_SENHA).send_keys(SENHA)
    driver.find_element(By.ID, ID_BOTAO_LOGIN).click()

    print("✅ Login efetuado")
    time.sleep(3)
    return driver



# === Loop principal para cada coleta ===
for coleta in coletas:
    try:
        print(f"\n🚀 Iniciando processamento da coleta {coleta}")

        driver = login_ssw()

        # === Campo f2 ===
        campo_f2 = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.ID, '2'))
        )
        campo_f2.click()
        campo_f2.clear()
        campo_f2.send_keys("SLI")
        driver.execute_script(
            "arguments[0].dispatchEvent(new Event('change'));", campo_f2
        )
        print("✅ Campo 'f2' atualizado para SLI")

        # === Digitar 003 e abrir nova guia ===
        campo_103 = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, '3'))
        )
        campo_103.send_keys("003")
        time.sleep(3)

        handles = driver.window_handles
        if len(handles) > 1:
            driver.switch_to.window(handles[-1])
        print("✅ Nova guia aberta e foco alterado")

        # === Processar coleta ===
        campo_f7 = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, '7'))
        )
        campo_f7.click()
        campo_f7.clear()
        campo_f7.send_keys(str(coleta))
        time.sleep(3)

        btn_pesquisar = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, '8'))
        )
        btn_pesquisar.click()
        time.sleep(3)

        driver.switch_to.window(driver.window_handles[-1])

        campo_ocor = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, 'fld_cod_ocor'))
        )
        campo_ocor.click()
        campo_ocor.clear()
        campo_ocor.send_keys("11")
        time.sleep(1)

        btn_confirmar = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, 'env_oco'))
        )
        btn_confirmar.click()
        time.sleep(3)

        print(f"✅ Coleta {coleta} processada com sucesso")

    except Exception as e:
        print(f"❌ Erro na coleta {coleta}: {e}")

    finally:
        driver.quit()
        print(f"🔒 Navegador fechado após coleta {coleta}")
        time.sleep(2)
