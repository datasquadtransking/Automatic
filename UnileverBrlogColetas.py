import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

# --- CONFIGURAÇÃO INICIAL E LEITURA DO ARQUIVO ---

caminho = r"C:\Users\flavi\OneDrive\Weslley\AcompanhamentoUnilever.xlsx"

df = pd.read_excel(caminho, sheet_name="Acompanhamento", skiprows=1)

df["ID"] = df["ID"].astype(str).apply(lambda x: x.split("/")[-1])

df["CHECKIN"] = None
df["DATA_CAMPO"] = None

# --- FILTRAGEM ---

df_coleta = df[df["TIPO"] == "COLETA"].copy()
df_coleta["SLA"] = pd.to_numeric(df_coleta["SLA"], errors="coerce")

filtro_sla = (df_coleta["SLA"].isna()) | ((df_coleta["SLA"] != 1) & (df_coleta["SLA"] != 0))
df_coleta = df_coleta[filtro_sla].copy()

print(f"Total de linhas a processar: {len(df_coleta)}")

# --- CONTROLE ---

total = len(df_coleta)
processados = 0
ColetasConsultadas = []

# --- SELENIUM ---

options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")

driver = None

try:
    driver = webdriver.Chrome(options=options)
    driver.get("https://unilever2.brasilrisk.com.br/Login/Logout")

    driver.find_element(By.ID, "usuario").send_keys("MONITORAMENTO1")
    driver.find_element(By.ID, "senha").send_keys("Transking@2026")
    driver.find_element(By.ID, "Login").click()

    time.sleep(2)

    driver.find_element(By.ID, "botaosearch").click()
    time.sleep(1)

    campo_search = driver.find_element(By.ID, "input-search")

    # --- LOOP PRINCIPAL ---

    for index, row in df_coleta.iterrows():
        valor_id = str(row["ID"])
        processados += 1

        print(f"\n-> ID {valor_id} | {processados}/{total}")

        campo_search.clear()
        campo_search.send_keys(valor_id)
        campo_search.send_keys(Keys.RETURN)

        time.sleep(1)

        modal_aberto = False
        valor_checkin = None
        valor_data_campo = None

        try:
            wait = WebDriverWait(driver, 10)
            chevron = wait.until(EC.presence_of_element_located((By.ID, "chevViagem")))
            chevron.click()
            modal_aberto = True
            time.sleep(1)

            # --- CHECKIN ---
            try:
                checkin_div = driver.find_element(
                    By.XPATH,
                    "//div[@class='boxFloat'][.//span[text()='CheckIn CD']]"
                )
                checkin_element = checkin_div.find_element(By.CLASS_NAME, "data-value")
                valor_checkin = checkin_element.text.strip() or "Hora vazia"

            except NoSuchElementException:
                valor_checkin = "CheckIn não encontrado"

            df.loc[df["ID"] == valor_id, "CHECKIN"] = valor_checkin

            # --- DATA CAMPO ---
            try:
                data_div = driver.find_element(
                    By.XPATH,
                    "//div[@class='boxFloat'][.//span[text()='Dock']]"
                )
                data_element = data_div.find_element(By.CLASS_NAME, "data-value")
                valor_data_campo = data_element.text.strip() or "Data vazia"

            except NoSuchElementException:
                valor_data_campo = "Data não encontrada"

            df.loc[df["ID"] == valor_id, "DATA_CAMPO"] = valor_data_campo

            print(f"CheckIn: {valor_checkin} | Data Campo: {valor_data_campo}")

        except (NoSuchElementException, TimeoutException):
            valor_checkin = "SEM_VIAGEM"
            valor_data_campo = "SEM_VIAGEM"

            df.loc[df["ID"] == valor_id, "CHECKIN"] = valor_checkin
            df.loc[df["ID"] == valor_id, "DATA_CAMPO"] = valor_data_campo

        # --- REGISTRO PARA NOVA PLANILHA ---
        ColetasConsultadas.append({
            "ID": valor_id,
            "CHECKIN": valor_checkin,
            "DATA_CAMPO": valor_data_campo
        })

        # --- FECHA MODAL ---
        if modal_aberto:
            driver.find_element(By.TAG_NAME, "body").send_keys(Keys.ESCAPE)
            time.sleep(1)

finally:
    if driver:
        driver.quit()

# --- SALVA NOVA PLANILHA ---

df_novo = pd.DataFrame(ColetasConsultadas)

with pd.ExcelWriter(caminho, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
    df_novo.to_excel(writer, sheet_name="ColetasConsultadas", index=False)

print("\nProcesso finalizado com sucesso!")
