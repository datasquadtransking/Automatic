# -*- coding: utf-8 -*-
import pandas as pd
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import os

# ===============================================================
# Caminho do SEU arquivo
# ===============================================================
CAMINHO_PLANILHA = r"C:\Users\flavi\OneDrive\Weslley\AcompanhamentoUnilever.xlsx"

# ===============================================================
# 1) Ler planilha
# ===============================================================
try:
    df = pd.read_excel(CAMINHO_PLANILHA, sheet_name="Acompanhamento", skiprows=1)
except FileNotFoundError:
    print(f"ERRO: Arquivo não encontrado no caminho: {CAMINHO_PLANILHA}")
    exit()
except Exception as e:
    print(f"ERRO ao ler o arquivo Excel: {e}")
    exit()

# Ajustar coluna ID para só o que vem depois do último "/"
df["ID"] = df["ID"].astype(str).apply(lambda x: x.split("/")[-1])

# Criar coluna de resultado para entregas no DataFrame principal
df.loc[:, "CHECK_ENTREGA"] = None

# ===============================================================
# 2) Filtrar somente ENTREGAS com SLA pendente
# ===============================================================
df_entrega = df[df["TIPO"] == "ENTREGA"].copy()
df_entrega.loc[:, "SLA"] = pd.to_numeric(df_entrega["SLA"], errors='coerce')

filtro_sla_entrega = (df_entrega["SLA"].isna()) | (
    (df_entrega["SLA"] != 1) & (df_entrega["SLA"] != 0)
)
df_entrega = df_entrega[filtro_sla_entrega].copy()

print(f"Total de ENTREGAS a processar: {len(df_entrega)}")

# ===============================================================
# 3) Iniciar navegador e Login
# ===============================================================
options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
driver = None

try:
    print("\nIniciando navegador e fazendo login...")
    driver = webdriver.Chrome(options=options)

    driver.get("https://unilever2.brasilrisk.com.br/Login/Logout")

    WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.ID, "usuario"))
    )

    usuario = driver.find_element(By.ID, "usuario")
    usuario.send_keys("MONITORAMENTO1")

    senha = driver.find_element(By.ID, "senha")
    senha.send_keys("Transking@2026")

    login_btn = driver.find_element(By.ID, "Login")
    login_btn.click()

    time.sleep(2)

    botao_search = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, "botaosearch"))
    )
    botao_search.click()

    campo_search = driver.find_element(By.ID, "input-search")

    # ===============================================================
    # 4) Loop para processar cada entrega
    # ===============================================================
    total_entregas = len(df_entrega)
    processadas = 0
    consultas_entrega = []

    for index, row in df_entrega.iterrows():
        id_viagem = str(row["ID"])
        idx_original = row.name

        processadas += 1
        faltam = total_entregas - processadas
        print(f"\n-> Processando ENTREGA ID: {id_viagem}")
        print(f"Processadas: {processadas} / {total_entregas}  |  Faltam: {faltam}")

        try:
            campo_search.clear()
        except Exception:
            campo_search = driver.find_element(By.ID, "input-search")
            campo_search.clear()

        campo_search.send_keys(id_viagem)
        time.sleep(1)
        campo_search.send_keys(Keys.RETURN)
        time.sleep(2)

        try:
            aba_entregas = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, "tabEntrega"))
            )
            aba_entregas.click()
            time.sleep(1)
        except (NoSuchElementException, TimeoutException):
            print(f"ERRO: Não foi possível abrir a aba 'Entregas' para o ID {id_viagem}.")
            df.loc[idx_original, "CHECK_ENTREGA"] = "SEM_ABA_ENTREGAS"
            consultas_entrega.append({"ID": id_viagem, "CHECK_ENTREGA": "SEM_ABA_ENTREGAS"})
            continue

        # ===============================================================
        # 5) CAPTURAR A ÚLTIMA LINHA da tabela (COLUNA 8 e 10)
        # ===============================================================
        valor_chegada = None
        valor_dme = None

        try:
            xpath_valor_chegada = "//div[@id='entregasnf']//tbody/tr[last()]/td[10]"
            xpath_valor_dme = "//div[@id='entregasnf']//tbody/tr[last()]/td[8]"

            elemento_valor_chegada = WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.XPATH, xpath_valor_chegada))
            )
            valor_chegada = elemento_valor_chegada.text.strip()

            elemento_valor_dme = WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.XPATH, xpath_valor_dme))
            )
            valor_dme = elemento_valor_dme.text.strip()

            if not valor_chegada:
                raise Exception("Valor de chegada vazio.")

            print(f"Valor 'DME' da ÚLTIMA linha encontrado: {valor_dme}")
            print(f"Valor 'Chegada no cliente' da ÚLTIMA linha encontrado: {valor_chegada}")

        except Exception:
            print("Falha na captura por XPath direto da última linha. Tentando fallback.")
            try:
                linhas = WebDriverWait(driver, 10).until(
                    EC.presence_of_all_elements_located((By.XPATH, "//div[@id='entregasnf']//tbody/tr"))
                )

                if linhas:
                    ultima_linha = linhas[-1]
                    colunas = ultima_linha.find_elements(By.TAG_NAME, "td")

                    if len(colunas) >= 10:
                        valor_dme = colunas[7].text.strip()
                        valor_chegada = colunas[9].text.strip()
                        print(f"Valor 'DME' via Fallback: {valor_dme}")
                        print(f"Valor 'Chegada no cliente' via Fallback: {valor_chegada}")
                    else:
                        valor_chegada = "COLUNA_10_NAO_EXISTE"
                else:
                    valor_chegada = "TABELA_VAZIA"

            except Exception:
                df.loc[idx_original, "CHECK_ENTREGA"] = "ERRO_COLETA"
                consultas_entrega.append({"ID": id_viagem, "CHECK_ENTREGA": "ERRO_COLETA"})
                continue

        # ===============================================================
        # Registro do resultado
        # ===============================================================
        df.loc[idx_original, "CHECK_ENTREGA"] = f"DME: {valor_dme} | CHEGADA: {valor_chegada}"
        consultas_entrega.append({
            "ID": id_viagem,
            "DME": valor_dme,
            "CHEGADA_CLIENTE": valor_chegada
        })

        try:
            driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.ESCAPE)
            time.sleep(1)
        except Exception:
            pass

finally:
    if driver:
        driver.quit()

print("\n--- Processo ENTREGAS concluído ---")

# ===============================================================
# 6) Salvar o log de consultas
# ===============================================================
df_consultas = pd.DataFrame(consultas_entrega)

try:
    with pd.ExcelWriter(
        CAMINHO_PLANILHA,
        engine='openpyxl',
        mode='a',
        if_sheet_exists='replace'
    ) as writer:
        df_consultas.to_excel(writer, sheet_name="EntregasConsultadas", index=False)
    print("\nAba 'EntregasConsultadas' salva com sucesso!")
except Exception as e:
    print(f"\nERRO ao salvar a aba 'EntregasConsultadas': {e}")

print("\nLog de Consultas:")
print(df_consultas.head())
