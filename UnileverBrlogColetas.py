import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException 
import time
import numpy as np 
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException 

# --- CONFIGURAÇÃO INICIAL E LEITURA DO ARQUIVO ---

# Caminho do arquivo
caminho = r"C:\Users\flavi\OneDrive\Weslley\AcompanhamentoUnilever.xlsx"

# Lendo a aba "Acompanhamento" e ignorando a primeira linha
df = pd.read_excel(caminho, sheet_name="Acompanhamento", skiprows=1)

# Ajustando a coluna ID para conter apenas o que vem depois do último "/"
df["ID"] = df["ID"].astype(str).apply(lambda x: x.split("/")[-1])

# Adiciona a nova coluna 'CHECKIN' ao DataFrame
df["CHECKIN"] = None

# --- FILTRAGEM DE DADOS ---

df_coleta = df[df["TIPO"] == "COLETA"].copy()

df_coleta.loc[:, "SLA"] = pd.to_numeric(df_coleta["SLA"], errors='coerce') 

filtro_sla = (df_coleta["SLA"].isna()) | ((df_coleta["SLA"] != 1) & (df_coleta["SLA"] != 0))

df_coleta = df_coleta[filtro_sla].copy()

print(f"Total de linhas a processar (TIPO = COLETA E SLA VAZIO/DIFERENTE DE 1 e 0): {len(df_coleta)}")


# --- ADIÇÃO: CONTADOR E LISTA PARA PLANILHA NOVA ---

total = len(df_coleta)
processados = 0
ColetasConsultadas = []


# --- INICIALIZAÇÃO DO NAVEGADOR E LOGIN ---

options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
driver = None
try:
    driver = webdriver.Chrome(options=options)

    driver.get("https://unilever2.brasilrisk.com.br/Login/Logout")

    usuario = driver.find_element(By.ID, "usuario")
    usuario.send_keys("MONITORAMENTO1")

    senha = driver.find_element(By.ID, "senha")
    senha.send_keys("Transking@2026")

    login_btn = driver.find_element(By.ID, "Login")
    login_btn.click()

    time.sleep(1)

    botao_search = driver.find_element(By.ID, "botaosearch")
    botao_search.click()

    time.sleep(1)

    campo_search = driver.find_element(By.ID, "input-search")

    janela_principal = driver.current_window_handle 


    # --- LOOP PARA PROCESSAR CADA ID ---

    for index, row in df_coleta.iterrows():
        valor_id = str(row["ID"])
        print(f"\n-> Processando ID: {valor_id}")

        # --- ADIÇÃO: CONTADOR ---
        processados += 1
        faltam = total - processados
        print(f"Consultados: {processados} / {total}  |  Faltam: {faltam}")

        campo_search.clear()
        campo_search.send_keys(valor_id)

        time.sleep(1)

        campo_search.send_keys(Keys.RETURN)

        time.sleep(1) 
        
        modal_aberto = False 

        try:
            wait = WebDriverWait(driver, 10)
            chevron = wait.until(EC.presence_of_element_located((By.ID, "chevViagem")))
            chevron.click()

            print(f"Modal de detalhe da viagem aberto para ID {valor_id}")
            modal_aberto = True
            
            time.sleep(1) 
            
            try:
                checkin_parent_div = driver.find_element(
                    By.XPATH, 
                    "//div[@class='boxFloat'][.//span[text()='CheckIn CD']]"
                )

                checkin_element = checkin_parent_div.find_element(
                    By.CLASS_NAME, 
                    "data-value"
                )
                
                valor_checkin = checkin_element.text
                
                if not valor_checkin.strip():
                    valor_checkin = "Hora de chegada Vazia na Tela"
                    
                print(f"Hora de chegada encontrada: {valor_checkin}")

                df.loc[df["ID"] == valor_id, "CHECKIN"] = valor_checkin
                
            except NoSuchElementException as checkin_e:
                valor_checkin = "Hora de chegada não encontrada"
                print(f"Erro ao encontrar Hora de chegada para ID {valor_id}. {checkin_e}")
                df.loc[df["ID"] == valor_id, "CHECKIN"] = valor_checkin

            except Exception as checkin_e:
                valor_checkin = "Erro de coleta de Hora de chegada"
                print(f"Erro inesperado ao coletar Hora de chegada para ID {valor_id}: {checkin_e}")
                df.loc[df["ID"] == valor_id, "CHECKIN"] = valor_checkin
                
        except (NoSuchElementException, TimeoutException) as chevron_e:
            print(f"ERRO: Ícone 'chevViagem' não encontrado para ID {valor_id}. {chevron_e}")
            valor_checkin = "SEM_VIAGEM"
            df.loc[df["ID"] == valor_id, "CHECKIN"] = valor_checkin

        except Exception as e:
            print(f"ERRO geral para ID {valor_id}. {e}")
            valor_checkin = "ERRO_GERAL"
            df.loc[df["ID"] == valor_id, "CHECKIN"] = valor_checkin


        # --- ADIÇÃO: REGISTRA CONSULTA PARA A PLANILHA NOVA ---
        ColetasConsultadas.append({
            "ID": valor_id,
            "CHECKIN": valor_checkin
        })


        # --- FECHA MODAL ---
        if modal_aberto:
            try:
                driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.ESCAPE)
                time.sleep(1)
                print("Modal fechado via ESC.")
            except Exception as close_e:
                print(f"Aviso: Falha ao fechar o modal via ESC. {close_e}")

finally:
    if driver:
        driver.quit()

print("\n--- Processo concluído ---")

print("\nDataFrame com a nova coluna CHECKIN:")
print(df.head())


# --- ADIÇÃO: SALVAR SOMENTE AS CONSULTAS REALIZADAS ---

df_novo = pd.DataFrame(ColetasConsultadas)

try:
    with pd.ExcelWriter(caminho, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        df_novo.to_excel(writer, sheet_name="ColetasConsultadas", index=False)
    print("\nPlanilha 'ColetasConsultadas' salva com sucesso!")
except Exception as e:
    print(f"\nERRO ao salvar planilha de consultas: {e}")