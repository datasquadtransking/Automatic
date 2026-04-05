import pandas as pd
import numpy as np
import os
import time
from datetime import datetime, timedelta
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# --- CONFIGURAÇÕES DE CAMINHO ---
PASTA_PROJETO = os.path.dirname(os.path.abspath(__file__))
ARQUIVO_ENTRADA = os.path.join(PASTA_PROJETO, "RelatorioAtualizado.xlsx")
ARQUIVO_SAIDA = os.path.join(PASTA_PROJETO, "Base_Chegada_Unilever.xlsx")

def preparar_dados():
    try:
        df = pd.read_excel(ARQUIVO_ENTRADA, engine='openpyxl')
        df.columns = df.columns.str.strip().str.upper()

        df_unilever = df[df['CLIENTE'].str.contains('UNILEVER', na=False, case=False)].copy()
        if df_unilever.empty:
            return None

        # Tratamento de Datas
        df_unilever['DT_COLETA_OBJ'] = pd.to_datetime(df_unilever['PROGRAMAÇÃO. COLETA'], dayfirst=True)
        df_unilever['DT_ENTREGA_OBJ'] = pd.to_datetime(df_unilever['PROGRAMAÇÃO. ENTREGA'], dayfirst=True)
        
        # Ajuste para data alvo (Feriado/Próximos dias)
        data_alvo = datetime.now().date() + timedelta(days=0) 

        condicoes = [
            (df_unilever['DT_COLETA_OBJ'].dt.date == data_alvo),
            (df_unilever['DT_ENTREGA_OBJ'].dt.date == data_alvo)
        ]
        df_unilever['TIPO'] = np.select(condicoes, ['COLETA', 'ENTREGA'], default='OUTROS')

        base_final = df_unilever[df_unilever['TIPO'] != 'OUTROS'].copy()
        base_final["ID_CONSULTA"] = base_final["ID"].astype(str).apply(lambda x: x.split("/")[-1])
        
        # Inicializa colunas conforme os dois códigos originais
        base_final["HORA_CHEGADA"] = "---"
        base_final["CHECKIN_CD"] = "---"
        base_final["DOCK_DATA"] = "---"
        base_final["DME"] = "---"
        base_final["CHEGADA_CLIENTE"] = "---"

        return base_final
    except Exception as e:
        print(f"❌ Erro na leitura: {e}")
        return None

def realizar_consultas_portal(df):
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    driver = webdriver.Chrome(options=options)
    wait = WebDriverWait(driver, 20)

    try:
        driver.get("https://unilever2.brasilrisk.com.br/Login/Logout")
        wait.until(EC.presence_of_element_located((By.ID, "usuario"))).send_keys("MONITORAMENTO1")
        driver.find_element(By.ID, "senha").send_keys("@Tk@2026")
        driver.find_element(By.ID, "Login").click()
        
        wait.until(EC.element_to_be_clickable((By.ID, "botaosearch"))).click()
        time.sleep(1)

        for index, row in df.iterrows():
            id_viagem = row["ID_CONSULTA"]
            tipo = row["TIPO"]

            try:
                campo_search = driver.find_element(By.ID, "input-search")
                campo_search.clear()
                campo_search.send_keys(id_viagem)
                campo_search.send_keys(Keys.RETURN)
                
                if tipo == "COLETA":
                    chevron = wait.until(EC.element_to_be_clickable((By.ID, "chevViagem")))
                    chevron.click()
                    time.sleep(2)
                    
                    # Captura conforme Parte2_UnileverCol.py
                    try:
                        df.at[index, "HORA_CHEGADA"] = driver.find_element(By.XPATH, "//div[@class='boxFloat'][.//span[contains(text(),'Hora de chegada')]]").find_element(By.CLASS_NAME, "data-value").text.strip()
                        df.at[index, "CHECKIN_CD"] = driver.find_element(By.XPATH, "//div[@class='boxFloat'][.//span[contains(text(),'CheckIn CD')]]").find_element(By.CLASS_NAME, "data-value").text.strip()
                        df.at[index, "DOCK_DATA"] = driver.find_element(By.XPATH, "//div[@class='boxFloat'][.//span[contains(text(),'Dock')]]").find_element(By.CLASS_NAME, "data-value").text.strip()
                    except: pass

                elif tipo == "ENTREGA":
                    aba_entregas = wait.until(EC.element_to_be_clickable((By.ID, "tabEntrega")))
                    aba_entregas.click()
                    time.sleep(3)

                    # Captura conforme Parte2_UnileverEnt.py
                    try:
                        df.at[index, "DME"] = driver.find_element(By.XPATH, "//div[@id='entregasnf']//tbody/tr[last()]/td[8]").text.strip()
                        df.at[index, "CHEGADA_CLIENTE"] = driver.find_element(By.XPATH, "//div[@id='entregasnf']//tbody/tr[last()]/td[10]").text.strip()
                    except: pass

                driver.find_element(By.TAG_NAME, "body").send_keys(Keys.ESCAPE)
                time.sleep(1)
            except: continue

    finally:
        driver.quit()
    return df

if __name__ == "__main__":
    base = preparar_dados()
    if base is not None:
        base_processada = realizar_consultas_portal(base)
        
        # Formatação das colunas de data para o padrão original
        base_processada['PROGRAMAÇÃO. COLETA'] = base_processada['DT_COLETA_OBJ'].dt.strftime('%d/%m/%Y %H:%M')
        base_processada['PROGRAMAÇÃO. ENTREGA'] = base_processada['DT_ENTREGA_OBJ'].dt.strftime('%d/%m/%Y %H:%M')

        # LISTA EXATA DE COLUNAS DOS DOIS CÓDIGOS ORIGINAIS
        colunas_final = [
            'ID', 'TIPO', 'PLACA', 'MOTORISTA', 'HORA_CHEGADA', 'CHECKIN_CD', 'DOCK_DATA', 
            'DME', 'CHEGADA_CLIENTE', 'PROGRAMAÇÃO. COLETA', 'PROGRAMAÇÃO. ENTREGA', 'ORIGEM', 'DESTINO'
        ]
        
        # Filtra apenas as colunas que existem e estão na ordem desejada
        resultado = base_processada[[c for c in colunas_final if c in base_processada.columns]]
        
        resultado.to_excel(ARQUIVO_SAIDA, index=False)
        print(f"✅ Arquivo atualizado com as colunas corretas: {ARQUIVO_SAIDA}")