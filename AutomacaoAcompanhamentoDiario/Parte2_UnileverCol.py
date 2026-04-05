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
ARQUIVO_SAIDA_UNILEVER = os.path.join(PASTA_PROJETO, "Base_Chegada_Unilever.xlsx")

def preparar_base_chegada_unilever():
    print(f"📂 Passo 1: Lendo base para extração Unilever...")

    try:
        # 1. Leitura do arquivo gerado pela Parte 1
        df = pd.read_excel(ARQUIVO_ENTRADA, engine='openpyxl')
        df.columns = df.columns.str.strip().str.upper()

        # 2. Filtro de Cliente
        df_unilever = df[df['CLIENTE'].str.contains('UNILEVER', na=False, case=False)].copy()

        if df_unilever.empty:
            print("⚠️ Nenhum dado da Unilever encontrado.")
            return

        # 3. Tratamento de Datas e Ajuste do Alvo (FERIADO: HOJE + 1)
        df_unilever['DT_COLETA_OBJ'] = pd.to_datetime(df_unilever['PROGRAMAÇÃO. COLETA'], dayfirst=True)
        df_unilever['DT_ENTREGA_OBJ'] = pd.to_datetime(df_unilever['PROGRAMAÇÃO. ENTREGA'], dayfirst=True)
        
        # AJUSTE PROVISÓRIO PARA AMANHÃ
        data_alvo = datetime.now().date() + timedelta(days=+3) 
        print(f"📅 Data Alvo definida para: {data_alvo.strftime('%d/%m/%Y')} (Feriado)")

        condicoes = [
            (df_unilever['DT_COLETA_OBJ'].dt.date == data_alvo),
            (df_unilever['DT_ENTREGA_OBJ'].dt.date == data_alvo)
        ]
        df_unilever['TIPO'] = np.select(condicoes, ['COLETA', 'ENTREGA'], default='OUTROS')

        base_final = df_unilever[df_unilever['TIPO'] != 'OUTROS'].copy()
        base_final["ID_CONSULTA"] = base_final["ID"].astype(str).apply(lambda x: x.split("/")[-1])
        
        # Preparando colunas de retorno do sistema
        base_final["HORA_CHEGADA"] = "---"
        base_final["CHECKIN_CD"] = "---"
        base_final["DOCK_DATA"] = "---"

        # --- PASSO 4: CONSULTA PORTAL (Apenas COLETAS) ---
        df_coletas = base_final[base_final["TIPO"] == "COLETA"].copy()
        
        if not df_coletas.empty:
            print(f"🌐 Passo 2: Consultando {len(df_coletas)} coletas no BrasilRisk...")
            base_final = consultar_portal_brasilrisk(base_final, df_coletas)
        
        # 5. Formatação das datas originais para o Excel
        base_final['PROGRAMAÇÃO. COLETA'] = base_final['DT_COLETA_OBJ'].dt.strftime('%d/%m/%Y %H:%M')
        base_final['PROGRAMAÇÃO. ENTREGA'] = base_final['DT_ENTREGA_OBJ'].dt.strftime('%d/%m/%Y %H:%M')

        # 6. Seleção de colunas e salvamento
        colunas_final = [
            'ID', 'TIPO', 'PLACA', 'MOTORISTA', 'HORA_CHEGADA', 'CHECKIN_CD', 'DOCK_DATA',
            'PROGRAMAÇÃO. COLETA', 'PROGRAMAÇÃO. ENTREGA', 'ORIGEM', 'DESTINO'
        ]
        
        base_final[[c for c in colunas_final if c in base_final.columns]].to_excel(ARQUIVO_SAIDA_UNILEVER, index=False)
        print(f"✅ Processo concluído! Arquivo gerado: {ARQUIVO_SAIDA_UNILEVER}")

    except Exception as e:
        print(f"❌ Erro Geral no Script: {e}")

def consultar_portal_brasilrisk(df_completo, df_coletas):
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    driver = webdriver.Chrome(options=options)
    wait = WebDriverWait(driver, 20) # Aumentado tempo de espera para conexões lentas

    try:
        # Login no Sistema
        driver.get("https://unilever2.brasilrisk.com.br/Login/Logout")
        wait.until(EC.presence_of_element_located((By.ID, "usuario"))).send_keys("MONITORAMENTO1")
        driver.find_element(By.ID, "senha").send_keys("@Tk@2026")
        driver.find_element(By.ID, "Login").click()
        
        # Clica na Lupa de Busca
        wait.until(EC.element_to_be_clickable((By.ID, "botaosearch"))).click()
        time.sleep(1)

        for index, row in df_coletas.iterrows():
            id_viagem = row["ID_CONSULTA"]
            print(f"🔍 Pesquisando ID: {id_viagem}")

            try:
                campo_search = driver.find_element(By.ID, "input-search")
                campo_search.clear()
                campo_search.send_keys(id_viagem)
                campo_search.send_keys(Keys.RETURN)
                
                # Abre os detalhes da viagem
                chevron = wait.until(EC.element_to_be_clickable((By.ID, "chevViagem")))
                chevron.click()
                time.sleep(2) # Espera o modal carregar as informações internas

                # --- CAPTURA COM XPATH FLEXÍVEL ---
                
                # 1. Hora de Chegada
                try:
                    el_chegada = driver.find_element(By.XPATH, "//div[@class='boxFloat'][.//span[contains(text(),'Hora de chegada')]]")
                    val_chegada = el_chegada.find_element(By.CLASS_NAME, "data-value").text.strip()
                except: val_chegada = "Vazio"

                # 2. CheckIn CD
                try:
                    el_checkin = driver.find_element(By.XPATH, "//div[@class='boxFloat'][.//span[contains(text(),'CheckIn CD')]]")
                    val_checkin = el_checkin.find_element(By.CLASS_NAME, "data-value").text.strip()
                except: val_checkin = "Vazio"

                # 3. Dock
                try:
                    el_dock = driver.find_element(By.XPATH, "//div[@class='boxFloat'][.//span[contains(text(),'Dock')]]")
                    val_dock = el_dock.find_element(By.CLASS_NAME, "data-value").text.strip()
                except: val_dock = "Vazio"

                # Grava no DataFrame principal
                df_completo.at[index, "HORA_CHEGADA"] = val_chegada
                df_completo.at[index, "CHECKIN_CD"] = val_checkin
                df_completo.at[index, "DOCK_DATA"] = val_dock
                
                print(f"   📊 Resultado -> Chegada: {val_chegada} | Checkin: {val_checkin}")

                # Fecha o modal (Esc) para liberar a próxima busca
                driver.find_element(By.TAG_NAME, "body").send_keys(Keys.ESCAPE)
                time.sleep(1)

            except Exception as e:
                print(f"   ⚠️ ID {id_viagem} não localizado ou erro na tela.")
                df_completo.at[index, "HORA_CHEGADA"] = "N/D"

    finally:
        driver.quit()
    
    return df_completo

if __name__ == "__main__":
    preparar_base_chegada_unilever()