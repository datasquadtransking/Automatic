import pandas as pd
import numpy as np
import os
import time
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# --- CONFIGURAÇÕES DE CAMINHO ---
PASTA_PROJETO = os.path.dirname(os.path.abspath(__file__))
# O arquivo que já contém a coluna TIPO (Coleta/Entrega)
ARQUIVO_BASE = os.path.join(PASTA_PROJETO, "Base_Chegada_Unilever.xlsx")

def processar_entregas_unilever():
    print(f"📂 Passo 1: Lendo a base Base_Chegada_Unilever...")

    try:
        # 1. Leitura da base existente
        df = pd.read_excel(ARQUIVO_BASE, engine='openpyxl')
        
        # 2. Filtrar apenas o que é ENTREGA
        # Verificamos se a coluna TIPO existe e filtramos
        if 'TIPO' not in df.columns:
            print("❌ Erro: Coluna 'TIPO' não encontrada na base.")
            return

        df_entregas = df[df['TIPO'] == 'ENTREGA'].copy()

        if df_entregas.empty:
            print("⚠️ Nenhuma linha de 'ENTREGA' encontrada para processar.")
            return

        # Preparar colunas de destino se não existirem
        if "DME" not in df.columns: df["DME"] = "---"
        if "CHEGADA_CLIENTE" not in df.columns: df["CHEGADA_CLIENTE"] = "---"

        # Criar ID de consulta (limpando a barra se houver)
        df_entregas["ID_CONSULTA"] = df_entregas["ID"].astype(str).apply(lambda x: x.split("/")[-1])

        print(f"🌐 Passo 2: Consultando {len(df_entregas)} entregas no BrasilRisk...")
        
        # 3. Chamar a função de consulta
        df_final = consultar_entregas_portal(df, df_entregas)

        # 4. Salvar de volta na mesma planilha
        df_final.to_excel(ARQUIVO_BASE, index=False)
        print(f"✅ Processo de Entregas concluído! Arquivo atualizado: {ARQUIVO_BASE}")

    except Exception as e:
        print(f"❌ Erro Geral: {e}")

def consultar_entregas_portal(df_completo, df_somente_entregas):
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    driver = webdriver.Chrome(options=options)
    wait = WebDriverWait(driver, 20)

    try:
        # Login
        driver.get("https://unilever2.brasilrisk.com.br/Login/Logout")
        wait.until(EC.presence_of_element_located((By.ID, "usuario"))).send_keys("MONITORAMENTO1")
        driver.find_element(By.ID, "senha").send_keys("@Tk@2026")
        driver.find_element(By.ID, "Login").click()
        
        # Abrir busca
        wait.until(EC.element_to_be_clickable((By.ID, "botaosearch"))).click()
        time.sleep(1)

        for index, row in df_somente_entregas.iterrows():
            id_viagem = row["ID_CONSULTA"]
            print(f"🔍 Consultando Entrega ID: {id_viagem}")

            try:
                campo_search = driver.find_element(By.ID, "input-search")
                campo_search.clear()
                campo_search.send_keys(id_viagem)
                campo_search.send_keys(Keys.RETURN)
                
                # Clicar na aba "Entregas"
                aba_entregas = wait.until(EC.element_to_be_clickable((By.ID, "tabEntrega")))
                aba_entregas.click()
                time.sleep(3) # Tempo para carregar a tabela interna

                # Captura dos valores (DME na coluna 8 e Chegada na coluna 10)
                try:
                    # Tenta pegar direto da última linha da tabela de entregas
                    valor_dme = driver.find_element(By.XPATH, "//div[@id='entregasnf']//tbody/tr[last()]/td[8]").text.strip()
                    valor_chegada = driver.find_element(By.XPATH, "//div[@id='entregasnf']//tbody/tr[last()]/td[10]").text.strip()
                except:
                    valor_dme = "N/D"
                    valor_chegada = "N/D"

                # Grava a informação no DataFrame original usando o índice correto
                df_completo.at[index, "DME"] = valor_dme
                df_completo.at[index, "CHEGADA_CLIENTE"] = valor_chegada
                
                print(f"   📊 OK -> DME: {valor_dme} | Chegada: {valor_chegada}")

                # Fecha o modal (Esc)
                driver.find_element(By.TAG_NAME, "body").send_keys(Keys.ESCAPE)
                time.sleep(1)

            except Exception as e:
                print(f"   ⚠️ Erro no ID {id_viagem}: Verifique se a aba 'Entregas' existe.")
                df_completo.at[index, "DME"] = "ERRO"
                df_completo.at[index, "CHEGADA_CLIENTE"] = "ERRO"

    finally:
        driver.quit()
    
    return df_completo

if __name__ == "__main__":
    processar_entregas_unilever()