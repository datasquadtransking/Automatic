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
# ATENÇÃO: Verifique se este é o caminho correto no seu computador.
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
# Filtrar onde SLA é NaN, ímpar, ou não é 0 nem 1
filtro_sla_entrega = (df_entrega["SLA"].isna()) | ((df_entrega["SLA"] != 1) & (df_entrega["SLA"] != 0))
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

    # Usando a mesma URL e campos do seu script original
    driver.get("https://unilever2.brasilrisk.com.br/Login/Logout")

    # Esperar os campos de login carregarem
    WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.ID, "usuario"))
    )

    usuario = driver.find_element(By.ID, "usuario")
    usuario.send_keys("MONITORAMENTO1")

    senha = driver.find_element(By.ID, "senha")
    senha.send_keys("Transking@2026")

    login_btn = driver.find_element(By.ID, "Login")
    login_btn.click()

    time.sleep(2) # Espera após login

    # Clicar no botão de busca (lupa) para abrir o campo de pesquisa
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
    consultas_entrega = [] # Lista para armazenar o log de consultas

    for index, row in df_entrega.iterrows():
        id_viagem = str(row["ID"])
        # Obtém o índice original no DataFrame completo 'df' para atualização
        idx_original = row.name 
        
        processadas += 1
        faltam = total_entregas - processadas
        print(f"\n-> Processando ENTREGA ID: {id_viagem}")
        print(f"Processadas: {processadas} / {total_entregas}  |  Faltam: {faltam}")

        # PESQUISAR O ID
        try:
            campo_search.clear()
        except Exception:
            # Em caso de falha no .clear(), tenta reobter o elemento
            campo_search = driver.find_element(By.ID, "input-search")
            campo_search.clear()

        campo_search.send_keys(id_viagem)
        time.sleep(1)
        campo_search.send_keys(Keys.RETURN)
        time.sleep(2) # Tempo para a busca carregar

        # CLICAR NA ABA "ENTREGAS"
        try:
            aba_entregas = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, "tabEntrega"))
            )
            aba_entregas.click()
            # Tempo de espera para o carregamento da tabela na aba
            time.sleep(1) 
        except (NoSuchElementException, TimeoutException) as e:
            print(f"ERRO: Não foi possível abrir a aba 'Entregas' para o ID {id_viagem}.")
            # Registra o erro no DataFrame (Embora o df não seja salvo, o log de consultas será)
            df.loc[idx_original, "CHECK_ENTREGA"] = "SEM_ABA_ENTREGAS" 
            consultas_entrega.append({"ID": id_viagem, "CHECK_ENTREGA": "SEM_ABA_ENTREGAS"})
            continue

        # ===============================================================
        # 5) CAPTURAR A ÚLTIMA LINHA da tabela de entregas (10ª coluna)
        # ===============================================================
        valor_chegada = None
        
        try:
            # XPath para selecionar a 10ª coluna (td[10]) da ÚLTIMA linha (tr[last()])
            xpath_valor_chegada = "//div[@id='entregasnf']//tbody/tr[last()]/td[10]"
            
            # Espera até que o elemento da última linha esteja presente
            elemento_valor_chegada = WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.XPATH, xpath_valor_chegada))
            )
            
            valor_chegada = elemento_valor_chegada.text.strip()
            
            if not valor_chegada:
                raise Exception("Valor encontrado por XPath da última linha estava vazio.")
            
            print(f"Valor 'Chegada no cliente' da ÚLTIMA linha encontrado: {valor_chegada}")

        except Exception as e_xpath:
            # Fallback: Caso o XPath direto falhe, tenta localizar todas as linhas
            print(f"Falha na captura por XPath direto da última linha. Tentando fallback.")
            try:
                linhas = WebDriverWait(driver, 10).until(
                    EC.presence_of_all_elements_located((By.XPATH, "//div[@id='entregasnf']//tbody/tr"))
                )
                
                if linhas:
                    # Pega a última linha da lista
                    ultima_linha = linhas[-1]
                    colunas = ultima_linha.find_elements(By.TAG_NAME, "td")
                    
                    if len(colunas) >= 10:
                        # Coluna 10 (Chegada no cliente) => índice 9
                        valor_chegada = colunas[9].text.strip()
                        print(f"Valor 'Chegada no cliente' da ÚLTIMA linha encontrado via Fallback: {valor_chegada}")
                    else:
                        valor_chegada = "COLUNA_10_NAO_EXISTE"
                        print(f"A última linha não tem 10 colunas para ID {id_viagem}.")
                else:
                    valor_chegada = "TABELA_VAZIA"
                    print(f"Tabela de entregas está vazia para ID {id_viagem}.")
            
            except Exception as e_fallback:
                print(f"Erro no Fallback para ID {id_viagem}: {e_fallback}")
                df.loc[idx_original, "CHECK_ENTREGA"] = "ERRO_COLETA"
                consultas_entrega.append({"ID": id_viagem, "CHECK_ENTREGA": "ERRO_COLETA"})
                continue
        
        # Lógica de registro do resultado (após as tentativas)
        if valor_chegada is None or valor_chegada in ("COLUNA_10_NAO_EXISTE", "TABELA_VAZIA", "VAZIO"):
            status = valor_chegada if valor_chegada else "SEM_VALOR_FINAL"
            print(f"Valor de Chegada no cliente não encontrado/inconsistente. Status: {status}")
            df.loc[idx_original, "CHECK_ENTREGA"] = status
            consultas_entrega.append({"ID": id_viagem, "CHECK_ENTREGA": status})
        else:
            print(f"Valor final 'Chegada no cliente' registrado: {valor_chegada}")
            df.loc[idx_original, "CHECK_ENTREGA"] = valor_chegada
            consultas_entrega.append({"ID": id_viagem, "CHECK_ENTREGA": valor_chegada})


        # Opcional: fechar modal (se houver) — tentativa por ESC
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
# 6) Salvar o log de consultas na mesma planilha (APENAS EntregasConsultadas)
# ===============================================================

df_consultas = pd.DataFrame(consultas_entrega)

# Salvamos APENAS o log de consultas na aba 'EntregasConsultadas'
try:
    # Usamos 'openpyxl' para manter o arquivo e substituir APENAS a aba 'EntregasConsultadas'
    with pd.ExcelWriter(CAMINHO_PLANILHA, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        df_consultas.to_excel(writer, sheet_name="EntregasConsultadas", index=False)
    print("\nAba 'EntregasConsultadas' (Log de Consultas) salva com sucesso no arquivo principal!")
except Exception as e:
    print(f"\nERRO ao salvar a aba 'EntregasConsultadas': {e}")

# Exibir as primeiras linhas do Log de Consultas
print("\nLog de Consultas (EntregasConsultadas):")
print(df_consultas.head())