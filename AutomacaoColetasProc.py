import os
import time
import shutil
import pandas as pd
from datetime import datetime, timedelta
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import sqlite3


# ====================================================================
# CONFIGURAÇÕES DO SISTEMA
# ====================================================================
LOGIN_URL = "http://sistema.ssw.inf.br"
DOM = 'TKI'
CPF = '11744601640'
USUARIO = 'weslley'
SENHA = 'W@tki111'

ID_DOM = '1'
ID_CPF = '2'
ID_USUARIO = '3'
ID_SENHA = '4'
ID_BOTAO_LOGIN = '5'

# ===== CONFIG PASTAS =====
PASTA_DOWNLOADS = os.path.join(os.path.expanduser("~"), "Downloads")
PASTA_DESKTOP = os.path.join(os.path.expanduser("~"), "Desktop")

# ====================================================================
# FUNÇÃO: ESPERAR DOWNLOAD AUTOMATICAMENTE
# ====================================================================
def esperar_download(nome_inicio, nome_fim, pasta, timeout=120):
    print("⏳ Aguardando download finalizar...")
    inicio = time.time()

    while True:
        for arquivo in os.listdir(pasta):
            if arquivo.startswith(nome_inicio) and arquivo.endswith(nome_fim):
                caminho = os.path.join(pasta, arquivo)
                print(f"📥 Arquivo encontrado: {caminho}")
                return caminho

        if time.time() - inicio > timeout:
            print("❌ Tempo limite atingido! Arquivo não encontrado.")
            return None
        time.sleep(1)

# ====================================================================
# FUNÇÃO DE LOGIN + TROCA DE JANELAS
# ====================================================================
def login_ssw():
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service)
    driver.maximize_window()
    driver.get(LOGIN_URL)

    WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.ID, ID_DOM)))

    driver.find_element(By.ID, ID_DOM).send_keys(DOM)
    driver.find_element(By.ID, ID_CPF).send_keys(CPF)
    driver.find_element(By.ID, ID_USUARIO).send_keys(USUARIO)
    driver.find_element(By.ID, ID_SENHA).send_keys(SENHA)

    driver.find_element(By.ID, ID_BOTAO_LOGIN).click()
    print("✅ Login efetuado")

    time.sleep(2)
    handles = driver.window_handles
    if len(handles) > 1:
        driver.switch_to.window(handles[-1])
        print("🔄 1ª nova janela aberta")
    else:
        print("⚠️ Nenhuma nova janela após login!")

    return driver

# ====================================================================
# INICIAR LOGIN
# ====================================================================
print("\n🚀 Iniciando login e mudança de janela...")
driver = login_ssw()

# ====================================================================
# DIGITAR 103 e ABRIR NOVA JANELA
# ====================================================================
try:
    campo_103 = WebDriverWait(driver, 15).until(
        EC.presence_of_element_located((By.ID, '3'))
    )
    campo_103.clear()
    campo_103.send_keys("103")
    print("✅ Campo '3' preenchido com 103")
except Exception as e:
    print("❌ Erro no campo 103:", e)

time.sleep(2)

handles = driver.window_handles
if len(handles) > 1:
    driver.switch_to.window(handles[-1])
    print("🔄 2ª nova janela aberta")
else:
    print("⚠️ 2ª janela não detectada!")

# ====================================================================
# CALCULAR DATA -120 DIAS
# ====================================================================
data_hoje = datetime.now()
data_menos_120 = data_hoje - timedelta(days=120)
data_formatada = data_menos_120.strftime("%d%m%y")

print(f"📅 Data calculada: {data_formatada}")

# ====================================================================
# DIGITAR A DATA NO CAMPO 14
# ====================================================================
try:
    campo_14 = WebDriverWait(driver, 15).until(
        EC.presence_of_element_located((By.ID, '14'))
    )
    campo_14.clear()
    campo_14.send_keys(data_formatada)
    print("✅ Campo 14 preenchido")
except Exception as e:
    print("❌ Erro campo 14:", e)

# ====================================================================
# DIGITAR "i" NO CAMPO 16
# ====================================================================
try:
    campo_16 = WebDriverWait(driver, 15).until(
        EC.presence_of_element_located((By.ID, '16'))
    )
    campo_16.clear()
    campo_16.send_keys("i")
    print("✅ Campo 16 preenchido com 'i'")
except:
    print("❌ Erro no campo 16")

# ====================================================================
# DIGITAR "e" NO CAMPO 17
# ====================================================================
try:
    campo_17 = WebDriverWait(driver, 15).until(
        EC.presence_of_element_located((By.ID, '17'))
    )
    campo_17.clear()
    campo_17.send_keys("e")
    print("✅ Campo 17 preenchido com 'e'")
except:
    print("❌ Erro no campo 17")

# ====================================================================
# CLICAR NO BOTÃO (id 20)
# ====================================================================
try:
    botao_20 = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.ID, '20'))
    )
    botao_20.click()
    print("▶️ Botão de gerar relatório clicado")
except:
    print("❌ Erro ao clicar no botão")
time.sleep(10)

# ====================================================================
# AGUARDAR DOWNLOAD
# ====================================================================
arquivo_ssw = esperar_download(
    nome_inicio="CSVssw0166TKI",
    nome_fim=".sswweb",
    pasta=PASTA_DOWNLOADS,
    timeout=120
)

if not arquivo_ssw:
    print("❌ Arquivo não encontrado, encerrando.")
    driver.quit()
    exit()

# ====================================================================
# ====================================================================
# CARREGAR O ARQUIVO E TRATAR DADOS
# ====================================================================
print("⚙️ Processando dados e aplicando novas regras...")

# Lendo o CSV
df = pd.read_csv(arquivo_ssw, sep=";", encoding="latin1", dtype=str)

# 1. Filtro: Coluna "1" deve ser igual a 3
df = df[df["1"] == "3"].copy()

# 2. Filtros de Pagadores (Remover indesejados)
chaves_remover = ["AMAZON", "BRSP02", "MERCADO ENVIOS", "SHPX"]
regex_pagadores = "|".join(chaves_remover)
df = df[~df["PAGADOR"].fillna("").str.contains(regex_pagadores, case=False, regex=True)].copy()

# 3. Formatar NUMERO_COLETA como Inteiro
# preenchemos vazios com 0 para não dar erro na conversão
df["NUMERO_COLETA"] = pd.to_numeric(df["NUMERO_COLETA"], errors='coerce').fillna(0).astype(int)

# 4. Tratamento de Datas e Lógica de SLA
def processar_sla_avancado(df):
    # Criamos objetos DATETIME reais para comparação
    dt_prazo = pd.to_datetime(df["COLETAR_DATA"] + " " + df["COLETAR_HORA"], dayfirst=True, errors='coerce')
    dt_realizado = pd.to_datetime(df["COLETADO_DATA"] + " " + df["COLETADO_HORA"], dayfirst=True, errors='coerce')

    # Criamos as colunas visuais formatadas para o Excel
    df["DATA_HORA_COLETAR"] = dt_prazo.dt.strftime('%d/%m/%Y %H:%M')
    df["DATA_HORA_COLETADO"] = dt_realizado.dt.strftime('%d/%m/%Y %H:%M')

    # Função interna para aplicar as regras linha por linha
    def checar_regra(row):
        idx = row.name
        pago = dt_prazo[idx]
        real = dt_realizado[idx]
        pagador = str(row["PAGADOR"]).upper()

        # Se não houver data de coleta, precisa verificar
        if pd.isna(pago) or pd.isna(real):
            return "VERIFICAR"

        # REGRA GERAL: Coletado antes ou igual ao prazo (Data e Hora)
        if real <= pago:
            return "NO PRAZO"

        # REGRA LG: Se contiver LG e for no MESMO DIA (ignora a hora)
        if "LG" in pagador:
            if real.date() == pago.date():
                return "NO PRAZO"

        return "VERIFICAR"

    df["SITUACAO_COLETA"] = df.apply(checar_regra, axis=1)
    return df

df = processar_sla_avancado(df)

# Reordenar colunas para facilitar a leitura
cols_foco = ["SITUACAO_COLETA", "NUMERO_COLETA", "PAGADOR", "DATA_HORA_COLETAR", "DATA_HORA_COLETADO"]
outras = [c for c in df.columns if c not in cols_foco]
df = df[cols_foco + outras]

# ====================================================================
# SALVAR EM EXCEL (PRESERVANDO OUTRAS ABAS)
# ====================================================================
xlsx_final = os.path.join(PASTA_DESKTOP, "ColetasAut.xlsx")


# ====================================================================
# SALVAR EM BANCO SQL (SQLite)
# ====================================================================
caminho_db = os.path.join(PASTA_DESKTOP, "ColetasAut.db")

try:
    conn = sqlite3.connect(caminho_db)

    # Salva o DataFrame como tabela SQL
    df.to_sql(
        name="base_coletas",
        con=conn,
        if_exists="replace",  # substitui toda vez que rodar
        index=False
    )

    conn.close()

    print(f"✅ Banco SQL criado/atualizado com sucesso em: {caminho_db}")

except Exception as e:
    print("❌ Erro ao salvar no banco SQL:", e)

try:
    if os.path.exists(xlsx_final):
        # Abre o arquivo e substitui apenas a aba Base_Coletas
        with pd.ExcelWriter(xlsx_final, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df.to_excel(writer, sheet_name='Base_Coletas', index=False)
        print(f"✅ Aba 'Base_Coletas' atualizada com sucesso!")
    else:
        df.to_excel(xlsx_final, sheet_name='Base_Coletas', index=False)
        print(f"✅ Novo arquivo criado na Área de Trabalho.")

    # === NOVA LINHA ADICIONADA AQUI ===
    if os.path.exists(arquivo_ssw):
        os.remove(arquivo_ssw)
        print(f"🗑️ Arquivo temporário {arquivo_ssw} removido da pasta Downloads.")
    # =================================

except PermissionError:
    print("❌ ERRO: Feche o arquivo Excel antes de rodar o código!")