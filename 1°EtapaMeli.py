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


# ====================================================================
#   CONFIGURAÇÕES DO SISTEMA
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
#   FUNÇÃO: ESPERAR DOWNLOAD AUTOMATICAMENTE
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
#   FUNÇÃO DE LOGIN + TROCA DE JANELAS
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
#   INICIAR LOGIN
# ====================================================================
print("\n🚀 Iniciando login e mudança de janela...")
driver = login_ssw()


# ====================================================================
#   DIGITAR 103 e ABRIR NOVA JANELA
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
#   CALCULAR DATA -90 DIAS
# ====================================================================
data_hoje = datetime.now()
data_menos_90 = data_hoje - timedelta(days=90)
data_formatada = data_menos_90.strftime("%d%m%y")

print(f"📅 Data calculada: {data_formatada}")


# ====================================================================
#   DIGITAR A DATA NO CAMPO 14
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
#   DIGITAR "i" NO CAMPO 16
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
#   DIGITAR "e" NO CAMPO 17
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
#   CLICAR NO BOTÃO (id 20)
# ====================================================================
try:
    botao_20 = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.ID, '20'))
    )
    botao_20.click()
    print("▶️ Botão de gerar relatório clicado")
except:
    print("❌ Erro ao clicar no botão")


# ====================================================================
#   AGUARDAR DOWNLOAD
# ====================================================================
arquivo_ssw = esperar_download(
    nome_inicio="CSVssw0166TKI",
    nome_fim=".sswweb",
    pasta=PASTA_DOWNLOADS,
    timeout=120
)

if not arquivo_ssw:
    print("❌ Arquivo não encontrado, encerrando.")
    exit()


# ====================================================================
#   CARREGAR O ARQUIVO E CONVERTER PARA XLSX
# ====================================================================
df = pd.read_csv(arquivo_ssw, sep=";", encoding="latin1")

xlsx_final = os.path.join(PASTA_DESKTOP, "Baixar dedicado.xlsx")


# ====================================================================
#   MOSTRAR QUANTIDADE DE COLETAS POR PAGADOR
# ====================================================================
print("\n📊 QUANTIDADE DE COLETAS POR PAGADOR:")
print(df["PAGADOR"].value_counts(dropna=False))
print(f"\n🔢 TOTAL DE COLETAS: {len(df)}\n")


# ====================================================================
#   FILTRAR SOMENTE OS PAGADORES SOLICITADOS
# ====================================================================
chaves_pagador = ["AMAZON", "BRSP02", "MERCADO ENVIOS", "SHPX"]
regex_pagadores = "|".join(chaves_pagador)

df = df[
    df["PAGADOR"].fillna("").str.contains(regex_pagadores, case=False, regex=True)
]

df = df[
    (df["ULTIMA_OCORRENCIA"] != "11 - VALIDACAO CONCLUIDA") &
    (df["ULTIMA_OCORRENCIA"] != "99 - COLETA CANCELADA")
]

df = df[["NUMERO_COLETA", "PAGADOR", "ULTIMA_OCORRENCIA", "OBSERVACAO_2"]]


# -------------------------------------------------------------------------
# TRATAMENTO DA COLUNA OBSERVACAO_2 (DIVIDIR EM COLETA E ENTREGA)
# -------------------------------------------------------------------------

df[["COLETA_RAW", "ENTREGA_RAW"]] = df["OBSERVACAO_2"].str.split(" - ", n=1, expand=True)

df["COLETA_RAW"] = df["COLETA_RAW"].fillna("")
df["ENTREGA_RAW"] = df["ENTREGA_RAW"].fillna("")

df["COLETA_RAW"] = df["COLETA_RAW"].str.replace("COLETA: ", "", regex=False)
df["ENTREGA_RAW"] = df["ENTREGA_RAW"].str.replace("ENTREGA: ", "", regex=False)


# -------------------------------------------------------------------------
# EXTRAÇÃO DAS DATAS
# -------------------------------------------------------------------------

df["DATA_COLETA_DT"] = df["COLETA_RAW"].str[:10]
df["DATA_ENTREGA_DT"] = df["ENTREGA_RAW"].str[:10]


# -------------------------------------------------------------------------
# CONVERTER DATAS PARA DATETIME
# -------------------------------------------------------------------------

df["DATA_COLETA_DT"] = pd.to_datetime(
    df["DATA_COLETA_DT"], format="%d/%m/%Y", errors="coerce"
)

df["DATA_ENTREGA_DT"] = pd.to_datetime(
    df["DATA_ENTREGA_DT"], format="%d/%m/%Y", errors="coerce"
)


# -------------------------------------------------------------------------
# CÁLCULO DIAS PARA ENTREGA
# -------------------------------------------------------------------------

hoje = pd.Timestamp.today().normalize()

df["DIAS_PARA_ENTREGA"] = (df["DATA_ENTREGA_DT"] - hoje).dt.days

df["DIAS_PARA_ENTREGA"] = df["DIAS_PARA_ENTREGA"].fillna(-999).astype(int)


# ====================================================================
#   🎯 FILTRO SOLICITADO: MANTER APENAS DIAS_PARA_ENTREGA <= -8
# ====================================================================
df = df[(df["DIAS_PARA_ENTREGA"] <= -8) & (df["DIAS_PARA_ENTREGA"] > -800)]


# ====================================================================
#   SALVAR XLSX
# ====================================================================
df.to_excel(xlsx_final, index=False)
print(f"💾 Arquivo final salvo em: {xlsx_final}")


# ====================================================================
#   REMOVER TODOS .sswweb DA PASTA DOWNLOADS
# ====================================================================
try:
    for arq in os.listdir(PASTA_DOWNLOADS):
        if arq.endswith(".sswweb"):
            os.remove(os.path.join(PASTA_DOWNLOADS, arq))
    print("🗑 Todos os arquivos .sswweb foram removidos da pasta Downloads.")
except Exception as e:
    print("⚠ Erro ao remover arquivos .sswweb:", e)


print("\n✅ PROCESSO FINALIZADO COM SUCESSO!")
input("Pressione ENTER para sair...")
