
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
# CONFIGURAÇÕES DO SISTEMA
# ====================================================================
LOGIN_URL = "http://sistema.ssw.inf.br"
DOM = 'TKI'
CPF = '10620126663'
USUARIO = 'automac'
SENHA = 'A@tki123'

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
# CARREGAR O ARQUIVO E TRATAR DADOS
# ====================================================================
print("⚙️ Processando dados e aplicando novas regras...")

df = pd.read_csv(arquivo_ssw, sep=";", encoding="latin1", dtype=str)

# 1. Filtro: Coluna "1" deve ser igual a 3
df = df[df["1"] == "3"].copy()

# 2. Filtros de Pagadores
chaves_remover = ["AMAZON", "BRSP02", "MERCADO ENVIOS", "SHPX","SAL EXPRESS SOLUCOES E TRANSPO"]
regex_pagadores = "|".join(chaves_remover)
df = df[~df["PAGADOR"].fillna("").str.contains(regex_pagadores, case=False, regex=True)].copy()

# 3. Formatar NUMERO_COLETA
df["NUMERO_COLETA"] = pd.to_numeric(df["NUMERO_COLETA"], errors='coerce').fillna(0).astype(int)

# 4. Tratamento de Datas e Lógica de SLA
def processar_sla_avancado(df):
    dt_prazo = pd.to_datetime(df["COLETAR_DATA"] + " " + df["COLETAR_HORA"], dayfirst=True, errors='coerce')
    dt_realizado = pd.to_datetime(df["COLETADO_DATA"] + " " + df["COLETADO_HORA"], dayfirst=True, errors='coerce')
    dt_UltOco = pd.to_datetime(df["DATA_ULTIMA_OCORRENCIA"] + " " + df["HORA_ULTIMA_OCORRENCIA"], dayfirst=True, errors='coerce')


    df["DATA_HORA_COLETAR"] = dt_prazo.dt.strftime('%d/%m/%Y %H:%M')
    df["DATA_HORA_COLETADO"] = dt_realizado.dt.strftime('%d/%m/%Y %H:%M')
    df["DATA_HORA_ULTIMAOCORRENCIA"] = dt_realizado.dt.strftime('%d/%m/%Y %H:%M')


    def checar_regra(row):
        idx = row.name
        pago = dt_prazo[idx]
        real = dt_realizado[idx]
        pagador = str(row["PAGADOR"]).upper()

        if pd.isna(pago) or pd.isna(real):
            return "VERIFICAR"

        if real <= pago:
            return "NO PRAZO"

        if "LG" in pagador:
            if real.date() == pago.date():
                return "NO PRAZO"

        return "VERIFICAR"

    df["SITUACAO_COLETA"] = df.apply(checar_regra, axis=1)
    return df

df = processar_sla_avancado(df)

# CARREGAR O ARQUIVO E TRATAR DADOS
# ====================================================================
print("⚙️ Processando dados e aplicando novas regras...")

df = pd.read_csv(arquivo_ssw, sep=";", encoding="latin1", dtype=str)

# 1. Filtro: Coluna "1" deve ser igual a 3
df = df[df["1"] == "3"].copy()

# 2. Renomear PEDIDO para COTAÇÃO (se existir)
if "PEDIDO" in df.columns:
    df = df.rename(columns={"PEDIDO": "COTAÇÃO"})

# 3. Filtros de Pagadores
chaves_remover = ["AMAZON", "BRSP02", "MERCADO ENVIOS", "SHPX"]
regex_pagadores = "|".join(chaves_remover)
df = df[~df["PAGADOR"].fillna("").str.contains(regex_pagadores, case=False, regex=True)].copy()

# 4. Tratamento de Datas e Lógica de SLA
df = processar_sla_avancado(df)

# ====================================================================
# NOVAS CONCATENAÇÕES (Data/Hora, Origem, Destino e Rota)
# ====================================================================
print("🔗 Agrupando colunas de Data, Origem e Destino...")

# Juntar DATA_INCLUSAO e HORA_INCLUSAO
df["DATA_HORA_INCLUSAO"] = df["DATA_INCLUSAO"].fillna("") + " " + df["HORA_INCLUSAO"].fillna("")

# Juntar Origem e Destino
df["Origem"] = df["CIDADE_REMETENTE"].fillna("") + "/" + df["UF_REMETENTE"].fillna("")
df["DESTINO"] = df["CIDADE_DESTINO"].fillna("") + "/" + df["UF_DESTINO"].fillna("")

# Criar ROTA
df["ROTA"] = df["Origem"] + " - " + df["DESTINO"]

# ====================================================================
# MAPEAMENTO DE STATUS
# ====================================================================
print("📊 Gerando coluna Status...")

mapeamento_status = {
    "1 - COLETA REALIZADA (CARREGADA NO VEICULO)": "3. Aguardando Coleta",
    "10 - EM ROTA DE ENTREGA": "6. Em Rota",
    "17 - CHEGADA NA ENTREGA": "7. Aguardando Descarga",
    "2 - COLETA COMANDADA AO VEICULO": "2. Em GR",
    "20 - DEVOLUCAO TOTAL": "Outros",
    "25 - COMPROVANTE RETIDO": "8. Aguardando Comprovante",
    "3 - DOC EMITIDO": "5. Aguardando A.E",
    "30 - ENTREGA REALIZADA": "Validação Final",
    "33 - ESTADIA - ENTREGA": "Outros",
    "34 - A.E. EMITIDA": "6. Em Rota",
    "50 - STATUS CARREGAMENTO": "3. Aguardando Coleta",
    "82 - APROVADO GR": "3. Aguardando Coleta",
    "79 - EM PESQUISA GR": "2. Em GR",
    "81 - STATUS GR": "2. Em GR",
    "97 - DESCOMANDADA": "1. Aguardando Planejamento",
    "98 - COLETA CADASTRADA": "1. Aguardando Planejamento"
}

# Criando a coluna Status com segurança
if "ULTIMA_OCORRENCIA" in df.columns:
    df["Status"] = df["ULTIMA_OCORRENCIA"].map(mapeamento_status).fillna("Outros")
else:
    df["Status"] = "Verificar Ocorrência"

# ====================================================================
# APAGAR COLUNAS INDIVIDUAIS E LIMPEZA FINAL
# ====================================================================
colunas_para_remover = [
    "DATA_INCLUSAO", "HORA_INCLUSAO", 
    "CIDADE_REMETENTE", "UF_REMETENTE", 
    "CIDADE_DESTINO", "UF_DESTINO",
    "11 - VALIDACAO CONCLUIDA", "99 - COLETA CANCELADA", # Caso existissem como colunas
    "22 - OPORTUNIDADE PERDIDA", "8 - CARGA DECLINADA PELO CLIENTE"
]
df = df.drop(columns=[c for c in colunas_para_remover if c in df.columns])

# Organização de colunas (Garantindo que 'Status' existe antes de organizar)
cols_foco = ["Status", "SITUACAO_COLETA", "ROTA", "COTAÇÃO", "NUMERO_COLETA", "PAGADOR"]
cols_foco = [c for c in cols_foco if c in df.columns] # Só mantém as que realmente existem
outras = [c for c in df.columns if c not in cols_foco]
df = df[cols_foco + outras]

# ====================================================================
# EXCLUIR OCORRÊNCIAS E USUÁRIOS INDESEJADOS
# ====================================================================
print("🧹 Removendo ocorrências indesejadas e usuários...")

status_para_remover = [
    "11 - VALIDACAO CONCLUIDA", 
    "99 - COLETA CANCELADA",
    "22 - OPORTUNIDADE PERDIDA",
    "8 - CARGA DECLINADA PELO CLIENTE"
]
usuarios_para_remover = ["matheusc", "robson"]

df = df[
    (~df["ULTIMA_OCORRENCIA"].isin(status_para_remover)) & 
    (~df["USUARIO"].isin(usuarios_para_remover))
].copy()

# ====================================================================
# NOVO: MAPEAMENTO DE STATUS (CONFORME TABELA DA FOTO)
# ====================================================================
print("📊 Mapeando coluna Status baseada na ULTIMA_OCORRENCIA...")

mapeamento_status = {
    "1 - COLETA REALIZADA (CARREGADA NO VEICULO)": "3. Aguardando Coleta",
    "10 - EM ROTA DE ENTREGA": "6. Em Rota",
    "17 - CHEGADA NA ENTREGA": "7. Aguardando Descarga",
    "2 - COLETA COMANDADA AO VEICULO": "2. Em GR",
    "20 - DEVOLUCAO TOTAL": "Outros",
    "25 - COMPROVANTE RETIDO": "8. Aguardando Comprovante",
    "3 - DOC EMITIDO": "5. Aguardando A.E",
    "30 - ENTREGA REALIZADA": "Validação Final",
    "33 - ESTADIA - ENTREGA": "Outros",
    "34 - A.E. EMITIDA": "6. Em Rota",
    "50 - STATUS CARREGAMENTO": "3. Aguardando Coleta",
    "82 - APROVADO GR": "3. Aguardando Coleta",
    "79 - EM PESQUISA GR": "2. Em GR",
    "81 - STATUS GR": "2. Em GR",
    "97 - DESCOMANDADA": "1. Aguardando Planejamento",
    "98 - COLETA CADASTRADA": "1. Aguardando Planejamento"
}

# Cria a coluna Status. Se não encontrar o valor no dicionário, coloca "Outros"
df["Status"] = df["ULTIMA_OCORRENCIA"].map(mapeamento_status).fillna("Outros")

# ====================================================================
# ORGANIZAÇÃO FINAL DE COLUNAS
# ====================================================================
# Adicionada a nova coluna 'Status' no início do foco
cols_foco = ["Status", "SITUACAO_COLETA", "NUMERO_COLETA", "PAGADOR", "DATA_HORA_COLETAR", "DATA_HORA_COLETADO"]
outras = [c for c in df.columns if c not in cols_foco]
df = df[cols_foco + outras]


# SELEÇÃO E ORDENAÇÃO FINAL DE COLUNAS (CONFORME SOLICITADO)
# ====================================================================
print("🎯 Filtrando e ordenando colunas finais...")

# Lista exata das colunas desejadas na ordem correta
colunas_finais = [
    "Status", "DATA_HORA_INCLUSAO",  "NUMERO_COLETA",  "COTAÇÃO", "PAGADOR",  "ROTA","SITUACAO", 
    "DATA_HORA_COLETAR", "DATA_HORA_COLETADO","USUARIO_ULTIMA_OCORRENCIA", "ULTIMA_OCORRENCIA", "INSTRUCAO_SITUACAO", 
    "NOTAS_FISCAIS", "CTRC", "DESTINATARIO", "CNPJ_DESTINATARIO", 
    "CNPJ_PAGADOR", "PESO_CALCULO(KG)", "PESO_REAL(KG)", "M3", 
    "QUANT_VOL", "VALOR_MERCADORIA", "TIPO_MERCADORIA", 
    "PLACA_CAVALO", "PLACA_CARRETA", "CPF_MOTORISTA", 
    "MOTORISTA", "Origem", "DESTINO","SITUACAO_COLETA","DATA_HORA_ULTIMAOCORRENCIA"
]

# 1. Verificar se todas as colunas existem no DF (evita erro se alguma vier vazia do SSW)
colunas_presentes = [c for c in colunas_finais if c in df.columns]

# 2. Filtrar o DataFrame: manter APENAS estas colunas e NESTA ordem
df = df[colunas_presentes].copy()

# ====================================================================
# SALVAR EM EXCEL
# ====================================================================
xlsx_final = os.path.join(PASTA_DESKTOP, "AcompanhamentoOperacionalBi.xlsx")

try:
    if os.path.exists(xlsx_final):
        with pd.ExcelWriter(xlsx_final, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df.to_excel(writer, sheet_name='Base_Coletas', index=False)
        print(f"✅ Aba 'Base_Coletas' atualizada com sucesso!")
    else:
        df.to_excel(xlsx_final, sheet_name='Base_Coletas', index=False)
        print(f"✅ Novo arquivo criado na Área de Trabalho.")

    if os.path.exists(arquivo_ssw):
        os.remove(arquivo_ssw)
        print(f"🗑️ Arquivo temporário {arquivo_ssw} removido.")

except PermissionError:
    print("❌ ERRO: Feche o arquivo Excel antes de rodar o código!")

print("\n✨ Processo concluído!")