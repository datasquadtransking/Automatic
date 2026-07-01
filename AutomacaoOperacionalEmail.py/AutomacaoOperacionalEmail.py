

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
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

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

# Marca o horário de início do download
hora_inicio_aba = datetime.now().strftime("%d-%m-%Y_%H-%M")

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

# ====================================================================
# RENOMEAR PEDIDO PARA COTAÇÃO
# ====================================================================
if "PEDIDO" in df.columns:
    df = df.rename(columns={"PEDIDO": "COTAÇÃO"})

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
    "11 - VALIDACAO CONCLUIDA", "99 - COLETA CANCELADA",
    "22 - OPORTUNIDADE PERDIDA", "8 - CARGA DECLINADA PELO CLIENTE"
]
df = df.drop(columns=[c for c in colunas_para_remover if c in df.columns])

# Organização de colunas
cols_foco = ["Status", "SITUACAO_COLETA", "ROTA", "COTAÇÃO", "NUMERO_COLETA", "PAGADOR"]
cols_foco = [c for c in cols_foco if c in df.columns]
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
# Mapeamento de Status
# ====================================================================
df["Status"] = df["ULTIMA_OCORRENCIA"].map(mapeamento_status).fillna("Outros")

cols_foco = ["Status", "SITUACAO_COLETA", "NUMERO_COLETA", "PAGADOR", "DATA_HORA_COLETAR", "DATA_HORA_COLETADO"]
outras = [c for c in df.columns if c not in cols_foco]
df = df[cols_foco + outras]

# ====================================================================
# SELEÇÃO E ORDENAÇÃO FINAL DE COLUNAS
# ====================================================================
colunas_finais = [
    "Status", "DATA_HORA_INCLUSAO", "NUMERO_COLETA", "COTAÇÃO", "PAGADOR", "ROTA", "SITUACAO", 
    "DATA_HORA_COLETAR", "DATA_HORA_COLETADO", "USUARIO_ULTIMA_OCORRENCIA", "ULTIMA_OCORRENCIA", "INSTRUCAO_SITUACAO", 
    "NOTAS_FISCAIS", "CTRC", "DESTINATARIO", "CNPJ_DESTINATARIO", 
    "CNPJ_PAGADOR", "PESO_CALCULO(KG)", "PESO_REAL(KG)", "M3", 
    "QUANT_VOL", "VALOR_MERCADORIA", "TIPO_MERCADORIA", 
    "PLACA_CAVALO", "PLACA_CARRETA", "CPF_MOTORISTA", 
    "MOTORISTA", "Origem", "DESTINO", "SITUACAO_COLETA", "DATA_HORA_ULTIMAOCORRENCIA"
]

colunas_presentes = [c for c in colunas_finais if c in df.columns]
df = df[colunas_presentes].copy()

# ====================================================================
# ORDENAR POR STATUS EM ORDEM CRESCENTE
# ====================================================================
df = df.sort_values(by='Status', ascending=True)

# SALVAR EM EXCEL
# ====================================================================
xlsx_final = os.path.join(PASTA_DESKTOP, "AcompanhamentoOperacionalBi.xlsx")
nome_da_aba = f"Base {hora_inicio_aba}"

try:
    df.to_excel(xlsx_final, sheet_name=nome_da_aba, index=False)
    print(f"✅ Planilha salva no arquivo fixo!")
    print(f"✅ Nome da aba definido como: {nome_da_aba}")

    # ====================================================================
    # COPIAR/MOVER ARQUIVO PARA OUTRA PASTA SUBSTITUINDO SE EXISTIR
    # ====================================================================
    destino_final = r"C:\Users\flavi\OneDrive\Weslley\AcompanhamentoOperacionalBi.xlsx"

    try:
        shutil.copy2(xlsx_final, destino_final)
        print(f"📁 Arquivo copiado para: {destino_final}")
    except Exception as e:
        print(f"❌ Erro ao copiar arquivo: {e}")

    if os.path.exists(arquivo_ssw):
        os.remove(arquivo_ssw)
        print(f"🗑️ Arquivo temporário removido.")

except PermissionError:
    print("❌ ERRO: O Excel está aberto! Feche o arquivo para salvar.")



# ====================================================================
# ENVIAR E-MAIL
# ====================================================================
EMAIL_REMETENTE = "automac@transking.com.br"
SENHA_EMAIL = "P3dr0Tk2025"
EMAIL_DESTINATARIO = "weslley.transking@gmail.com"
EMAILS_COPIA = " "

SMTP_SERVIDOR = "smtp.task.com.br"
SMTP_PORTA = 587

def estilizar_tabela_html(df, largura_total="100%", cor_cabecalho="#1f4e78"):
    if df.empty: return ""
    
    df_copy = df.copy()
    for col in df_copy.columns:
        if df_copy[col].dtype in ['float64', 'int64']:
             df_copy[col] = df_copy[col].fillna(0).astype(int).astype(str)

    html = f'<div style="width: 100%; overflow-x: auto;">'
    html += f'<table style="border-collapse: collapse; width: {largura_total}; font-family: Segoe UI, Arial; font-size: 10px; margin-bottom: 10px; border: 1px solid #dee2e6;">'
    html += f'<thead style="background-color: {cor_cabecalho}; color: white;"><tr>'
    
    for col in df_copy.columns:
        html += f'<th style="padding: 3px 8px; border: 1px solid #dee2e6; text-align: left; white-space: nowrap;">{col}</th>'
    html += '</tr></thead><tbody>'
    
    for _, row in df_copy.iterrows():
        html += '<tr>'
        for val in row:
            val_str = str(val) if pd.notnull(val) else ""
            html += f'<td style="padding: 2px 8px; border: 1px solid #dee2e6; text-align: left; white-space: nowrap;">{val_str}</td>'
        html += '</tr>'
    
    html += '</tbody></table></div>'
    return html


def gerar_corpo_painel(caminho_arquivo):
    df = pd.read_excel(caminho_arquivo)
    col_status = 'Status'
    col_data = 'DATA_HORA_COLETAR'
    
    df[col_data] = pd.to_datetime(df[col_data], errors='coerce')
    
    agora = datetime.now()
    hoje_meia_noite = agora.replace(hour=0, minute=0, second=0, microsecond=0)
    data_limite = (hoje_meia_noite + timedelta(days=2)).replace(hour=23, minute=59, second=59)
    
    saudacao = "Boa tarde" if 12 <= agora.hour < 18 else "Boa noite" if 18 <= agora.hour <= 23 else "Bom dia"
    
    # Ajuste: Exibe somente a data, sem a hora
    html_body = f"""
    <html>
    <body style="margin: 0; padding: 20px; font-family: Segoe UI, Arial; color: #333;">
        <p>{saudacao}, Weslley.</p>
        <h2 style="color: #1f4e78; border-bottom: 2px solid #1f4e78; padding-bottom: 5px;">
            📊 Acompanhamento Operacional - {agora.strftime('%d/%m/%Y')}
        </h2>
        <p style='font-size: 12px;'>
            Exibindo registros de Planejamento até: <b>{data_limite.strftime('%d/%m/%Y')}</b> (Hoje + 2 dias).
        </p>
    """

    df = df.sort_values(by=[col_status, col_data])
    status_unicos = df[col_status].unique()
    
    for status in status_unicos:
        if "AGUARDANDO PLANEJAMENTO" in str(status).upper():
            df_filtrado = df[(df[col_status] == status) & (df[col_data] <= data_limite)].copy()
        else:
            df_filtrado = df[df[col_status] == status].copy()
        
        if not df_filtrado.empty:
            df_exibicao = df_filtrado.copy()
            df_exibicao[col_data] = df_exibicao[col_data].dt.strftime('%d/%m/%Y %H:%M')
            
            html_body += f"""
            <div style="margin-top: 20px;">
                <h4 style="color: #c45911; margin-bottom: 5px; text-transform: uppercase;">
                    📌 ETAPA: {status} <span style="color: #777; font-weight: normal; font-size: 12px;">({len(df_exibicao)} itens)</span>
                </h4>
                {estilizar_tabela_html(df_exibicao)}
            </div>
            <hr style="border: 0; border-top: 1px solid #eee; margin: 10px 0;">
            """

    html_body += """
        <br>
        <p style="font-size: 11px; color: #999;"><i>E-mail gerado automaticamente pelo Sistema de Automação Transking.</i></p>
    </body>
    </html>
    """
    return html_body

def enviar_email_com_anexo(caminho_arquivo):
    if not os.path.exists(caminho_arquivo):
        print(f"Erro: Arquivo {caminho_arquivo} não encontrado.")
        return

    msg = MIMEMultipart()
    msg['From'] = EMAIL_REMETENTE
    msg['To'] = EMAIL_DESTINATARIO
    msg['Cc'] = EMAILS_COPIA
    # Ajuste: Assunto do e-mail sem a hora
    msg['Subject'] = f"Acompanhamento Operacional - {datetime.now().strftime('%d/%m/%Y')}"

    todos_destinatarios = [EMAIL_DESTINATARIO] + [email.strip() for email in EMAILS_COPIA.split(",")]

    try:
        corpo_html = gerar_corpo_painel(caminho_arquivo)
        msg.attach(MIMEText(corpo_html, 'html'))
    except Exception as e:
        print(f"Erro ao processar dados do Excel: {e}")
        return

    filename = os.path.basename(caminho_arquivo)
    try:
        with open(caminho_arquivo, "rb") as anexo:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(anexo.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f'attachment; filename={filename}')
            msg.attach(part)
    except Exception as e:
        print(f"Erro ao anexar arquivo: {e}")

    try:
        server = smtplib.SMTP(SMTP_SERVIDOR, SMTP_PORTA)
        server.starttls()
        server.login(EMAIL_REMETENTE, SENHA_EMAIL)
        server.sendmail(EMAIL_REMETENTE, todos_destinatarios, msg.as_string())
        server.quit()
        print(f"Relatório enviado com sucesso para {EMAIL_DESTINATARIO} e em cópia para {EMAILS_COPIA}!")
    except Exception as e:
        print(f"Falha ao enviar e-mail: {e}")

if __name__ == "__main__":
    enviar_email_com_anexo(xlsx_final)
