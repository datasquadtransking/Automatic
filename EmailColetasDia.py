import pandas as pd
from datetime import datetime
import os
import re
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

# ====== CONFIGURAÇÕES ======
ARQUIVO_ORIGINAL = r"C:\Users\flavi\OneDrive\001\BI\planilha_coleta.xlsx"

EMAIL_REMETENTE = "automac@transking.com.br"
SENHA_EMAIL = "P3dr0Tk2025"  # Senha ou senha de app
EMAIL_DESTINATARIO = "weslley.transking@gmail.com"
SMTP_SERVIDOR = "smtp.task.com.br"
SMTP_PORTA = 587


# ====== ETAPA 1 - TRATAMENTO ======
df = pd.read_excel(ARQUIVO_ORIGINAL)
df['DATA_INCLUSAO'] = pd.to_datetime(df['DATA_INCLUSAO'], errors='coerce')


# ====== ETAPA 1 - LEITURA E NORMALIZAÇÃO ======
df = pd.read_excel(ARQUIVO_ORIGINAL)
df.columns = df.columns.str.strip().str.upper()
df['DATA_INCLUSAO'] = pd.to_datetime(df['DATA_INCLUSAO'], dayfirst=True, errors='coerce')

remetentes_indesejados = [
    "AMAZON LOGISTICA DO BRASIL LTD",
    "JADLOG LOGISTICA S.A",
    "MERCADO ENVIOS SERVICOS DE LOG",
    "GRADIN E ALMEIDA LOGISTICA",
    "SHPX LOGISTICA LTDA."
]

abreviacoes_clientes = {
    "RECKITT BENCKISER BRASIL COM D": "RECKITT",
    "UNILEVER BRASIL IND. LTDA": "UNILEVER",
    "INGLEZA INDUSTRIA DE PRODUTOS": "INGLEZA",
    "PEGAKI TECNOLOGIA DE ENTREGAS": "PEGAKI",
    "LG ELECTRONICS DO BRASIL LTDA": "LG",
    "SUPPLOG ARMAZENS GERAIS E ENTR": "SUPPLOG",
    "MARLUVAS EQUIPAMENTOS DE SEGUR": "MARLUVAS",
    "ANTILHAS EMBALAGENS FLEXIVEIS": "ANTILHAS ",
    "ETIC DISTRIBUIDORA HCP LTDA": "ETIC",
    "PROCTER EAMP GAMBLE INDUSTRIAL": "P&G",
    "SANCHEZ CANO LTDA": "FINI",
    "HEALTH LOGISTICA HOSPITALAR SA": "HEALTH"
}

# ====== ETAPA 2 - FILTRAGEM ======
padrao = '|'.join([re.escape(r) for r in remetentes_indesejados])
df_filtrado = df[~df['REMETENTE'].str.contains(padrao, case=False, na=False)]
df_filtrado = df_filtrado[df_filtrado['SITUACAO'].str.upper() != 'CANCELADA']
df_filtrado = df_filtrado[df_filtrado['ULTIMA_OCORRENCIA'].str.strip() != '8 - CARGA DECLINADA PELO CLIENTE']
df_filtrado['PAGADOR'] = df_filtrado['PAGADOR'].map(abreviacoes_clientes).fillna(df_filtrado['PAGADOR'])

# ====== ETAPA 3 - SELEÇÃO E RENOMEAÇÃO ======
colunas_desejadas = [
    "NUMERO_COLETA", "COTACAO", "DATA_INCLUSAO", "HORA_INCLUSAO", "USUARIO", "PAGADOR",
    "SITUACAO", "DATA_ULTIMA_OCORRENCIA", "HORA_ULTIMA_OCORRENCIA", "USUARIO_ULTIMA_OCORRENCIA",
    "ULTIMA_OCORRENCIA", "OBSERVACAO_2"
]

df_final = df_filtrado[colunas_desejadas].rename(columns={
    "NUMERO_COLETA": "Coleta",
    "DATA_INCLUSAO": "DataInc",
    "HORA_INCLUSAO": "HoraInc",
    "USUARIO": "Login",
    "DATA_ULTIMA_OCORRENCIA": "DataUlt",
    "HORA_ULTIMA_OCORRENCIA": "HoraUlt",
    "ULTIMA_OCORRENCIA": "Ult.Oco",
    "USUARIO_ULTIMA_OCORRENCIA": "LoginUlt"
}).copy()

# ====== ETAPA 4 - EXTRAÇÃO DATAS/HORAS COLETA E ENTREGA ======
coleta_entrega = df_final['OBSERVACAO_2'].str.extract(
    r'(\d{2}/\d{2}/\d{4}) (\d{2}:\d{2}) - (\d{2}/\d{2}/\d{4}) (\d{2}:\d{2})'
)
coleta_entrega.columns = ['DataCol', 'HoraCol', 'DataEnt', 'HoraEnt']
df_final = pd.concat([df_final.drop(columns=['OBSERVACAO_2']), coleta_entrega], axis=1)

# ====== ETAPA 5 - NOVAS COLUNAS BASEADAS NO CÓDIGO ======
mapeamento = {
    '97': 'Ag.Contratação',
    '98': 'Ag.Contratação',
    '80': 'Ag.Contratação',
    '81': 'Em Gr',
    '82': 'Deslocando para Coleta',
    '34': 'Acompanhar e dar o 10',
    '1': 'Ag Emissão',
    '10': 'Acompanhar Entrega',
    '35': 'Ag.Coleta',
    '50': 'Acompanhar Coleta',
    '17': 'Ag.Descarga',
    '3': 'Ag.AE',
    '30': 'Entregue',
    '11': 'Finalizada',
}

df_final['CodigoOco'] = df_final['Ult.Oco'].str.split('-').str[0].str.strip()
df_final['DescricaoOco'] = df_final['CodigoOco'].map(mapeamento)

# ====== FORMATAR TODAS AS DATAS PARA dd/mm/aa ======
for col in ['DataInc', 'DataUlt', 'DataCol', 'DataEnt']:
    df_final[col] = pd.to_datetime(df_final[col], dayfirst=True, errors='coerce').dt.strftime('%d/%m/%y')

# ====== FUNÇÃO PARA CENTRALIZAR CABEÇALHO HTML ======
def centralizar_cabecalho(df):
    html = df.to_html(index=False, escape=False)
    html = html.replace('<th>', '<th style="text-align:center">')
    return html

# ====== ETAPA 6 - FILTRAGEM PARA DATA DE HOJE ======
data_hoje = datetime.today().strftime('%d/%m/%y')
df_final = df_final[
    (df_final['DataCol'] == data_hoje) |
    (df_final['DataEnt'] == data_hoje)
]

df_final['TIPO'] = df_final.apply(
    lambda x: 'COLETA' if x['DataCol'] == data_hoje else 'ENTREGA',
    axis=1
)

df_final['HORA_BASE'] = df_final.apply(
    lambda x: x['HoraCol'] if x['TIPO'] == 'COLETA' else x['HoraEnt'],
    axis=1
)
df_final = df_final.sort_values(
    by='HORA_BASE',
    key=lambda col: pd.to_datetime(col, format='%H:%M', errors='coerce')
).drop(columns=['HORA_BASE'])

# ====== ETAPA 7 - CRIAÇÃO DAS LISTAS PARA O EMAIL ======
coletas = df_final[df_final['TIPO'] == 'COLETA']
entregas = df_final[df_final['TIPO'] == 'ENTREGA']

codigos_coleta_acomp = ['97', '98', '80', '81', '82', '50']
codigos_entrega_acomp = ['34', '10', '97', '98', '80', '81', '82', '50']

coletas_acomp = coletas[coletas['CodigoOco'].isin(codigos_coleta_acomp)]
entregas_acomp = entregas[entregas['CodigoOco'].isin(codigos_entrega_acomp)]

# ====== REMOVER COLUNA 'CodigoOco' DOS DATAFRAMES DO EMAIL ======
coletas = coletas.drop(columns=['CodigoOco'])
entregas = entregas.drop(columns=['CodigoOco'])
coletas_acomp = coletas_acomp.drop(columns=['CodigoOco'])
entregas_acomp = entregas_acomp.drop(columns=['CodigoOco'])

# ====== GERAÇÃO DO HTML ======
html_coletas = centralizar_cabecalho(coletas)
html_entregas = centralizar_cabecalho(entregas)
html_coletas_acomp = centralizar_cabecalho(coletas_acomp)
html_entregas_acomp = centralizar_cabecalho(entregas_acomp)

# ====== REMOVER 'CodigoOco' DO df_final ANTES DE SALVAR ======
df_final = df_final.drop(columns=['CodigoOco'])

# ====== ETAPA 8 - SAUDAÇÃO ======
hora_atual = datetime.now().hour
if 12 <= hora_atual < 18:
    saudacao = "Boa tarde"
elif 18 <= hora_atual <= 23:
    saudacao = "Boa noite"
else:
    saudacao = "Bom dia"


# ====== ETAPA 9 - CORPO DO E-MAIL ======
mensagem_html = f"""
<p>{saudacao}.</p>
<p>Hoje temos <b>{len(coletas)}</b> coletas e <b>{len(entregas)}</b> entregas. Total de <b>{len(coletas) + len(entregas)}</b> carregamentos.</p>

<h3>📦 Coletas ({len(coletas)}):</h3>
{html_coletas}
<br>

<h3>📦 Entregas ({len(entregas)}):</h3>
{html_entregas}
<br>

<h3>🚨 Coletas que precisam acompanhamento ({len(coletas_acomp)}):</h3>
{html_coletas_acomp}
<br>

<h3>🚨 Entregas que precisam acompanhamento ({len(entregas_acomp)}):</h3>
{html_entregas_acomp}
"""

# ====== ETAPA 10 - SALVA PLANILHA ======
arquivo_saida = os.path.join(os.path.expanduser("~"), "Downloads", "coletas_entregas.xlsx")
df_final.to_excel(arquivo_saida, index=False)

# ====== ETAPA 11 - ENVIO DE E-MAIL ======
msg = MIMEMultipart()
msg["Subject"] = f"Coletas e Entregas - {data_hoje}"
msg["From"] = EMAIL_REMETENTE
msg["To"] = EMAIL_DESTINATARIO
msg.attach(MIMEText(mensagem_html, "html"))

with open(arquivo_saida, "rb") as f:
    parte = MIMEBase("application", "octet-stream")
    parte.set_payload(f.read())
encoders.encode_base64(parte)
parte.add_header(
    "Content-Disposition",
    f"attachment; filename={os.path.basename(arquivo_saida)}"
)
msg.attach(parte)

with smtplib.SMTP(SMTP_SERVIDOR, SMTP_PORTA) as server:
    server.starttls()
    server.login(EMAIL_REMETENTE, SENHA_EMAIL)
    server.sendmail(EMAIL_REMETENTE, EMAIL_DESTINATARIO, msg.as_string())

print("✅ E-mail enviado com sucesso!")