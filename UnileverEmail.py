import pandas as pd
from datetime import datetime, timedelta
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

# ====== CONFIGURAÇÕES ======
ARQUIVO_ORIGINAL = r"C:\Users\flavi\OneDrive\001\BI\planilha_coleta.xlsx"

EMAIL_REMETENTE = "automac@transking.com.br"
SENHA_EMAIL = "#Tkisli2023@"  # Senha ou senha de app
EMAIL_DESTINATARIO = "weslleyworksilva@yahoo.com"
SMTP_SERVIDOR = "smtp.task.com.br"
SMTP_PORTA = 587

# ====== ETAPA 1 - LEITURA E NORMALIZAÇÃO ====== 
df = pd.read_excel(ARQUIVO_ORIGINAL)
df.columns = df.columns.str.strip().str.upper()
df['DATA_INCLUSAO'] = pd.to_datetime(df['DATA_INCLUSAO'], dayfirst=True, errors='coerce')

# ====== ETAPA 2 - FILTRAGEM ======
df_filtrado = df[df['SITUACAO'].str.upper() != 'CANCELADA']
df_filtrado = df_filtrado[df_filtrado['ULTIMA_OCORRENCIA'].str.strip() != '8 - CARGA DECLINADA PELO CLIENTE']

# >>> Só manter linhas com "UNILE" em REMETENTE ou PAGADOR
df_filtrado = df_filtrado[
    df_filtrado['REMETENTE'].str.contains("UNILE", case=False, na=False) |
    df_filtrado['PAGADOR'].str.contains("UNILE", case=False, na=False)
]

# ====== ETAPA 2.1 - EXTRAIR CÓDIGO DA OCORRÊNCIA ======
df_filtrado['CodigoOco'] = df_filtrado['ULTIMA_OCORRENCIA'].str.extract(r'^(\d+)\s*-\s*')[0]

# ====== ETAPA 2.2 - MAPEAMENTO STATUS ======
map_status = {
    '97': '97.Ag.Contratação',
    '98': '98.Ag.Contratação',
    '80': '80.Ag.Contratação',
    '81': '81.Em Gr',
    '82': '82.Deslocando para Coleta',
    '34': '34.Acompanhar e dar o 10',
    '1': '1 Ag.Emissão',
    '10': '10.Acompanhar Entrega',
    '35': '35.Ag.Coleta',
    '50': '50.Acompanhar Coleta',
    '17': '17.Ag.Descarga',
    '3': '3.Ag.AE',
    '30': '30.Entregue',
    '11': '11.Finalizada',
}

df_filtrado['Status'] = df_filtrado['CodigoOco'].map(map_status).fillna('Desconhecido')

# ====== ETAPA 3 - SELEÇÃO E RENOMEAÇÃO ======
colunas_desejadas = [
    "NUMERO_COLETA", "COTACAO", "DATA_INCLUSAO", "HORA_INCLUSAO", "USUARIO",
    "SITUACAO", "DATA_ULTIMA_OCORRENCIA", "HORA_ULTIMA_OCORRENCIA",
    "USUARIO_ULTIMA_OCORRENCIA", "ULTIMA_OCORRENCIA", "OBSERVACAO_2",
    "CIDADE_REMETENTE", "UF_REMETENTE", "CIDADE_DESTINO", "UF_DESTINO",
    "CodigoOco", "Status"
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

# ====== FORMATAR DATAS ======
for col in ['DataInc', 'DataUlt', 'DataCol', 'DataEnt']:
    df_final[col] = pd.to_datetime(df_final[col], dayfirst=True, errors='coerce')

# ====== ETAPA 4.1 - DATA E HORA DA ÚLTIMA OCORRÊNCIA ======
df_final['DataHoraUltimaOco'] = df_final.apply(
    lambda row: pd.to_datetime(str(row['DataUlt'].date()) + " " + str(row['HoraUlt']))
    if pd.notnull(row['DataUlt']) and pd.notnull(row['HoraUlt']) else pd.NaT,
    axis=1
)

# ====== ETAPA 4.2 - TEMPO DECORRIDO DA ÚLTIMA OCORRÊNCIA ======
df_final['TempoUltimaOco'] = df_final['DataHoraUltimaOco'].apply(
    lambda x: datetime.now() - x if pd.notnull(x) else pd.NaT
)
df_final['TempoUltimaOco'] = df_final['TempoUltimaOco'].apply(
    lambda td: f"{td.days}d {td.seconds//3600}h {(td.seconds//60)%60}m" if pd.notnull(td) else ""
)

# ====== DEFINIR PERÍODOS ======
hoje = datetime.today().date()
ontem = hoje - timedelta(days=1)
amanha = hoje + timedelta(days=1)

def classificar_data(row):
    if pd.notnull(row['DataCol']):
        if row['DataCol'].date() == ontem: return "Ontem"
        if row['DataCol'].date() == hoje: return "Hoje"
        if row['DataCol'].date() == amanha: return "Amanhã"
    if pd.notnull(row['DataEnt']):
        if row['DataEnt'].date() == ontem: return "Ontem"
        if row['DataEnt'].date() == hoje: return "Hoje"
        if row['DataEnt'].date() == amanha: return "Amanhã"
    return None

df_final["Periodo"] = df_final.apply(classificar_data, axis=1)
df_final = df_final[df_final["Periodo"].notna()]

# ====== DEFINIR TIPO (COLETA/ENTREGA) ======
def definir_tipo(row):
    if pd.notnull(row['DataCol']) and row['DataCol'].date() in [ontem, hoje, amanha]:
        return "COLETA"
    elif pd.notnull(row['DataEnt']) and row['DataEnt'].date() in [ontem, hoje, amanha]:
        return "ENTREGA"
    return "N/A"

df_final["TIPO"] = df_final.apply(definir_tipo, axis=1)

# ====== NOVA COLUNA ROTA ======
df_final["Rota"] = df_final["CIDADE_REMETENTE"].str.upper() + "/" + df_final["UF_REMETENTE"].str.upper() + " - " + \
                   df_final["CIDADE_DESTINO"].str.upper() + "/" + df_final["UF_DESTINO"].str.upper()

# ====== NOVA COLUNA DATA + HORA DO EVENTO ======
def definir_datahora_evento(row):
    if row["TIPO"] == "COLETA" and pd.notnull(row["DataCol"]):
        return pd.to_datetime(str(row["DataCol"].date()) + " " + str(row["HoraCol"]), errors="coerce")
    elif row["TIPO"] == "ENTREGA" and pd.notnull(row["DataEnt"]):
        return pd.to_datetime(str(row["DataEnt"].date()) + " " + str(row["HoraEnt"]), errors="coerce")
    return None

df_final["DataHoraEvento"] = df_final.apply(definir_datahora_evento, axis=1)

# ====== AJUSTE COLUNAS FINAIS ======
df_final = df_final[["Coleta", "COTACAO", "Rota", "DataHoraEvento", "TIPO", "Status", "DataHoraUltimaOco", "TempoUltimaOco", "Periodo"]]
df_final = df_final.sort_values(by="DataHoraEvento")

# ====== FORMATAR DataHoraEvento ======
df_final["DataHoraEvento"] = df_final["DataHoraEvento"].dt.strftime("%d/%m/%y %H:%M")
df_final["DataHoraUltimaOco"] = df_final["DataHoraUltimaOco"].dt.strftime("%d/%m/%y %H:%M")

# ====== FUNÇÃO PARA TABELA HTML ======
def centralizar_cabecalho(df):
    html = df.to_html(index=False, escape=False)
    html = html.replace('<th>', '<th style="text-align:center">')
    return html

# ====== RESUMO NUMÉRICO ======
summary_text = ""
for periodo in ["Ontem", "Hoje", "Amanhã"]:
    subset = df_final[df_final["Periodo"] == periodo]
    if not subset.empty:
        coletas = (subset["TIPO"] == "COLETA").sum()
        entregas = (subset["TIPO"] == "ENTREGA").sum()
        total = len(subset)
        summary_text += f"<p>Para {periodo.lower()}: {coletas} coletas e {entregas} entregas — total de {total} atendimentos.</p>"

# ====== TABELAS DETALHADAS ======
tables_html = ""
for periodo in ["Ontem", "Hoje", "Amanhã"]:
    subset = df_final[df_final["Periodo"] == periodo].drop(columns=["Periodo"])
    if not subset.empty:
        tables_html += f"<h3>📅 {periodo}</h3>"
        tables_html += centralizar_cabecalho(subset)

# ====== SAUDAÇÃO ======
hora_atual = datetime.now().hour
if 12 <= hora_atual < 18:
    saudacao = "Boa tarde"
elif 18 <= hora_atual <= 23:
    saudacao = "Boa noite"
else:
    saudacao = "Bom dia"

# ====== CORPO DO E-MAIL ======
mensagem_html = f"""
<p>{saudacao}.</p>
<p>Segue o resumo das operações da UNILEVER separadas por período:</p>
{summary_text}
<p><b>Abaixo o detalhamento:</b></p>
{tables_html}
"""

# ====== SALVA PLANILHA ======
arquivo_saida = os.path.join(os.path.expanduser("~"), "Downloads", "coletas_entregas.xlsx")
df_final.to_excel(arquivo_saida, index=False)

# ====== ENVIO DE E-MAIL ======
msg = MIMEMultipart()
msg["Subject"] = f"Resumo Operacional UNILEVER - {hoje.strftime('%d/%m/%Y')}"
msg["From"] = EMAIL_REMETENTE
msg["To"] = EMAIL_DESTINATARIO
msg.attach(MIMEText(mensagem_html, "html"))

with open(arquivo_saida, "rb") as f:
    parte = MIMEBase("application", "octet-stream")
    parte.set_payload(f.read())
encoders.encode_base64(parte)
parte.add_header("Content-Disposition", f"attachment; filename={os.path.basename(arquivo_saida)}")
msg.attach(parte)

with smtplib.SMTP(SMTP_SERVIDOR, SMTP_PORTA) as server:
    server.starttls()
    server.login(EMAIL_REMETENTE, SENHA_EMAIL)
    server.sendmail(EMAIL_REMETENTE, EMAIL_DESTINATARIO, msg.as_string())

print("✅ E-mail enviado com sucesso!")
