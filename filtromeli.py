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
SENHA_EMAIL = "P3dr0Tk2025"  # Senha ou senha de app
EMAIL_DESTINATARIO = "weslleyworksilva@yahoo.com"
SMTP_SERVIDOR = "smtp.task.com.br"
SMTP_PORTA = 587

# ====== ETAPA 1 - LEITURA E NORMALIZAÇÃO =====
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
    '22': '22Oportunidade Perdida',
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

print(df_final)