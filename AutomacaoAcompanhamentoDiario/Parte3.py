import pandas as pd
import smtplib
import os
import re
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime, timedelta

# --- CONFIGURAÇÕES EXISTENTES ---
EMAIL_REMETENTE = "automac@transking.com.br"
SENHA_EMAIL = "P3dr0Tk2025"
EMAIL_DESTINATARIO = "weslley.transking@gmail.com"
SMTP_SERVIDOR = "smtp.task.com.br"
SMTP_PORTA = 587

PASTA_PROJETO = r"C:\Users\flavi\Documents\GitHub\Automatic\AutomacaoAcompanhamentoDiario"
ARQUIVO_TRATADO = os.path.join(PASTA_PROJETO, "Relatorio_Pronto_Para_Colar.xlsx")
ARQUIVO_UNILEVER = os.path.join(PASTA_PROJETO, "Base_Chegada_Unilever.xlsx")

# --- NOVA CONFIGURAÇÃO (PLANILHA DE COLETA) ---
ARQUIVO_COLETA_ORIGINAL = r"C:\Users\flavi\OneDrive\001\BI\planilha_coleta.xlsx"

def estilizar_tabela(df, largura_total="100%", larguras_cols=None, cor_cabecalho="#1f4e78"):
    if df.empty: return ""
    df_copy = df.copy()
    for col in df_copy.columns:
        if df_copy[col].dtype in ['float64', 'int64']:
             df_copy[col] = df_copy[col].fillna(0).astype(int).astype(str)

    html = f'<table style="border-collapse: collapse; width: {largura_total}; font-family: Segoe UI, Arial; font-size: 10px; margin-bottom: 10px; border: 1px solid #dee2e6;">'
    html += f'<thead style="background-color: {cor_cabecalho}; color: white;"><tr>'
    for col in df_copy.columns:
        largura = f' width="{larguras_cols[col]}"' if larguras_cols and col in larguras_cols else ''
        html += f'<th{largura} style="padding: 3px; border: 1px solid #dee2e6; text-align: left;">{col}</th>'
    html += '</tr></thead><tbody>'
    
    for _, row in df_copy.iterrows():
        html += '<tr>'
        for val in row:
            html += f'<td style="padding: 2px; border: 1px solid #dee2e6; text-align: left;">{val}</td>'
        html += '</tr>'
    html += '</tbody></table>'
    return html

def tratar_planilha_coleta():
    """Processa a lógica da planilha de coleta (Etapas 1-7 do seu código 2)"""
    if not os.path.exists(ARQUIVO_COLETA_ORIGINAL): return "", "", "", "", pd.DataFrame()
    
    df = pd.read_excel(ARQUIVO_COLETA_ORIGINAL)
    df.columns = df.columns.str.strip().str.upper()
    
    # CORREÇÃO DO AVISO (UserWarning): Definindo o formato explicitamente
    df['DATA_INCLUSAO'] = pd.to_datetime(df['DATA_INCLUSAO'], format='%d/%m/%Y', errors='coerce')

    remetentes_indesejados = ["AMAZON LOGISTICA DO BRASIL LTD", "JADLOG LOGISTICA S.A", "MERCADO ENVIOS SERVICOS DE LOG", "SHPX LOGISTICA LTDA."]
    abreviacoes_clientes = {
        "RECKITT BENCKISER BRASIL COM D": "RECKITT", "UNILEVER BRASIL IND. LTDA": "UNILEVER",
        "INGLEZA INDUSTRIA DE PRODUTOS": "INGLEZA", "PEGAKI TECNOLOGIA DE ENTREGAS": "PEGAKI",
        "LG ELECTRONICS DO BRASIL LTDA": "LG", "SUPPLOG ARMAZENS GERAIS E ENTR": "SUPPLOG",
        "MARLUVAS EQUIPAMENTOS DE SEGUR": "MARLUVAS", "ANTILHAS EMBALAGENS FLEXIVEIS": "ANTILHAS ",
        "ETIC DISTRIBUIDORA HCP LTDA": "ETIC", "PROCTER EAMP GAMBLE INDUSTRIAL": "P&G",
        "SANCHEZ CANO LTDA": "FINI", "HEALTH LOGISTICA HOSPITALAR SA": "HEALTH"
    }

    padrao = '|'.join([re.escape(r) for r in remetentes_indesejados])
    df_f = df[~df['REMETENTE'].str.contains(padrao, case=False, na=False)].copy()
    df_f = df_f[df_f['SITUACAO'].str.upper() != 'CANCELADA']
    df_f = df_f[df_f['ULTIMA_OCORRENCIA'].str.strip() != '8 - CARGA DECLINADA PELO CLIENTE']
    df_f['PAGADOR'] = df_f['PAGADOR'].map(abreviacoes_clientes).fillna(df_f['PAGADOR'])

    colunas_desejadas = ["NUMERO_COLETA", "COTACAO", "DATA_INCLUSAO", "HORA_INCLUSAO", "USUARIO", "PAGADOR", "SITUACAO", "DATA_ULTIMA_OCORRENCIA", "HORA_ULTIMA_OCORRENCIA", "USUARIO_ULTIMA_OCORRENCIA", "ULTIMA_OCORRENCIA", "OBSERVACAO_2"]
    df_final = df_f[colunas_desejadas].rename(columns={"NUMERO_COLETA": "Coleta", "DATA_INCLUSAO": "DataInc", "HORA_INCLUSAO": "HoraInc", "USUARIO": "Login", "DATA_ULTIMA_OCORRENCIA": "DataUlt", "HORA_ULTIMA_OCORRENCIA": "HoraUlt", "ULTIMA_OCORRENCIA": "Ult.Oco", "USUARIO_ULTIMA_OCORRENCIA": "LoginUlt"}).copy()

    coleta_entrega = df_final['OBSERVACAO_2'].str.extract(r'(\d{2}/\d{2}/\d{4}) (\d{2}:\d{2}) - (\d{2}/\d{2}/\d{4}) (\d{2}:\d{2})')
    coleta_entrega.columns = ['DataCol', 'HoraCol', 'DataEnt', 'HoraEnt']
    df_final = pd.concat([df_final.drop(columns=['OBSERVACAO_2']), coleta_entrega], axis=1)

    mapeamento = {'97': 'Ag.Contratação', '98': 'Ag.Contratação', '80': 'Ag.Contratação', '81': 'Em Gr', '82': 'Deslocando para Coleta', '34': 'Acompanhar e dar o 10', '1': 'Ag Emissão', '10': 'Acompanhar Entrega', '35': 'Ag.Coleta', '50': 'Acompanhar Coleta', '17': 'Ag.Descarga', '3': 'Ag.AE', '30': 'Entregue', '11': 'Finalizada'}
    df_final['CodigoOco'] = df_final['Ult.Oco'].str.split('-').str[0].str.strip()
    df_final['DescricaoOco'] = df_final['CodigoOco'].map(mapeamento)

    # CORREÇÃO DO AVISO (UserWarning): Forçando o formato ao converter strings de datas
    for col in ['DataInc', 'DataUlt', 'DataCol', 'DataEnt']:
        df_final[col] = pd.to_datetime(df_final[col], format='%d/%m/%Y', errors='coerce').dt.strftime('%d/%m/%y')

    data_hoje = datetime.today().strftime('%d/%m/%y')
    df_hoje = df_final[(df_final['DataCol'] == data_hoje) | (df_final['DataEnt'] == data_hoje)].copy()
    df_hoje['TIPO'] = df_hoje.apply(lambda x: 'COLETA' if x['DataCol'] == data_hoje else 'ENTREGA', axis=1)
    
    coletas = df_hoje[df_hoje['TIPO'] == 'COLETA'].drop(columns=['CodigoOco'])
    entregas = df_hoje[df_hoje['TIPO'] == 'ENTREGA'].drop(columns=['CodigoOco'])
    
    # Acompanhamentos
    cod_col_ac = ['97', '98', '80', '81', '82', '50']
    cod_ent_ac = ['34', '10', '97', '98', '80', '81', '82', '50']
    coletas_ac = df_hoje[(df_hoje['TIPO'] == 'COLETA') & (df_hoje['CodigoOco'].isin(cod_col_ac))].drop(columns=['CodigoOco'])
    entregas_ac = df_hoje[(df_hoje['TIPO'] == 'ENTREGA') & (df_hoje['CodigoOco'].isin(cod_ent_ac))].drop(columns=['CodigoOco'])

    return estilizar_tabela(coletas), estilizar_tabela(entregas), estilizar_tabela(coletas_ac), estilizar_tabela(entregas_ac), df_hoje.drop(columns=['CodigoOco'])

def gerar_bloco_chegada_unilever():
    if not os.path.exists(ARQUIVO_UNILEVER): return ""
    df_uni = pd.read_excel(ARQUIVO_UNILEVER)
    colunas_uni = ['ID', 'TIPO', 'PLACA', 'HORA_CHEGADA', 'CHECKIN_CD', 'DOCK_DATA', 'DME', 'CHEGADA_CLIENTE']
    df_uni = df_uni[[c for c in colunas_uni if c in df_uni.columns]]
    html = f"<div style='background-color: #1f4e78; color: white; padding: 4px; font-size: 11px; margin-top: 10px;'><b>🚛 Status Chegada/Docking UNILEVER (Portal)</b></div>"
    html += estilizar_tabela(df_uni)
    return html

def gerar_bloco_cliente(df, cliente, hoje, amanha):
    df_cli = df[df['CLIENTE'].str.contains(cliente, na=False, case=False)]
    if df_cli.empty: return ""
    
    # CORREÇÃO DO ERRO CRÍTICO: Garante que o script não quebre se a coluna estiver ausente
    colunas_det = ['ID', 'COLETA', 'COTACAO', 'ROTA', 'DOCK/AGENDAMENTO']
    colunas_existentes = [c for c in colunas_det if c in df_cli.columns]
    
    html = f"<div style='background-color: #1f4e78; color: white; padding: 4px; font-size: 11px; margin-top: 10px;'><b>🏢 {cliente}</b></div>"
    
    # Filtros por data (hoje/amanhã)
    for tit, d_f in [("🚚 Coletas Hoje", df_cli[df_cli['DT_COLETA_OK'] == hoje]), ("📦 Entregas Hoje", df_cli[df_cli['DT_ENTREGA_OK'] == hoje]), ("🚚 Coletas Amanhã", df_cli[df_cli['DT_COLETA_OK'] == amanha]), ("📅 Entregas Amanhã", df_cli[df_cli['DT_ENTREGA_OK'] == amanha])]:
        if not d_f.empty:
            html += f"<p style='margin: 5px 0 2px 0; font-size: 10px;'><b>{tit}: {len(d_f)}</b></p>"
            html += estilizar_tabela(d_f[colunas_existentes]) # Usa apenas as colunas que realmente existem
    return html

def enviar_email_operacao():
    # 1. Processamento Base Geral
    if not os.path.exists(ARQUIVO_TRATADO): return
    df = pd.read_excel(ARQUIVO_TRATADO)
    
    # CORREÇÃO EXTRA: Padroniza as colunas da planilha geral para evitar erros de maiúsculas/minúsculas
    df.columns = df.columns.str.strip().str.upper()
    
    # Ajustando os nomes chamados na sequência por causa do .upper()
    df['DT_COLETA_OK'] = pd.to_datetime(df['PROG. COLETA'], format='%d/%m/%Y', errors='coerce').dt.date
    df['DT_ENTREGA_OK'] = pd.to_datetime(df['PROG. ENTREGA'], format='%d/%m/%Y', errors='coerce').dt.date
    hoje = datetime.now().date()
    amanha = hoje + timedelta(days=1)

    # Resumo Geral
    df['RESUMO_TIPO'] = 'OUTROS'
    df.loc[df['DT_COLETA_OK'] == hoje, 'RESUMO_TIPO'] = 'COLETA'
    df.loc[df['DT_ENTREGA_OK'].isin([hoje, amanha]), 'RESUMO_TIPO'] = 'ENTREGA'
    pivot = df[df['RESUMO_TIPO'] != 'OUTROS'].pivot_table(index='CLIENTE', columns='RESUMO_TIPO', values='ID', aggfunc='count', fill_value=0).reset_index()
    for c in ['COLETA', 'ENTREGA']: 
        if c not in pivot.columns: pivot[c] = 0
    pivot['TOTAL'] = pivot['COLETA'] + pivot['ENTREGA']
    
    # 2. Processamento Base Coletas/Entregas (OneDrive)
    h_col, h_ent, h_col_ac, h_ent_ac, df_coleta_excel = tratar_planilha_coleta()

    # 3. Montagem do HTML
    saudacao = "Boa tarde" if 12 <= datetime.now().hour < 18 else "Boa noite" if 18 <= datetime.now().hour <= 23 else "Bom dia"
    
    html_body = f"""
    <html>
    <body style="margin: 0; padding: 10px; font-family: Segoe UI, Arial;">
        <p>{saudacao}, Weslley.</p>
        <h3 style="color: #1f4e78;">📊 Resumo Operacional - {hoje.strftime('%d/%m/%Y')}</h3>
        {estilizar_tabela(pivot, largura_total="350px")}
        
        <hr style="border: 0; border-top: 1px solid #eee; margin: 15px 0;">
        <h3 style="color: #1f4e78;">🚚 Coletas e Entregas (Base Geral)</h3>
        <p style='font-size: 10px;'>Total de hoje: <b>{len(df_coleta_excel)}</b> carregamentos.</p>
        <h4>📦 Coletas:</h4>{h_col}
        <h4>📦 Entregas:</h4>{h_ent}
        <h4 style="color: #d44;">🚨 Acompanhamento Crítico (Coletas):</h4>{h_col_ac}
        <h4 style="color: #d44;">🚨 Acompanhamento Crítico (Entregas):</h4>{h_ent_ac}

        <hr style="border: 0; border-top: 1px solid #eee; margin: 15px 0;">
        <h3 style="color: #1f4e78;">📌 Status Unilever Portal</h3>
        {gerar_bloco_chegada_unilever()}

        <hr style="border: 0; border-top: 1px solid #eee; margin: 15px 0;">
        <h3 style="color: #c45911;">🔍 Detalhamento por Cliente</h3>
        <table width="100%" border="0">
            <tr>
                <td width="49%" valign="top">{gerar_bloco_cliente(df, 'UNILEVER', hoje, amanha)}</td>
                <td width="2%">&nbsp;</td>
                <td width="49%" valign="top">{gerar_bloco_cliente(df, 'RECKITT', hoje, amanha)}</td>
            </tr>
        </table>
    </body>
    </html>
    """

    # 4. Envio e Anexos
    msg = MIMEMultipart()
    msg['Subject'] = f"Acompanhamento Operação - {hoje.strftime('%d/%m/%Y')}"
    msg['From'], msg['To'] = EMAIL_REMETENTE, EMAIL_DESTINATARIO
    msg.attach(MIMEText(html_body, 'html'))

    # Salva temporariamente a nova planilha tratada para anexar
    path_tmp_coleta = os.path.join(os.path.expanduser("~"), "Downloads", "coletas_entregas_hoje.xlsx")
    df_coleta_excel.to_excel(path_tmp_coleta, index=False)

    for cam, nome in [(ARQUIVO_TRATADO, "Relatorio_Geral.xlsx"), (ARQUIVO_UNILEVER, "Base_Chegada_Unilever.xlsx"), (path_tmp_coleta, "Coletas_Entregas_Tratado.xlsx")]:
        if os.path.exists(cam):
            with open(cam, "rb") as f:
                part = MIMEBase('application', 'octet-stream'); part.set_payload(f.read())
                encoders.encode_base64(part); part.add_header('Content-Disposition', f'attachment; filename={nome}'); msg.attach(part)

    try:
        server = smtplib.SMTP(SMTP_SERVIDOR, SMTP_PORTA); server.starttls()
        server.login(EMAIL_REMETENTE, SENHA_EMAIL); server.send_message(msg); server.quit()
        if os.path.exists(path_tmp_coleta): os.remove(path_tmp_coleta)
        print("✅ E-mail consolidado enviado com sucesso!")
    except Exception as e: print(f"❌ Erro: {e}")

if __name__ == "__main__":
    enviar_email_operacao()