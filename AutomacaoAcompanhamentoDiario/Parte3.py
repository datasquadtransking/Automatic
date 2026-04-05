

# import pandas as pd
# import smtplib
# import os
# from email.mime.multipart import MIMEMultipart
# from email.mime.text import MIMEText
# from email.mime.base import MIMEBase
# from email import encoders
# from datetime import datetime, timedelta

# # --- CONFIGURAÇÕES ---
# EMAIL_REMETENTE = "automac@transking.com.br"
# SENHA_EMAIL = "P3dr0Tk2025"
# EMAIL_DESTINATARIO = "weslley.transking@gmail.com"
# SMTP_SERVIDOR = "smtp.task.com.br"
# SMTP_PORTA = 587

# PASTA_PROJETO = r"C:\Users\flavi\Documents\GitHub\Automatic\AutomacaoAcompanhamentoDiario"
# ARQUIVO_TRATADO = os.path.join(PASTA_PROJETO, "Relatorio_Pronto_Para_Colar.xlsx")

# def estilizar_tabela(df, largura_total="100%", larguras_cols=None):
#     df_copy = df.copy()
#     for col in df_copy.columns:
#         if df_copy[col].dtype in ['float64', 'int64']:
#              df_copy[col] = df_copy[col].fillna(0).astype(int).astype(str)

#     html = f'<table style="border-collapse: collapse; width: {largura_total}; font-family: Segoe UI, Arial; font-size: 10px; margin-bottom: 10px; border: 1px solid #dee2e6;">'
#     html += '<thead style="background-color: #1f4e78; color: white;"><tr>'
#     for col in df_copy.columns:
#         largura = f' width="{larguras_cols[col]}"' if larguras_cols and col in larguras_cols else ''
#         html += f'<th{largura} style="padding: 3px; border: 1px solid #dee2e6; text-align: left;">{col}</th>'
#     html += '</tr></thead><tbody>'
    
#     for _, row in df_copy.iterrows():
#         html += '<tr>'
#         for val in row:
#             html += f'<td style="padding: 2px; border: 1px solid #dee2e6; text-align: left;">{val}</td>'
#         html += '</tr>'
#     html += '</tbody></table>'
#     return html

# def gerar_bloco_cliente(df, cliente, hoje, amanha):
#     # Filtra pelo cliente alvo
#     df_cli = df[df['CLIENTE'].str.contains(cliente, na=False, case=False)]
#     if df_cli.empty: return ""

#     colunas_det = ['ID', 'COLETA', 'COTACAO', 'ROTA', 'DOCK/AGENDAMENTO']
#     larguras_det = {'ID': '100px', 'COLETA': '50px', 'COTACAO': '50px', 'ROTA': '180px', 'DOCK/AGENDAMENTO': '100px'}
    
#     html = f"<div style='background-color: #1f4e78; color: white; padding: 4px; font-size: 11px; margin-top: 10px;'><b>🏢 {cliente}</b></div>"
    
#     # NOVA LÓGICA DE FILTRO BASEADA NAS DATAS DE COLETA E ENTREGA
#     # 1. Coletas Hoje (Sempre que a data de coleta for HOJE)
#     col_hoje = df_cli[df_cli['DT_COLETA_OK'] == hoje][colunas_det]
    
#     # 2. Entregas Hoje (Sempre que a data de entrega for HOJE)
#     ent_hoje = df_cli[df_cli['DT_ENTREGA_OK'] == hoje][colunas_det]
    
#     # 3. Entregas Amanhã (Sempre que a data de entrega for AMANHÃ)
#     ent_amanha = df_cli[df_cli['DT_ENTREGA_OK'] == amanha][colunas_det]

#     # 4. Coletas Amanhã (Sempre que a data de coleta for AMANHÃ)
#     col_amanha = df_cli[df_cli['DT_COLETA_OK'] == amanha][colunas_det]

    

#     for titulo, dados in [("🚚 Coletas Hoje", col_hoje), 
#                           ("📦 Entregas Hoje", ent_hoje), 
#                           ("🚚 Coletas Amanhã", col_amanha), 
#                           ("📅 Entregas Amanhã", ent_amanha)]:
#         if not dados.empty:
#             html += f"<p style='margin: 5px 0 2px 0; font-size: 10px;'><b>{titulo}: {len(dados)}</b></p>"
#             html += estilizar_tabela(dados, largura_total="100%", larguras_cols=larguras_det)
#     return html

# def enviar_email_operacao():
#     if not os.path.exists(ARQUIVO_TRATADO): return

#     df = pd.read_excel(ARQUIVO_TRATADO)
    
#     # Converte as colunas de data para comparação
#     df['DT_COLETA_OK'] = pd.to_datetime(df['PROGRAMAÇÃO. COLETA'], dayfirst=True).dt.date
#     df['DT_ENTREGA_OK'] = pd.to_datetime(df['PROGRAMAÇÃO. ENTREGA'], dayfirst=True).dt.date
    
#     hoje = datetime.now().date()
#     amanha = hoje + timedelta(days=1)

#     # 1. AJUSTE DA DINÂMICA (Baseada na nova lógica de datas)
#     # Identifica o que é Coleta ou Entrega hoje/amanhã para o resumo
#     df['RESUMO_TIPO'] = 'OUTROS'
#     df.loc[df['DT_COLETA_OK'] == hoje, 'RESUMO_TIPO'] = 'COLETA'
#     df.loc[df['DT_ENTREGA_OK'].isin([hoje, amanha]), 'RESUMO_TIPO'] = 'ENTREGA'
    
#     df_res = df[df['RESUMO_TIPO'] != 'OUTROS']
#     pivot = df_res.pivot_table(index='CLIENTE', columns='RESUMO_TIPO', values='ID', aggfunc='count', fill_value=0).reset_index()
#     pivot.columns.name = None
    
#     for col in ['COLETA', 'ENTREGA']:
#         if col not in pivot.columns: pivot[col] = 0
    
#     pivot['TOTAL'] = pivot['COLETA'] + pivot['ENTREGA']
#     total_geral = pd.DataFrame({'CLIENTE': ['TOTAL GERAL'], 'COLETA': [pivot['COLETA'].sum()], 
#                                 'ENTREGA': [pivot['ENTREGA'].sum()], 'TOTAL': [pivot['TOTAL'].sum()]})
#     pivot = pd.concat([pivot, total_geral], ignore_index=True)
    
#     html_dinamica = estilizar_tabela(pivot, largura_total="350px")

#     # 2. DETALHAMENTO LADO A LADO
#     detalhe_unilever = gerar_bloco_cliente(df, 'UNILEVER', hoje, amanha)
#     detalhe_reckitt = gerar_bloco_cliente(df, 'RECKITT', hoje, amanha)

#     html_body = f"""
#     <html>
#     <body style="margin: 0; padding: 10px; font-family: Segoe UI, Arial;">
#         <h3 style="color: #1f4e78; margin-bottom: 5px;">📊 Resumo Operacional - {hoje.strftime('%d/%m/%Y')}</h3>
#         {html_dinamica}
#         <hr style="border: 0; border-top: 1px solid #eee; margin: 15px 0;">
#         <table width="100%" border="0" cellspacing="0" cellpadding="0">
#             <tr>
#                 <td width="49%" valign="top" style="padding-right: 10px;">{detalhe_unilever}</td>
#                 <td width="2%" style="border-left: 1px solid #eee;">&nbsp;</td>
#                 <td width="49%" valign="top" style="padding-left: 10px;">{detalhe_reckitt}</td>
#             </tr>
#         </table>
#     </body>
#     </html>
#     """

#     msg = MIMEMultipart()
#     msg['From'], msg['To'], msg['Subject'] = EMAIL_REMETENTE, EMAIL_DESTINATARIO, f"Acompanhamento Operação - {hoje.strftime('%d/%m/%Y')}"
#     msg.attach(MIMEText(html_body, 'html'))

#     try:
#         with open(ARQUIVO_TRATADO, "rb") as f:
#             part = MIMEBase('application', 'octet-stream'); part.set_payload(f.read())
#             encoders.encode_base64(part); part.add_header('Content-Disposition', 'attachment; filename=Relatorio.xlsx'); msg.attach(part)
#         server = smtplib.SMTP(SMTP_SERVIDOR, SMTP_PORTA); server.starttls()
#         server.login(EMAIL_REMETENTE, SENHA_EMAIL); server.send_message(msg); server.quit()
#         print("✅ E-mail enviado com a nova lógica de filtro (por data)!")
#     except Exception as e: print(f"❌ Erro: {e}")

# if __name__ == "__main__":
#     enviar_email_operacao()

import pandas as pd
import smtplib
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime, timedelta

# --- CONFIGURAÇÕES ---
EMAIL_REMETENTE = "automac@transking.com.br"
SENHA_EMAIL = "P3dr0Tk2025"
EMAIL_DESTINATARIO = "weslley.transking@gmail.com"
SMTP_SERVIDOR = "smtp.task.com.br"
SMTP_PORTA = 587

PASTA_PROJETO = r"C:\Users\flavi\Documents\GitHub\Automatic\AutomacaoAcompanhamentoDiario"
ARQUIVO_TRATADO = os.path.join(PASTA_PROJETO, "Relatorio_Pronto_Para_Colar.xlsx")
ARQUIVO_UNILEVER = os.path.join(PASTA_PROJETO, "Base_Chegada_Unilever.xlsx")

def estilizar_tabela(df, largura_total="100%", larguras_cols=None, cor_cabecalho="#1f4e78"):
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

def gerar_bloco_chegada_unilever():
    """Lê a base específica da Unilever e gera o HTML em AZUL"""
    if not os.path.exists(ARQUIVO_UNILEVER):
        return "<p style='font-size: 10px; color: red;'>Arquivo Base_Chegada_Unilever não encontrado.</p>"
    
    df_uni = pd.read_excel(ARQUIVO_UNILEVER)
    colunas_uni = ['ID', 'TIPO', 'PLACA', 'HORA_CHEGADA', 'CHECKIN_CD', 'DOCK_DATA', 'DME', 'CHEGADA_CLIENTE']
    df_uni = df_uni[[c for c in colunas_uni if c in df_uni.columns]]
    
    # Alterado de #385623 (Verde) para #1f4e78 (Azul)
    html = f"<div style='background-color: #1f4e78; color: white; padding: 4px; font-size: 11px; margin-top: 10px;'><b>🚛 Status Chegada/Docking UNILEVER (Portal)</b></div>"
    html += estilizar_tabela(df_uni, cor_cabecalho="#1f4e78")
    return html

def gerar_bloco_cliente(df, cliente, hoje, amanha):
    df_cli = df[df['CLIENTE'].str.contains(cliente, na=False, case=False)]
    if df_cli.empty: return ""

    colunas_det = ['ID', 'COLETA', 'COTACAO', 'ROTA', 'DOCK/AGENDAMENTO']
    larguras_det = {'ID': '100px', 'COLETA': '50px', 'COTACAO': '50px', 'ROTA': '180px', 'DOCK/AGENDAMENTO': '100px'}
    
    html = f"<div style='background-color: #1f4e78; color: white; padding: 4px; font-size: 11px; margin-top: 10px;'><b>🏢 {cliente}</b></div>"
    
    col_hoje = df_cli[df_cli['DT_COLETA_OK'] == hoje][colunas_det]
    ent_hoje = df_cli[df_cli['DT_ENTREGA_OK'] == hoje][colunas_det]
    ent_amanha = df_cli[df_cli['DT_ENTREGA_OK'] == amanha][colunas_det]
    col_amanha = df_cli[df_cli['DT_COLETA_OK'] == amanha][colunas_det]

    for titulo, dados in [("🚚 Coletas Hoje", col_hoje), 
                          ("📦 Entregas Hoje", ent_hoje), 
                          ("🚚 Coletas Amanhã", col_amanha), 
                          ("📅 Entregas Amanhã", ent_amanha)]:
        if not dados.empty:
            html += f"<p style='margin: 5px 0 2px 0; font-size: 10px;'><b>{titulo}: {len(dados)}</b></p>"
            html += estilizar_tabela(dados, largura_total="100%", larguras_cols=larguras_det)
    return html

def enviar_email_operacao():
    if not os.path.exists(ARQUIVO_TRATADO): return

    df = pd.read_excel(ARQUIVO_TRATADO)
    df['DT_COLETA_OK'] = pd.to_datetime(df['PROGRAMAÇÃO. COLETA'], dayfirst=True).dt.date
    df['DT_ENTREGA_OK'] = pd.to_datetime(df['PROGRAMAÇÃO. ENTREGA'], dayfirst=True).dt.date
    
    hoje = datetime.now().date()
    amanha = hoje + timedelta(days=1)

    # 1. Resumo Dinâmico
    df['RESUMO_TIPO'] = 'OUTROS'
    df.loc[df['DT_COLETA_OK'] == hoje, 'RESUMO_TIPO'] = 'COLETA'
    df.loc[df['DT_ENTREGA_OK'].isin([hoje, amanha]), 'RESUMO_TIPO'] = 'ENTREGA'
    df_res = df[df['RESUMO_TIPO'] != 'OUTROS']
    pivot = df_res.pivot_table(index='CLIENTE', columns='RESUMO_TIPO', values='ID', aggfunc='count', fill_value=0).reset_index()
    
    # Garante colunas para não dar erro na soma
    for c in ['COLETA', 'ENTREGA']:
        if c not in pivot.columns: pivot[c] = 0
        
    pivot['TOTAL'] = pivot['COLETA'] + pivot['ENTREGA']
    html_dinamica = estilizar_tabela(pivot, largura_total="350px")

    # 2. Detalhamento
    detalhe_unilever = gerar_bloco_cliente(df, 'UNILEVER', hoje, amanha)
    detalhe_reckitt = gerar_bloco_cliente(df, 'RECKITT', hoje, amanha)
    tabela_chegada_unilever = gerar_bloco_chegada_unilever()

    html_body = f"""
    <html>
    <body style="margin: 0; padding: 10px; font-family: Segoe UI, Arial;">
        <h3 style="color: #1f4e78; margin-bottom: 5px;">📊 Resumo Operacional - {hoje.strftime('%d/%m/%Y')}</h3>
        {html_dinamica}
        
        <hr style="border: 0; border-top: 1px solid #eee; margin: 15px 0;">
        
        <h3 style="color: #1f4e78; margin-bottom: 5px;">📌 Status de Chegada/Portais (Unilever)</h3>
        {tabela_chegada_unilever}

        <hr style="border: 0; border-top: 1px solid #eee; margin: 15px 0;">
        
        <h3 style="color: #c45911; margin-bottom: 10px;">🔍 Planejamento de Carga/Descarga</h3>
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
            <tr>
                <td width="49%" valign="top" style="padding-right: 10px;">{detalhe_unilever}</td>
                <td width="2%" style="border-left: 1px solid #eee;">&nbsp;</td>
                <td width="49%" valign="top" style="padding-left: 10px;">{detalhe_reckitt}</td>
            </tr>
        </table>
    </body>
    </html>
    """

    msg = MIMEMultipart()
    msg['From'], msg['To'], msg['Subject'] = EMAIL_REMETENTE, EMAIL_DESTINATARIO, f"Acompanhamento Operação - {hoje.strftime('%d/%m/%Y')}"
    msg.attach(MIMEText(html_body, 'html'))

    # Anexos
    for caminho, nome_exibicao in [(ARQUIVO_TRATADO, "Relatorio_Geral.xlsx"), (ARQUIVO_UNILEVER, "Base_Chegada_Unilever.xlsx")]:
        if os.path.exists(caminho):
            try:
                with open(caminho, "rb") as f:
                    part = MIMEBase('application', 'octet-stream')
                    part.set_payload(f.read())
                    encoders.encode_base64(part)
                    part.add_header('Content-Disposition', f'attachment; filename={nome_exibicao}')
                    msg.attach(part)
            except Exception as e: print(f"Erro ao anexar {nome_exibicao}: {e}")

    try:
        server = smtplib.SMTP(SMTP_SERVIDOR, SMTP_PORTA); server.starttls()
        server.login(EMAIL_REMETENTE, SENHA_EMAIL); server.send_message(msg); server.quit()
        print("✅ E-mail enviado com sucesso (Tudo Azul)!")
    except Exception as e: print(f"❌ Erro: {e}")

if __name__ == "__main__":
    enviar_email_operacao()