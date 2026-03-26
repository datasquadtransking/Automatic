# # import pandas as pd
# # import smtplib
# # import os
# # from email.mime.multipart import MIMEMultipart
# # from email.mime.text import MIMEText
# # from email.mime.base import MIMEBase
# # from email import encoders
# # from datetime import datetime, timedelta

# # # --- CONFIGURAÇÕES DE ACESSO ---
# # EMAIL_REMETENTE = "automac@transking.com.br"
# # SENHA_EMAIL = "P3dr0Tk2025"
# # EMAIL_DESTINATARIO = "weslley.transking@gmail.com"
# # SMTP_SERVIDOR = "smtp.task.com.br"
# # SMTP_PORTA = 587

# # # --- CAMINHOS ---
# # PASTA_PROJETO = r"C:\Users\flavi\Documents\GitHub\Automatic\AutomacaoAcompanhamentoDiario"
# # ARQUIVO_TRATADO = os.path.join(PASTA_PROJETO, "Relatorio_Pronto_Para_Colar.xlsx")

# # def gerar_tabela_html(df_filtrado):
# #     """Transforma um DataFrame em uma tabela HTML bonita para o e-mail"""
# #     if df_filtrado.empty:
# #         return "<p style='color: gray;'><i>Nenhum registro encontrado para este filtro.</i></p>"
    
# #     # Seleciona apenas as colunas que você pediu
# #     colunas_exibicao = ['ID', 'COLETA', 'COTACAO', 'ROTA', 'DOCK/AGENDAMENTO', 'TIPO']
# #     df_exibir = df_filtrado[colunas_exibicao]
    
# #     # Gera o HTML com estilo de bordas e cores
# #     return df_exibir.to_html(index=False, border=1, classes='table', escape=False).replace(
# #         'table', 'table style="border-collapse: collapse; width: 100%; font-family: Arial; font-size: 12px;"'
# #     ).replace('<thead>', '<thead style="background-color: #d3d3d3;">').replace('<td>', '<td style="padding: 5px; border: 1px solid #ccc;">').replace('<th>', '<th style="padding: 5px; border: 1px solid #ccc;">')

# # def enviar_email_detalhado():
# #     if not os.path.exists(ARQUIVO_TRATADO):
# #         print("❌ Arquivo não encontrado.")
# #         return

# #     # 1. LER E PREPARAR DADOS
# #     df = pd.read_excel(ARQUIVO_TRATADO)
# #     df['DOCK_DT'] = pd.to_datetime(df['DOCK/AGENDAMENTO'], dayfirst=True).dt.date
    
# #     hoje = datetime.now().date()
# #     amanha = hoje + timedelta(days=1)
    
# #     # 2. INÍCIO DO CORPO DO EMAIL
# #     html_body = f"""
# #     <html>
# #     <body>
# #         <h2 style="color: #2c3e50;">Acompanhamento Diário de Operação - {hoje.strftime('%d/%m/%Y')}</h2>
# #         <p>Olá, segue o detalhamento das operações filtrado por cliente e data:</p>
# #     """

# #     # 3. FILTRAGEM DINÂMICA POR CLIENTE
# #     # Pegamos todos os clientes únicos da coluna 'CLIENTE'
# #     clientes_unicos = df['CLIENTE'].unique()

# #     for cliente in clientes_unicos:
# #         if pd.isna(cliente): continue
        
# #         df_cli = df[df['CLIENTE'] == cliente]
        
# #         # Filtros de Data e Tipo
# #         coletas_hoje = df_cli[(df_cli['TIPO'] == 'COLETA') & (df_cli['DOCK_DT'] == hoje)]
# #         entregas_hoje = df_cli[(df_cli['TIPO'] == 'ENTREGA') & (df_cli['DOCK_DT'] == hoje)]
# #         entregas_amanha = df_cli[(df_cli['TIPO'] == 'ENTREGA') & (df_cli['DOCK_DT'] == amanha)]

# #         # Só adiciona o cliente ao e-mail se ele tiver algo hoje ou amanhã
# #         if not (coletas_hoje.empty and entregas_hoje.empty and entregas_amanha.empty):
# #             html_body += f"<div style='background-color: #f8f9fa; padding: 10px; border-left: 5px solid #2c3e50; margin-top: 20px;'>"
# #             html_body += f"<h3 style='margin: 0;'>🏢 CLIENTE: {cliente}</h3></div>"

# #             html_body += "<h4>🚚 COLETAS PARA HOJE</h4>"
# #             html_body += gerar_tabela_html(coletas_hoje)

# #             html_body += "<h4>📦 ENTREGAS PARA HOJE</h4>"
# #             html_body += gerar_tabela_html(entregas_hoje)

# #             html_body += "<h4>📅 ENTREGAS PARA AMANHÃ</h4>"
# #             html_body += gerar_tabela_html(entregas_amanha)
# #             html_body += "<br><hr>"

# #     html_body += """
# #         <p><i>Relatório gerado automaticamente pelo sistema de monitoramento Transking.</i></p>
# #     </body>
# #     </html>
# #     """

# #     # 4. CONFIGURAÇÃO DO E-MAIL
# #     msg = MIMEMultipart()
# #     msg['From'] = EMAIL_REMETENTE
# #     msg['To'] = EMAIL_DESTINATARIO
# #     msg['Subject'] = f"DETALHAMENTO OPERAÇÃO - {hoje.strftime('%d/%m/%Y')}"
# #     msg.attach(MIMEText(html_body, 'html'))

# #     # 5. ANEXAR O ARQUIVO COMPLETO
# #     try:
# #         with open(ARQUIVO_TRATADO, "rb") as f:
# #             part = MIMEBase('application', 'octet-stream')
# #             part.set_payload(f.read())
# #             encoders.encode_base64(part)
# #             part.add_header('Content-Disposition', f"attachment; filename=Relatorio_Operacao.xlsx")
# #             msg.attach(part)
# #     except Exception as e: print(f"Erro anexo: {e}")

# #     # 6. ENVIO
# #     try:
# #         server = smtplib.SMTP(SMTP_SERVIDOR, SMTP_PORTA)
# #         server.starttls()
# #         server.login(EMAIL_REMETENTE, SENHA_EMAIL)
# #         server.send_message(msg)
# #         server.quit()
# #         print(f"✅ Detalhamento enviado com sucesso para {EMAIL_DESTINATARIO}!")
# #     except Exception as e:
# #         print(f"❌ Erro no envio: {e}")

# # if __name__ == "__main__":
# #     enviar_email_detalhado()











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
#     df_cli = df[df['CLIENTE'].str.contains(cliente, na=False, case=False)]
#     if df_cli.empty: return ""

#     colunas_det = ['ID', 'COLETA', 'COTACAO', 'ROTA', 'DOCK/AGENDAMENTO']
#     larguras_det = {'ID': '100px', 'COLETA': '50px', 'COTACAO': '50px', 'ROTA': '180px', 'DOCK/AGENDAMENTO': '100px'}
    
#     html = f"<div style='background-color: #1f4e78; color: white; padding: 4px; font-size: 11px; margin-top: 10px;'><b>🏢 {cliente}</b></div>"
    
#     for titulo, tipo, data in [("🚚 Coletas Hoje", "COLETA", hoje), 
#                               ("📦 Entregas Hoje", "ENTREGA", hoje), 
#                               ("📅 Entregas Amanhã", "ENTREGA", amanha)]:
#         subset = df_cli[(df_cli['TIPO'] == tipo) & (df_cli['DOCK_DT'] == data)][colunas_det]
#         if not subset.empty:
#             # ADICIONADO O TOTAL NO TÍTULO
#             total_bloco = len(subset)
#             html += f"<p style='margin: 5px 0 2px 0; font-size: 10px;'><b>{titulo}: {total_bloco}</b></p>"
#             html += estilizar_tabela(subset, largura_total="100%", larguras_cols=larguras_det)
#     return html

# def enviar_email_operacao():
#     if not os.path.exists(ARQUIVO_TRATADO): return

#     df = pd.read_excel(ARQUIVO_TRATADO)
#     df['DOCK_DT'] = pd.to_datetime(df['DOCK/AGENDAMENTO'], dayfirst=True).dt.date
#     hoje = datetime.now().date()
#     amanha = hoje + timedelta(days=1)

#     # 1. DINÂMICA COM TOTAL GERAL
#     df_res = df[df['DOCK_DT'].isin([hoje, amanha])]
#     pivot = df_res.pivot_table(index='CLIENTE', columns='TIPO', values='ID', aggfunc='count', fill_value=0).reset_index()
#     pivot.columns.name = None
    
#     # Garante que as colunas existam para somar
#     for col in ['COLETA', 'ENTREGA']:
#         if col not in pivot.columns: pivot[col] = 0
    
#     pivot['TOTAL'] = pivot['COLETA'] + pivot['ENTREGA']
    
#     # LINHA DE TOTAL GERAL NA DINÂMICA
#     total_geral = pd.DataFrame({
#         'CLIENTE': ['TOTAL GERAL'],
#         'COLETA': [pivot['COLETA'].sum()],
#         'ENTREGA': [pivot['ENTREGA'].sum()],
#         'TOTAL': [pivot['TOTAL'].sum()]
#     })
#     pivot = pd.concat([pivot, total_geral], ignore_index=True)
    
#     html_dinamica = estilizar_tabela(pivot, largura_total="350px")

#     # 2. DETALHAMENTO LADO A LADO
#     detalhe_unilever = gerar_bloco_cliente(df, 'UNILEVER', hoje, amanha)
#     detalhe_reckitt = gerar_bloco_cliente(df, 'RECKITT', hoje, amanha)

#     html_body = f"""
#     <html>
#     <body style="margin: 0; padding: 10px; font-family: Segoe UI, Arial;">
#         <h3 style="color: #1f4e78; margin-bottom: 5px;">📊 Resumo Geral - {hoje.strftime('%d/%m/%Y')}</h3>
#         {html_dinamica}
#         <hr style="border: 0; border-top: 1px solid #eee; margin: 15px 0;">
#         <h3 style="color: #c45911; margin-bottom: 10px;">🔍 Detalhamento Clientes </h3>
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
#         print("✅ E-mail enviado com todos os totais!")
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

def estilizar_tabela(df, largura_total="100%", larguras_cols=None):
    df_copy = df.copy()
    for col in df_copy.columns:
        if df_copy[col].dtype in ['float64', 'int64']:
             df_copy[col] = df_copy[col].fillna(0).astype(int).astype(str)

    html = f'<table style="border-collapse: collapse; width: {largura_total}; font-family: Segoe UI, Arial; font-size: 10px; margin-bottom: 10px; border: 1px solid #dee2e6;">'
    html += '<thead style="background-color: #1f4e78; color: white;"><tr>'
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

def gerar_bloco_cliente(df, cliente, hoje, amanha):
    # Filtra pelo cliente alvo
    df_cli = df[df['CLIENTE'].str.contains(cliente, na=False, case=False)]
    if df_cli.empty: return ""

    colunas_det = ['ID', 'COLETA', 'COTACAO', 'ROTA', 'DOCK/AGENDAMENTO']
    larguras_det = {'ID': '100px', 'COLETA': '50px', 'COTACAO': '50px', 'ROTA': '180px', 'DOCK/AGENDAMENTO': '100px'}
    
    html = f"<div style='background-color: #1f4e78; color: white; padding: 4px; font-size: 11px; margin-top: 10px;'><b>🏢 {cliente}</b></div>"
    
    # NOVA LÓGICA DE FILTRO BASEADA NAS DATAS DE COLETA E ENTREGA
    # 1. Coletas Hoje (Sempre que a data de coleta for HOJE)
    col_hoje = df_cli[df_cli['DT_COLETA_OK'] == hoje][colunas_det]
    
    # 2. Entregas Hoje (Sempre que a data de entrega for HOJE)
    ent_hoje = df_cli[df_cli['DT_ENTREGA_OK'] == hoje][colunas_det]
    
    # 3. Entregas Amanhã (Sempre que a data de entrega for AMANHÃ)
    ent_amanha = df_cli[df_cli['DT_ENTREGA_OK'] == amanha][colunas_det]

    for titulo, dados in [("🚚 Coletas Hoje", col_hoje), 
                          ("📦 Entregas Hoje", ent_hoje), 
                          ("📅 Entregas Amanhã", ent_amanha)]:
        if not dados.empty:
            html += f"<p style='margin: 5px 0 2px 0; font-size: 10px;'><b>{titulo}: {len(dados)}</b></p>"
            html += estilizar_tabela(dados, largura_total="100%", larguras_cols=larguras_det)
    return html

def enviar_email_operacao():
    if not os.path.exists(ARQUIVO_TRATADO): return

    df = pd.read_excel(ARQUIVO_TRATADO)
    
    # Converte as colunas de data para comparação
    df['DT_COLETA_OK'] = pd.to_datetime(df['PROGRAMAÇÃO. COLETA'], dayfirst=True).dt.date
    df['DT_ENTREGA_OK'] = pd.to_datetime(df['PROGRAMAÇÃO. ENTREGA'], dayfirst=True).dt.date
    
    hoje = datetime.now().date()
    amanha = hoje + timedelta(days=1)

    # 1. AJUSTE DA DINÂMICA (Baseada na nova lógica de datas)
    # Identifica o que é Coleta ou Entrega hoje/amanhã para o resumo
    df['RESUMO_TIPO'] = 'OUTROS'
    df.loc[df['DT_COLETA_OK'] == hoje, 'RESUMO_TIPO'] = 'COLETA'
    df.loc[df['DT_ENTREGA_OK'].isin([hoje, amanha]), 'RESUMO_TIPO'] = 'ENTREGA'
    
    df_res = df[df['RESUMO_TIPO'] != 'OUTROS']
    pivot = df_res.pivot_table(index='CLIENTE', columns='RESUMO_TIPO', values='ID', aggfunc='count', fill_value=0).reset_index()
    pivot.columns.name = None
    
    for col in ['COLETA', 'ENTREGA']:
        if col not in pivot.columns: pivot[col] = 0
    
    pivot['TOTAL'] = pivot['COLETA'] + pivot['ENTREGA']
    total_geral = pd.DataFrame({'CLIENTE': ['TOTAL GERAL'], 'COLETA': [pivot['COLETA'].sum()], 
                                'ENTREGA': [pivot['ENTREGA'].sum()], 'TOTAL': [pivot['TOTAL'].sum()]})
    pivot = pd.concat([pivot, total_geral], ignore_index=True)
    
    html_dinamica = estilizar_tabela(pivot, largura_total="350px")

    # 2. DETALHAMENTO LADO A LADO
    detalhe_unilever = gerar_bloco_cliente(df, 'UNILEVER', hoje, amanha)
    detalhe_reckitt = gerar_bloco_cliente(df, 'RECKITT', hoje, amanha)

    html_body = f"""
    <html>
    <body style="margin: 0; padding: 10px; font-family: Segoe UI, Arial;">
        <h3 style="color: #1f4e78; margin-bottom: 5px;">📊 Resumo Operacional - {hoje.strftime('%d/%m/%Y')}</h3>
        {html_dinamica}
        <hr style="border: 0; border-top: 1px solid #eee; margin: 15px 0;">
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

    try:
        with open(ARQUIVO_TRATADO, "rb") as f:
            part = MIMEBase('application', 'octet-stream'); part.set_payload(f.read())
            encoders.encode_base64(part); part.add_header('Content-Disposition', 'attachment; filename=Relatorio.xlsx'); msg.attach(part)
        server = smtplib.SMTP(SMTP_SERVIDOR, SMTP_PORTA); server.starttls()
        server.login(EMAIL_REMETENTE, SENHA_EMAIL); server.send_message(msg); server.quit()
        print("✅ E-mail enviado com a nova lógica de filtro (por data)!")
    except Exception as e: print(f"❌ Erro: {e}")

if __name__ == "__main__":
    enviar_email_operacao()