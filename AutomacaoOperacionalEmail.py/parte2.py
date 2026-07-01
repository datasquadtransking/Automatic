import os
import smtplib
import pandas as pd
from datetime import datetime, timedelta
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

# --- CONFIGURAÇÕES ---
PASTA_DESKTOP = os.path.join(os.path.expanduser("~"), "Desktop")
xlsx_final = os.path.join(PASTA_DESKTOP, "AcompanhamentoOperacionalBi.xlsx")

EMAIL_REMETENTE = "automac@transking.com.br"
SENHA_EMAIL = "P3dr0Tk2025"
EMAIL_DESTINATARIO = "weslley.transking@gmail.com"

# --- ADICIONE OS EMAILS EM CÓPIA AQUI (separados por vírgula) ---
EMAILS_COPIA = "coordenacao.monitoramento@transking.com.br,coordenacao.monitoramento1@transking.com.br " 

SMTP_SERVIDOR = "smtp.task.com.br"
SMTP_PORTA = 587

def estilizar_tabela_html(df, largura_total="100%", cor_cabecalho="#1f4e78"):
    """Estilização baseada no primeiro script com bloqueio de quebra de linha"""
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
    
    html_body = f"""
    <html>
    <body style="margin: 0; padding: 20px; font-family: Segoe UI, Arial; color: #333;">
        <p>{saudacao}, Weslley.</p>
        <h2 style="color: #1f4e78; border-bottom: 2px solid #1f4e78; padding-bottom: 5px;">
            📊 Acompanhamento Operacional- {agora.strftime('%d/%m/%Y')}
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
    msg['Cc'] = EMAILS_COPIA  # Adiciona os e-mails no cabeçalho visual do e-mail
    msg['Subject'] = f"REPORT: Acompanhamento Operacional - {datetime.now().strftime('%d/%m/%Y')}"

    # Lista consolidada de quem vai receber (To + Cc) para o servidor SMTP
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
        # O pulo do gato: enviamos para a lista consolidada 'todos_destinatarios'
        server.sendmail(EMAIL_REMETENTE, todos_destinatarios, msg.as_string())
        server.quit()
        print(f"Relatório enviado com sucesso para {EMAIL_DESTINATARIO} e em cópia para {EMAILS_COPIA}!")
    except Exception as e:
        print(f"Falha ao enviar e-mail: {e}")

if __name__ == "__main__":
    enviar_email_com_anexo(xlsx_final)