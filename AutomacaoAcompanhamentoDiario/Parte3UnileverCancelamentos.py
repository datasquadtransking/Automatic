# from datetime import datetime
# from email import encoders
# from email.mime.base import MIMEBase
# from email.mime.multipart import MIMEMultipart
# from email.mime.text import MIMEText
# import os
# import smtplib
# import pandas as pd

# # --- CONFIGURAÇÕES DE CAMINHO (Mesmo arquivo de saída da Parte 2) ---
# ARQUIVO_PROCESSADO = r"C:\Users\flavi\OneDrive - transking.com.br\PROCESSOS\01 - Equipe\04 - Weslley\Unilevercancelamentos_Processado.xlsx"

# # --- CONFIGURAÇÕES DE E-MAIL ---
# EMAIL_REMETENTE = "automac@transking.com.br"
# SENHA_EMAIL = "P3dr0Tk2025"
# EMAIL_DESTINATARIO = "weslley.transking@gmail.com"
# SMTP_SERVIDOR = "smtp.task.com.br"
# SMTP_PORTA = 587


# def enviar_email_alerta(df_cancelados):
#     """Estrutura e envia o e-mail de alerta com a tabela de cancelados"""
#     print("📧 Iniciando processo de envio de e-mail de alerta...")

#     try:
#         # Criação da mensagem
#         msg = MIMEMultipart()
#         msg["From"] = EMAIL_REMETENTE
#         msg["To"] = EMAIL_DESTINATARIO
#         msg[
#             "Subject"
#         ] = f"🚨 Alerta: IDs Cancelados Detectados no Portal - {datetime.now().strftime('%d/%m/%Y')}"

#         # Seleciona as colunas disponíveis para exibir no corpo do e-mail
#         colunas_exibir = [
#             col
#             for col in ["ID", "STATUS_PORTAL", "ROTA"]
#             if col in df_cancelados.columns
#         ]
#         if not colunas_exibir:
#             colunas_exibir = df_cancelados.columns

#         # Gera a tabela em HTML
#         tabela_html = df_cancelados[colunas_exibir].to_html(
#             index=False, classes="table table-striped"
#         )

#         corpo_email = f"""
#         <html>
#             <body style="font-family: Arial, sans-serif;">
#                 <h2 style="color: #cc0000;">Atenção: Viagens com Status Cancelado Identificadas</h2>
#                 <p>O sistema identificou registros cancelados no arquivo processado da Unilever.</p>
#                 <p><b>Detalhes dos itens cancelados:</b></p>
#                 {tabela_html}
#                 <br>
#                 <p><i>Este é um e-mail automático gerado pelo sistema de monitoramento Transking.</i></p>
#             </body>
#         </html>
#         """
#         msg.attach(MIMEText(corpo_email, "html"))

#         # Conexão com o servidor SMTP da Task
#         server = smtplib.SMTP(SMTP_SERVIDOR, SMTP_PORTA)
#         server.starttls()  # Ativa a segurança TLS
#         server.login(EMAIL_REMETENTE, SENHA_EMAIL)

#         # Envia o e-mail
#         server.sendmail(EMAIL_REMETENTE, EMAIL_DESTINATARIO, msg.as_string())
#         server.quit()

#         print("✅ E-mail de alerta enviado com sucesso!")

#     except Exception as e:
#         print(f"❌ Falha ao enviar o e-mail: {e}")


# def executar_parte_3():
#     print("🚀 Executando PARTE 3: Verificação de Cancelados e Envio de E-mail")

#     # 1. Verifica se o arquivo gerado pela Parte 2 realmente existe
#     if not os.path.exists(ARQUIVO_PROCESSADO):
#         print(
#             f"❌ Erro: O arquivo processado não foi encontrado em: {ARQUIVO_PROCESSADO}"
#         )
#         return

#     try:
#         # 2. Lê o arquivo gerado pela Parte 2
#         df = pd.read_excel(ARQUIVO_PROCESSADO, engine="openpyxl")

#         if "STATUS_PORTAL" not in df.columns:
#             print(
#                 "❌ Erro: Coluna 'STATUS_PORTAL' não encontrada na planilha processada."
#             )
#             return

#         # 3. Filtra buscando por qualquer variação de "CANCEL" (Cancelado, CANCELADO, etc.)
#         df_cancelados = df[
#             df["STATUS_PORTAL"].str.contains("CANCEL", case=False, na=False)
#         ]

#         # 4. Se houver cancelados, dispara a função de e-mail
#         if not df_cancelados.empty:
#             print(
#                 f"🚨 Alerta! Foram encontrados {len(df_cancelados)} itens cancelados no arquivo."
#             )
#             enviar_email_alerta(df_cancelados)
#         else:
#             print(
#                 "👍 Tudo limpo! Nenhuma ID cancelada foi detectada no arquivo processado."
#             )

#     except Exception as e:
#         print(f"❌ Erro ao processar a Parte 3: {e}")


# if __name__ == "__main__":
#     executar_parte_3()

import os
import smtplib
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import pandas as pd

# --- CONFIGURAÇÕES DE CAMINHO (Mesmo arquivo de saída da Parte 2) ---
ARQUIVO_PROCESSADO = r"C:\Users\flavi\OneDrive - transking.com.br\PROCESSOS\01 - Equipe\04 - Weslley\Unilevercancelamentos_Processado.xlsx"

# --- CONFIGURAÇÕES DE E-MAIL (Com múltiplos destinatários) ---
EMAIL_REMETENTE = "automac@transking.com.br"
SENHA_EMAIL = "P3dr0Tk2025"
SMTP_SERVIDOR = "smtp.task.com.br"
SMTP_PORTA = 587

# Lista com todos os e-mails que devem receber o alerta
EMAILS_DESTINATARIOS = [
    "weslley.transking@gmail.com",
    "equipe.monitoramento@transking.com.br",
    "equipe.comercial@transking.com.br",
]


def enviar_email_alerta(df_cancelados):
    """Estrutura e envia o e-mail de alerta com a tabela de cancelados para múltiplos destinatários"""
    print("📧 Iniciando processo de envio de e-mail de alerta...")

    try:
        # Criação da mensagem
        msg = MIMEMultipart()
        msg["From"] = EMAIL_REMETENTE

        # Transforma a lista de e-mails em um texto único separado por vírgulas para o cabeçalho do e-mail
        msg["To"] = ", ".join(EMAILS_DESTINATARIOS)

        msg[
            "Subject"
        ] = f"🚨 Alerta: IDs Cancelados Detectados no Portal - {datetime.now().strftime('%d/%m/%Y')}"

        # Seleciona as colunas disponíveis para exibir no corpo do e-mail
        colunas_exibir = [
            col
            for col in ["ID","COLETA", "ROTA","STATUS"]
            if col in df_cancelados.columns
        ]
        if not colunas_exibir:
            colunas_exibir = df_cancelados.columns

        # Gera a tabela em HTML formatada
        tabela_html = df_cancelados[colunas_exibir].to_html(
            index=False, classes="table table-striped"
        )

        corpo_email = f"""
        <html>
            <body style="font-family: Arial, sans-serif;">
                <h2 style="color: #cc0000;">Atenção: Viagens com Status Cancelado Identificadas</h2>
                <p>O sistema identificou registros cancelados no arquivo processado da Unilever.</p>
                <p><b>Detalhes dos itens cancelados:</b></p>
                {tabela_html}
                <br>
                <p><i>Este é um e-mail automático gerado pelo sistema de monitoramento Transking.</i></p>
            </body>
        </html>
        """
        msg.attach(MIMEText(corpo_email, "html"))

        # Conexão com o servidor SMTP da Task
        server = smtplib.SMTP(SMTP_SERVIDOR, SMTP_PORTA)
        server.starttls()  # Ativa a segurança TLS
        server.login(EMAIL_REMETENTE, SENHA_EMAIL)

        # Envia o e-mail passando a lista completa de destinatários
        server.sendmail(EMAIL_REMETENTE, EMAILS_DESTINATARIOS, msg.as_string())
        server.quit()

        print("✅ E-mail de alerta enviado com sucesso para toda a equipe!")

    except Exception as e:
        print(f"❌ Falha ao enviar o e-mail: {e}")


def executar_parte_3():
    print("🚀 Executando PARTE 3: Verificação de Cancelados e Envio de E-mail")

    # 1. Verifica se o arquivo gerado pela Parte 2 realmente existe
    if not os.path.exists(ARQUIVO_PROCESSADO):
        print(
            f"❌ Erro: O arquivo processado não foi encontrado em: {ARQUIVO_PROCESSADO}"
        )
        return

    try:
        # 2. Lê o arquivo gerado pela Parte 2
        df = pd.read_excel(ARQUIVO_PROCESSADO, engine="openpyxl")

        if "STATUS_PORTAL" not in df.columns:
            print(
                "❌ Erro: Coluna 'STATUS_PORTAL' não encontrada na planilha processada."
            )
            return

        # 3. Filtra buscando por qualquer variação de "CANCEL" (Cancelado, CANCELADO, etc.)
        df_cancelados = df[
            df["STATUS_PORTAL"].str.contains("CANCEL", case=False, na=False)
        ]

        # 4. Se houver cancelados, dispara a função de e-mail
        if not df_cancelados.empty:
            print(
                f"🚨 Alerta! Foram encontrados {len(df_cancelados)} itens cancelados no arquivo."
            )
            enviar_email_alerta(df_cancelados)
        else:
            print(
                "👍 Tudo limpo! Nenhuma ID cancelada foi detectada no arquivo processado."
            )

    except Exception as e:
        print(f"❌ Erro ao processar a Parte 3: {e}")


if __name__ == "__main__":
    executar_parte_3()