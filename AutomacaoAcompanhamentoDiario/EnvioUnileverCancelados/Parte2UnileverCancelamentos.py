
# import os
# import time
# from datetime import datetime
# import pandas as pd
# from selenium import webdriver
# from selenium.webdriver.common.by import By
# from selenium.webdriver.common.keys import Keys
# from selenium.webdriver.support.ui import WebDriverWait
# from selenium.webdriver.support import expected_conditions as EC

# # --- CONFIGURAÇÕES DE CAMINHO ---
# ARQUIVO_ENTRADA = r"C:\Users\flavi\OneDrive - transking.com.br\PROCESSOS\01 - Equipe\04 - Weslley\Unilevercancelamentos.xlsx"
# ARQUIVO_SAIDA = r"C:\Users\flavi\OneDrive - transking.com.br\PROCESSOS\01 - Equipe\04 - Weslley\Unilevercancelamentos_Processado.xlsx"

# def preparar_dados():
#     try:
#         # Lendo o arquivo específico de cancelamentos
#         df = pd.read_excel(ARQUIVO_ENTRADA, engine='openpyxl')
#         df.columns = df.columns.str.strip().str.upper()

#         if 'ID' not in df.columns:
#             print("❌ Erro: Coluna 'ID' não encontrada no arquivo.")
#             return None

#         # Tratando o ID para remover o '.0' caso seja lido como float e garantir apenas texto limpo
#         df['ID_CONSULTA'] = df['ID'].dropna().astype(str).apply(lambda x: x.split('.')[0].strip())
        
#         # Filtra linhas onde o ID ficou vazio após o tratamento
#         df = df[df['ID_CONSULTA'] != ''].copy()

#         # Cria a nova coluna para armazenar o status capturado do portal
#         df["STATUS_PORTAL"] = "---"

#         return df
#     except Exception as e:
#         print(f"❌ Erro na leitura do arquivo: {e}")
#         return None

# def realizar_consultas_portal(df):
#     if df is None or df.empty:
#         print("⚠ Nenhuma ID válida para consultar.")
#         return

#     options = webdriver.ChromeOptions()
#     options.add_argument("--start-maximized")
#     driver = webdriver.Chrome(options=options)
#     wait = WebDriverWait(driver, 15)

#     try:
#         # Login no portal
#         driver.get("https://unilever2.brasilrisk.com.br/Login/Logout")
#         wait.until(EC.presence_of_element_located((By.ID, "usuario"))).send_keys("MONITORAMENTO1")
#         driver.find_element(By.ID, "senha").send_keys("@Tk@2026")
#         driver.find_element(By.ID, "Login").click()
        
#         # Aguarda carregar a página inicial após login
#         time.sleep(10)

#         for index, row in df.iterrows():
#             id_viagem = row["ID_CONSULTA"]
            
#             tentativas_maximas = 5
#             sucesso = False
            
#             for tentativa in range(1, tentativas_maximas + 1):
#                 print(f"🔄 Consultando ID: {id_viagem} ({index + 1}/{len(df)}) - Tentativa {tentativa}/{tentativas_maximas}")

#                 try:
#                     # 1. Clica no botão para abrir/ativar a caixa de busca
#                     botao_busca = wait.until(EC.element_to_be_clickable((By.ID, "botaosearch")))
#                     botao_busca.click()
#                     time.sleep(1)

#                     # 2. Localiza o input de texto dentro da busca ativa
#                     campo_input = wait.until(EC.element_to_be_clickable((By.ID, "input-search")))
#                     campo_input.clear()
#                     time.sleep(1)
#                     campo_input.send_keys(id_viagem)
#                     time.sleep(1)
#                     campo_input.send_keys(Keys.ENTER)
                    
#                     # --- ESPERA DINÂMICA DE ATÉ 60 SEGUNDOS ---
#                     WebDriverWait(driver, 60).until(
#                         EC.any_of(
#                             EC.presence_of_element_located((By.ID, "txtStatus")),
#                             EC.presence_of_element_located((By.ID, "swal2-content"))
#                         )
#                     )
#                     time.sleep(1)

#                     # 3. Verifica se apareceu o alerta de "SM não encontrada"
#                     alertas = driver.find_elements(By.ID, "swal2-content")
#                     if alertas and "SM não encontrada" in alertas[0].text:
#                         print(f"⚠ ID {id_viagem} não encontrada no portal. Aplicando rotina de reset (F5)...")
                        
#                         # Rotina solicitada: ENTER -> Espera 2s -> ESC -> Espera 2s -> F5 -> Espera 3s
#                         webdriver.ActionChains(driver).send_keys(Keys.ENTER).perform()
#                         time.sleep(2)
#                         webdriver.ActionChains(driver).send_keys(Keys.ESCAPE).perform()
#                         time.sleep(2)
#                         driver.refresh()
#                         time.sleep(3)
                        
#                         df.at[index, "STATUS_PORTAL"] = "Não Encontrada"
#                         sucesso = True
#                         break # Sai do laço de tentativas pois o resultado foi conclusivo

#                     # 4. Se não deu erro, espera o elemento do status carregar na tela
#                     status_elemento = wait.until(EC.presence_of_element_located((By.ID, "txtStatus")))
#                     status_texto = status_elemento.text.strip()
                    
#                     print(f"✅ Status capturado: {status_texto}")
#                     df.at[index, "STATUS_PORTAL"] = status_texto
#                     sucesso = True
#                     break # Sai do laço de tentativas (sucesso!)

#                 except Exception as erro_loop:
#                     print(f"❌ Tentativa {tentativa} falhou para a ID {id_viagem}. Erro: {erro_loop}")
                    
#                     # Se ainda restarem tentativas, dá um F5 e espera o portal respirar antes de tentar de novo
#                     if tentativa < tentativas_maximas:
#                         print("🔄 Atualizando a página para tentar novamente...")
#                         driver.refresh()
#                         time.sleep(5) # Aumentei para 5 segundos para garantir estabilidade no refresh
#                     else:
#                         # Se esgotou todas as tentativas, marca como erro
#                         print(f"☠ Esgotadas as {tentativas_maximas} tentativas para a ID {id_viagem}.")
#                         df.at[index, "STATUS_PORTAL"] = "Erro na Consulta"
#                         driver.refresh()
#                         time.sleep(3)

#         # --- ALTERAÇÕES SOLICITADAS ANTES DE SALVAR ---
#         print("🔧 Tratando colunas finais antes de gerar o Excel...")
        
#         # 1. Junta as colunas ORIGEM e DESTINO separando por " - "
#         if 'ORIGEM' in df.columns and 'DESTINO' in df.columns:
#             df['ROTA'] = df['ORIGEM'].fillna('').astype(str) + " - " + df['DESTINO'].fillna('').astype(str)
#             df = df.drop(columns=['ORIGEM', 'DESTINO'], errors='ignore')

#         # 2. Remove as colunas CLIENTE, CURVA e a coluna temporária ID_CONSULTA
#         colunas_para_remover = ['CLIENTE', 'CURVA', 'ID_CONSULTA']
#         df = df.drop(columns=colunas_para_remover, errors='ignore')

#         # Salva o resultado final no novo arquivo Excel
#         df.to_excel(ARQUIVO_SAIDA, index=False, engine='openpyxl')
#         print(f"🏁 Processo concluído com sucesso! Arquivo salvo em: {ARQUIVO_SAIDA}")

#     except Exception as e:
#         print(f"❌ Erro geral no robô: {e}")
#     finally:
#         driver.quit()

# # --- EXECUÇÃO DO FLUXO ---
# if __name__ == "__main__":
#     base_dados = preparar_dados()
#     realizar_consultas_portal(base_dados)

import os
import time
from datetime import datetime
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Importações necessárias para o envio de e-mail
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

# --- CONFIGURAÇÕES DE CAMINHO ---
ARQUIVO_ENTRADA = r"C:\Users\flavi\OneDrive - transking.com.br\PROCESSOS\01 - Equipe\04 - Weslley\Unilevercancelamentos.xlsx"
ARQUIVO_SAIDA = r"C:\Users\flavi\OneDrive - transking.com.br\PROCESSOS\01 - Equipe\04 - Weslley\Unilevercancelamentos_Processado.xlsx"

# --- CONFIGURAÇÕES DE E-MAIL ---
EMAIL_REMETENTE = "automac@transking.com.br"
SENHA_EMAIL = "P3dr0Tk2025"
EMAIL_DESTINATARIO = "weslley.transking@gmail.com"
SMTP_SERVIDOR = "smtp.task.com.br"
SMTP_PORTA = 587


def preparar_dados():
    try:
        # Lendo o arquivo específico de cancelamentos
        df = pd.read_excel(ARQUIVO_ENTRADA, engine='openpyxl')
        df.columns = df.columns.str.strip().str.upper()

        if 'ID' not in df.columns:
            print("❌ Erro: Coluna 'ID' não encontrada no arquivo.")
            return None

        # Tratando o ID para remover o '.0' caso seja lido como float e garantir apenas texto limpo
        df['ID_CONSULTA'] = df['ID'].dropna().astype(str).apply(lambda x: x.split('.')[0].strip())
        
        # Filtra linhas onde o ID ficou vazio após o tratamento
        df = df[df['ID_CONSULTA'] != ''].copy()

        # Cria a nova coluna para armazenar o status capturado do portal
        df["STATUS_PORTAL"] = "---"

        return df
    except Exception as e:
        print(f"❌ Erro na leitura do arquivo: {e}")
        return None


def enviar_email_alerta(df_cancelados):
    """Função responsável por estruturar e enviar o e-mail de alerta com a tabela de cancelados"""
    print("📧 Iniciando processo de envio de e-mail de alerta...")
    
    try:
        # Criação da mensagem
        msg = MIMEMultipart()
        msg['From'] = EMAIL_REMETENTE
        msg['To'] = EMAIL_DESTINATARIO
        msg['Subject'] = f"🚨 Alerta: IDs Cancelados Detectados no Portal - {datetime.now().strftime('%d/%m/%Y')}"

        # Monta uma tabela simples em HTML para exibir no corpo do e-mail
        # Caso queira remover as colunas não tratadas, filtramos apenas as principais
        colunas_exibir = [col for col in ['ID', 'STATUS_PORTAL', 'ROTA'] if col in df_cancelados.columns]
        if not colunas_exibir:
            colunas_exibir = df_cancelados.columns

        tabela_html = df_cancelados[colunas_exibir].to_html(index=False, classes='table table-striped')

        corpo_email = f"""
        <html>
            <body style="font-family: Arial, sans-serif;">
                <h2 style="color: #cc0000;">Atenção: Viagens com Status Cancelado Identificadas</h2>
                <p>O robô identificou registros cancelados durante a última consulta ao portal Unilever.</p>
                <p><b>Detalhes dos itens cancelados:</b></p>
                {tabela_html}
                <br>
                <p><i>Este é um e-mail automático gerado pelo sistema de monitoramento Transking.</i></p>
            </body>
        </html>
        """
        msg.attach(MIMEText(corpo_email, 'html'))

        # Conexão com o servidor SMTP
        server = smtplib.SMTP(SMTP_SERVIDOR, SMTP_PORTA)
        server.starttls()  # Ativa a segurança TLS
        server.login(EMAIL_REMETENTE, SENHA_EMAIL)
        
        # Envia o e-mail
        server.sendmail(EMAIL_REMETENTE, EMAIL_DESTINATARIO, msg.as_string())
        server.quit()
        
        print("✅ E-mail de alerta enviado com sucesso!")
    except Exception as e:
        print(f"❌ Falha ao enviar o e-mail: {e}")


def realizar_consultas_portal(df):
    if df is None or df.empty:
        print("⚠ Nenhuma ID válida para consultar.")
        return

    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    driver = webdriver.Chrome(options=options)
    wait = WebDriverWait(driver, 15)

    try:
        # Login no portal
        driver.get("https://unilever2.brasilrisk.com.br/Login/Logout")
        wait.until(EC.presence_of_element_located((By.ID, "usuario"))).send_keys("MONITORAMENTO1")
        driver.find_element(By.ID, "senha").send_keys("@Tk@2026")
        driver.find_element(By.ID, "Login").click()
        
        # Aguarda carregar a página inicial após login
        time.sleep(10)

        for index, row in df.iterrows():
            id_viagem = row["ID_CONSULTA"]
            
            tentativas_maximas = 5
            sucesso = False
            
            for tentativa in range(1, tentativas_maximas + 1):
                print(f"🔄 Consultando ID: {id_viagem} ({index + 1}/{len(df)}) - Tentativa {tentativa}/{tentativas_maximas}")

                try:
                    # 1. Clica no botão para abrir/ativar a caixa de busca
                    botao_busca = wait.until(EC.element_to_be_clickable((By.ID, "botaosearch")))
                    botao_busca.click()
                    time.sleep(1)

                    # 2. Localiza o input de texto dentro da busca ativa
                    campo_input = wait.until(EC.element_to_be_clickable((By.ID, "input-search")))
                    campo_input.clear()
                    time.sleep(1)
                    campo_input.send_keys(id_viagem)
                    time.sleep(1)
                    campo_input.send_keys(Keys.ENTER)
                    
                    # --- ESPERA DINÂMICA DE ATÉ 60 SEGUNDOS ---
                    WebDriverWait(driver, 60).until(
                        EC.any_of(
                            EC.presence_of_element_located((By.ID, "txtStatus")),
                            EC.presence_of_element_located((By.ID, "swal2-content"))
                        )
                    )
                    time.sleep(1)

                    # 3. Verifica se apareceu o alerta de "SM não encontrada"
                    alertas = driver.find_elements(By.ID, "swal2-content")
                    if alertas and "SM não encontrada" in alertas[0].text:
                        print(f"⚠ ID {id_viagem} não encontrada no portal. Aplicando rotina de reset (F5)...")
                        
                        # Rotina solicitada: ENTER -> Espera 2s -> ESC -> Espera 2s -> F5 -> Espera 3s
                        webdriver.ActionChains(driver).send_keys(Keys.ENTER).perform()
                        time.sleep(2)
                        webdriver.ActionChains(driver).send_keys(Keys.ESCAPE).perform()
                        time.sleep(2)
                        driver.refresh()
                        time.sleep(3)
                        
                        df.at[index, "STATUS_PORTAL"] = "Não Encontrada"
                        sucesso = True
                        break  # Sai do laço de tentativas pois o resultado foi conclusivo

                    # 4. Se não deu erro, espera o elemento do status carregar na tela
                    status_elemento = wait.until(EC.presence_of_element_located((By.ID, "txtStatus")))
                    status_texto = status_elemento.text.strip()
                    
                    print(f"✅ Status capturado: {status_texto}")
                    df.at[index, "STATUS_PORTAL"] = status_texto
                    sucesso = True
                    break  # Sai do laço de tentativas (sucesso!)

                except Exception as erro_loop:
                    print(f"❌ Tentativa {tentativa} falhou para a ID {id_viagem}. Erro: {erro_loop}")
                    
                    # Se ainda restarem tentativas, dá um F5 e espera o portal respirar antes de tentar de novo
                    if tentativa < tentativas_maximas:
                        print("🔄 Atualizando a página para tentar novamente...")
                        driver.refresh()
                        time.sleep(5)
                    else:
                        # Se esgotou todas as tentativas, marca como erro
                        print(f"☠ Esgotadas as {tentativas_maximas} tentativas para a ID {id_viagem}.")
                        df.at[index, "STATUS_PORTAL"] = "Erro na Consulta"
                        driver.refresh()
                        time.sleep(3)

        # --- ALTERAÇÕES SOLICITADAS ANTES DE SALVAR ---
        print("🔧 Tratando colunas finais antes de gerar o Excel...")
        
        # 1. Junta as colunas ORIGEM e DESTINO separando por " - "
        if 'ORIGEM' in df.columns and 'DESTINO' in df.columns:
            df['ROTA'] = df['ORIGEM'].fillna('').astype(str) + " - " + df['DESTINO'].fillna('').astype(str)
            df = df.drop(columns=['ORIGEM', 'DESTINO'], errors='ignore')

        # 2. Remove as colunas CLIENTE, CURVA e a coluna temporária ID_CONSULTA
        colunas_para_remover = ['CLIENTE', 'CURVA', 'ID_CONSULTA']
        df = df.drop(columns=colunas_para_remover, errors='ignore')

        # --- NOVA ROTINA: VERIFICAÇÃO DE CANCELADOS ---
        # Filtra o DataFrame procurando por termos que contenham "CANCEL" (ignora maiúsculas/minúsculas)
        df_cancelados = df[df['STATUS_PORTAL'].str.contains('CANCEL', case=False, na=False)]

        if not df_cancelados.empty:
            print(f"🚨 Alerta! Foram encontrados {len(df_cancelados)} itens cancelados.")
            enviar_email_alerta(df_cancelados)
        else:
            print("👍 Nenhuma ID cancelada foi detectada nesta rodada.")

        # Salva o resultado final no novo arquivo Excel
        df.to_excel(ARQUIVO_SAIDA, index=False, engine='openpyxl')
        print(f"🏁 Processo concluído com sucesso! Arquivo salvo em: {ARQUIVO_SAIDA}")

    except Exception as e:
        print(f"❌ Erro geral no robô: {e}")
    finally:
        driver.quit()


# --- EXECUÇÃO DO FLUXO ---
if __name__ == "__main__":
    base_dados = preparar_dados()
    realizar_consultas_portal(base_dados)