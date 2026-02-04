import pandas as pd
import os
import time
import re
import win32com.client as win32
from datetime import datetime
from dateutil.relativedelta import relativedelta
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# =============================================================================
# FUNÇÃO PARA ENVIAR/MONTAR EMAIL COM VISÃO ESPECIAL
# =============================================================================
def enviar_email_destaque(df_especial, periodo, caminho_excel):
    try:
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.Subject = f"Resumo de Despesas - Reckitt e Unilever - {periodo['mes_ano_nome']}"
        mail.To = "seuemail@transking.com.br" # <--- AJUSTE O EMAIL AQUI

        # Criar tabela HTML para o corpo do email
        tabela_html = df_especial.drop(columns=['LinhaBruta']).to_html(index=False, classes='table table-striped')

        mail.HTMLBody = f"""
        <html>
        <head>
            <style>
                table {{ border-collapse: collapse; width: 100%; font-family: Arial; }}
                th {{ background-color: #004a99; color: white; padding: 8px; }}
                td {{ border: 1px solid #ddd; padding: 8px; }}
                tr:nth-child(even) {{ background-color: #f2f2f2; }}
            </style>
        </head>
        <body>
            <p>Olá,</p>
            <p>Conforme solicitado, segue abaixo a visão detalhada dos clientes <b>Reckitt e Unilever</b> para o período de {periodo['mes_ano_nome']}:</p>
            {tabela_html}
            <p>O arquivo Excel completo para estes clientes está em anexo.</p>
            <br>
            <p>Atenciosamente,<br>Automação SSW</p>
        </body>
        </html>
        """
        
        if os.path.exists(caminho_excel):
            mail.Attachments.Add(caminho_excel)
        
        mail.Display() # Abre o Outlook para você revisar antes de enviar
        print("📧 Email de destaque montado no Outlook.")
    except Exception as e:
        print(f"❌ Erro ao interagir com Outlook: {e}")

# =============================================================================
# FUNÇÃO PARA CRIAR NAVEGADOR COM PASTA DE DOWNLOAD ESPECÍFICA
# =============================================================================
def criar_navegador(pasta_download):
    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")
    prefs = {
        "download.default_directory": pasta_download,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    }
    chrome_options.add_experimental_option("prefs", prefs)
    return webdriver.Chrome(options=chrome_options)

# =============================================================================
# FUNÇÃO PARA DEFINIR PERÍODO
# =============================================================================
def obter_periodo(meses_atras=0):
    hoje = datetime.today()
    data_ref = hoje - relativedelta(months=meses_atras)
    primeiro_dia = data_ref.replace(day=1)
    if meses_atras == 0:
        data_fim = hoje
    else:
        data_fim = (primeiro_dia + relativedelta(months=1)) - relativedelta(days=1)
    return {
        "inicio_str": primeiro_dia.strftime("%d%m%y"),
        "fim_str": data_fim.strftime("%d%m%y"),
        "mes_ano_nome": data_ref.strftime("%m_%y"),
        "ano": data_ref.strftime("%Y")
    }

# =============================================================================
# FUNÇÃO PARA TRATAR ARQUIVO SSWWEB E GERAR CSV + EXCEL ESPECIAL
# =============================================================================
def tratar_sswweb_para_csv(caminho_sswweb, pasta_destino, periodo):
    registros = []
    placa_atual = None

    with open(caminho_sswweb, "r", encoding="latin-1") as f:
        for linha in f:
            linha = linha.rstrip()
            match_placa = re.search(r"PLACA:\s+([A-Z0-9]+)", linha)
            if match_placa:
                placa_atual = match_placa.group(1)
                continue

            if re.match(r"\d{2}/\d{2}/\d{2}", linha):
                data_emissao = linha[0:8]
                doc_match = re.search(r"-\s*(\d+)", linha)
                documento = doc_match.group(1) if doc_match else None
                evento_match = re.search(r"\b\d{4}\b", linha)
                evento = evento_match.group(0) if evento_match else None
                desp_match = re.search(r"(\d+,\d{2})$", linha)
                despesas = desp_match.group(1) if desp_match else None

                if despesas and evento:
                    registros.append({
                        "Placa": placa_atual,
                        "DataEmissao": data_emissao,
                        "Documento": documento,
                        "Evento": evento,
                        "Despesas": despesas,
                        "LinhaBruta": linha
                    })

    df = pd.DataFrame(registros)

    # 1. SALVAR ARQUIVO GERAL (CSV original mantido)
    nome_csv = f"DespesasFrota_{periodo['mes_ano_nome']}.csv"
    caminho_csv = os.path.join(pasta_destino, nome_csv)
    df.drop(columns=['LinhaBruta']).to_csv(caminho_csv, index=False, sep=";", encoding="utf-8-sig")
    print(f"📄 CSV gerado com sucesso: {caminho_csv}")

    # 2. VISÃO ESPECIAL: RECKITT E UNILEVER
    df_especial = df[df['LinhaBruta'].str.contains('RECKITT|UNILEVER', case=False, na=False)].copy()
    
    if not df_especial.empty:
        nome_excel = f"Especial_Reckitt_Unilever_{periodo['mes_ano_nome']}.xlsx"
        caminho_excel = os.path.join(pasta_destino, nome_excel)
        df_especial.drop(columns=['LinhaBruta']).to_excel(caminho_excel, index=False)
        print(f"⭐ Excel especial gerado: {caminho_excel}")
        
        # Chama a função de email enviando os dados especiais
        enviar_email_destaque(df_especial, periodo, caminho_excel)
    else:
        print("⚠️ Nenhum registro de Reckitt ou Unilever encontrado para este período.")

# =============================================================================
# FUNÇÃO PARA AGUARDAR DOWNLOAD, TRATAR E REMOVER ARQUIVO BRUTO
# =============================================================================
def renomear_arquivo(pasta_download, periodo):
    pasta_destino = fr"C:\Users\flavi\OneDrive - transking.com.br\PROCESSOS\02 - Bases\19 - Despesas Frota\{periodo['ano']}"
    os.makedirs(pasta_destino, exist_ok=True)

    timeout = 60
    inicio = time.time()
    while True:
        arquivos = [f for f in os.listdir(pasta_download) if not f.endswith(".crdownload")]
        if arquivos: break
        if time.time() - inicio > timeout: raise TimeoutError("Download não concluído.")
        time.sleep(1)

    arquivos_completo = [os.path.join(pasta_download, f) for f in arquivos]
    arquivo_mais_recente = max(arquivos_completo, key=os.path.getctime)
    
    tratar_sswweb_para_csv(arquivo_mais_recente, pasta_destino, periodo)
    os.remove(arquivo_mais_recente)

# =============================================================================
# FUNÇÃO PRINCIPAL E VARIÁVEIS (Mantidos conforme original)
# =============================================================================
CPF, USER_SSW, PASS_SSW = "10620126663", "automac", "A@tki123"

def main():
    for meses_atras in [0, 1]:
        periodo = obter_periodo(meses_atras)
        print(f"\n📅 Processando período: {periodo['mes_ano_nome']}")
        try:
            pasta_download = r"C:\Users\flavi\OneDrive - transking.com.br\PROCESSOS\02 - Bases\19 - Despesas Frota\01 - ArquivoBruto"
            navegador = criar_navegador(pasta_download)
            wait = WebDriverWait(navegador, 20)

            navegador.get("https://sistema.ssw.inf.br/bin/ssw0422")
            wait.until(EC.visibility_of_element_located((By.XPATH, '//input[1]'))).send_keys("TKI")
            wait.until(EC.visibility_of_element_located((By.XPATH, '//input[2]'))).send_keys(CPF)
            wait.until(EC.visibility_of_element_located((By.XPATH, '//input[3]'))).send_keys(USER_SSW)
            wait.until(EC.visibility_of_element_located((By.XPATH, '//input[4]'))).send_keys(PASS_SSW)
            wait.until(EC.element_to_be_clickable((By.XPATH, '//a'))).click()
            time.sleep(5)

            navegador.get("https://sistema.ssw.inf.br/bin/ssw0410")
            campo_inicio = navegador.find_element(By.ID, '3')
            campo_inicio.clear()
            campo_inicio.send_keys(periodo["inicio_str"])
            navegador.find_element(By.ID, '4').send_keys(periodo["fim_str"])
            time.sleep(3)
            navegador.find_element(By.ID, '7').click()

            navegador.get("https://sistema.ssw.inf.br/bin/ssw1440")
            time.sleep(20)
            linha = navegador.find_element(By.XPATH, '//tr[@rid="0"]')
            numero_seq = linha.get_attribute("seq")
            navegador.refresh()
            time.sleep(2)

            while True:
                botao = navegador.find_element(By.XPATH, f'//tr[@seq="{numero_seq}"]/td[9]//a')
                if botao.text.strip() == "Baixar":
                    botao.click()
                    renomear_arquivo(pasta_download, periodo)
                    break
                time.sleep(8)
                navegador.refresh()
            navegador.quit()
        except Exception as e:
            print("❌ Erro:", e)
            try: navegador.quit()
            except: pass

if __name__ == "__main__":
    main()