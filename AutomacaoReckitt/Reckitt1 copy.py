import pandas as pd
import chardet
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from datetime import datetime, timedelta
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import time
import os
import glob

# --- 1. CONFIGURAÇÃO ---
options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
download_dir = os.path.join(os.path.expanduser("~"), "Downloads")
arquivo_final_path = os.path.join(download_dir, "AutomacaoReckitt.xlsx")

# SUA LISTA OFICIAL DE COLUNAS (Rigidamente seguida agora)
COLUNAS_DESEJADAS = [
    "Data Criação Carga", "Data Chegada no Cliente", "Data Saída do Cliente", "Data da Entrega",
    "Data/Hora Avaliação", "Data de Carregamento", "Data de Início de Viagem", "Data de Fim de Viagem",
    "Data da Confirmação de Chegada", "Data de Início de Descarregamento", "Data de Término de Descarregamento",
    "Data de início do monitoramento", "Data de fim do monitoramento", "Data de início de carregamento",
    "Data de término de carregamento", "Data do primeiro espelhamento", "Data do último espelhamento",
    "Data de Confirmação da entrega pelo Usuário", "Data da Coleta", "Ordem Prevista", "Ordem Executada",
    "Aderência", "Tempo Permanência", "KM Planejado", "KM Realizado", "Confirmação via App",
    "Encerramento Manual da Viagem", "Filial", "Carga", "Tipo de Operação", "Transportador",
    "Motoristas", "Placas", "Situação da Entrega", "CPF/CNPJ Destinatário", "Cliente Destinatário",
    "Endereço Destinatário", "Bairro Destinatário", "Cidade Destinatário", "Estado Destinatário",
    "CEP Destinatário", "Notas Fiscais", "Pedidos", "Peso Bruto", "Resultado Avaliação",
    "Motivo Avaliação", "Observação Avaliação", "Entrega Fora do Raio", "Raio Médio Viagem (KM)",
    "Quantidade de Animais", "Quantidade de Mortalidade", "Cliente Remetente", "Endereço Remetente",
    "Bairro Remetente", "Cidade Remetente", "Estado Remetente", "Transbordo?", "Percentual Viagem",
    "Longitude Última Posição", "Latitude Última Posição", "Rota", "Origem Fim Monitoramento",
    "Motivo Fim Monitoramento", "Próxima Carga Programada", "Tipo Parada", "Valor Total das Notas",
    "Código Integração Cliente", "Prazo Cliente", "Prazo Transportador", "Previsão Entrega Cliente",
    "Agendamento Entrega Cliente", "Previsão Entrega Transportador", "Modelo Veicular",
    "CPF/CNPJ Remetente", "Finalizador", "Perfil Finalizador", "Número do pedido no cliente",
    "Situação do monitoramento", "Realizada no Prazo?", "Motivo da Retificação", "Shc - Tipo"
]

prefs = {
    "download.default_directory": download_dir,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
}
options.add_experimental_option("prefs", prefs)

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
wait = WebDriverWait(driver, 35) 

# --- FUNÇÕES DE TRATAMENTO ---

def processar_e_mesclar_dados(diretorio, caminho_final):
    """Trata o CSV e garante que APENAS as colunas desejadas existam no arquivo final."""
    time.sleep(5)
    arquivos_csv = glob.glob(os.path.join(diretorio, "*.csv"))
    if not arquivos_csv:
        print("❌ Nenhum CSV novo encontrado.")
        return
    
    ultimo_csv = max(arquivos_csv, key=os.path.getmtime)
    print(f"📄 Processando: {os.path.basename(ultimo_csv)}")

    try:
        # 1. Detectar Encoding
        with open(ultimo_csv, 'rb') as f:
            enc = chardet.detect(f.read(25000))['encoding']
        
        # 2. Ler CSV Novo
        df_novo = pd.read_csv(ultimo_csv, sep=None, engine='python', encoding=enc)

        # 3. Filtrar TRANSKING
        if 'Transportador' in df_novo.columns:
            df_novo = df_novo[df_novo['Transportador'].str.contains('TRANSKING', case=False, na=False)].copy()

        # 4. Criar Coluna "Shc - Tipo"
        if 'Carga' in df_novo.columns and 'Tipo Parada' in df_novo.columns:
            df_novo['Shc - Tipo'] = df_novo['Carga'].astype(str) + "/" + df_novo['Tipo Parada'].astype(str)

        # 5. Mesclar com a base existente
        if os.path.exists(caminho_final):
            print("🔄 Mesclando com a base AutomacaoReckitt.xlsx...")
            df_antigo = pd.read_excel(caminho_final)
            df_consolidado = pd.concat([df_antigo, df_novo], ignore_index=True)
        else:
            print("🆕 Criando nova base...")
            df_consolidado = df_novo

        # --- O AJUSTE CRÍTICO AQUI ---
        # 6. Forçar o DataFrame a ter APENAS as colunas da lista COLUNAS_DESEJADAS
        # Se houver colunas "lixo" no arquivo antigo ou no novo, elas morrem aqui.
        colunas_finais = [c for c in COLUNAS_DESEJADAS if c in df_consolidado.columns]
        df_consolidado = df_consolidado[colunas_finais]

        # 7. Remover Duplicados pela chave
        if 'Shc - Tipo' in df_consolidado.columns:
            df_consolidado = df_consolidado.drop_duplicates(subset=['Shc - Tipo'], keep='first').reset_index(drop=True)

        # 8. Salvar Final (Agora totalmente limpo)
        df_consolidado.to_excel(caminho_final, index=False)
        print(f"✅ Planilha salva e limpa em: {caminho_final}")

        # Limpeza do CSV
        os.remove(ultimo_csv)
        print("🗑️ CSV temporário removido.")

    except Exception as e:
        print(f"❌ Erro no processamento: {e}")

def esperar_download(diretorio):
    print("⏳ Aguardando download...")
    segundos = 0
    while segundos < 400:
        arquivos = os.listdir(diretorio)
        baixando = any([f.endswith(".crdownload") or f.endswith(".tmp") for f in arquivos])
        if any([f.endswith(".csv") for f in arquivos]) and not baixando:
            time.sleep(3)
            return True
        time.sleep(5)
        segundos += 5
    return False

# --- FLUXO PRINCIPAL ---
try:
    if os.path.exists(arquivo_final_path):
        data_alvo = (datetime.now() - timedelta(days=30)).strftime('%d/%m/%Y')
        print(f"📅 Modo Atualização (30 dias): {data_alvo}")
    else:
        data_alvo = f"01/01/{datetime.now().year}"
        print(f"📅 Modo Inicial (Ano Todo): {data_alvo}")

    driver.get("https://reckitt.multitransportador.com.br/Login")
    wait.until(EC.presence_of_element_located((By.ID, "Usuario"))).send_keys("20932-ES")
    driver.find_element(By.ID, "Senha").send_keys("Tk@2025")
    driver.find_element(By.CSS_SELECTOR, "button[type='submit']").click()
    
    time.sleep(5)
    wait.until(EC.element_to_be_clickable((By.XPATH, "//span[text()='Relatórios']"))).click()
    wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "a[title='Paradas']"))).click()
    time.sleep(8)

    campo_data = wait.until(EC.presence_of_element_located((By.XPATH, "//input[contains(@data-bind, 'DataInicial')]")))
    campo_data.clear()
    campo_data.send_keys(data_alvo)
    campo_data.send_keys(Keys.TAB)
    time.sleep(2)
    driver.find_element(By.CLASS_NAME, "subheader-title").click()

    print("👁️ Gerando Preview...")
    btn_preview = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Preview')]")))
    driver.execute_script("arguments[0].click();", btn_preview)
    time.sleep(25) 

    try:
        btn_todos = driver.find_element(By.XPATH, "//label[@for='marcar-desmarcar-colunas-gridPreviewRelatorio']")
        driver.execute_script("arguments[0].click();", btn_todos)
    except: pass

    print("📥 Gerando Excel...")
    btn_excel = wait.until(EC.presence_of_element_located((By.XPATH, "//button[contains(., 'Gerar Planilha Excel')]")))
    driver.execute_script("arguments[0].click();", btn_excel)
    
    if esperar_download(download_dir):
        processar_e_mesclar_dados(download_dir, arquivo_final_path)
    else:
        print("❌ Download falhou.")

except Exception as e:
    print(f"💥 Erro: {e}")

finally:
    driver.quit()
    print("🏁 Finalizado.")