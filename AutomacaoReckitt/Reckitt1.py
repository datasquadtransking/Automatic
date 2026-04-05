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

# LISTA DE COLUNAS OFICIAIS (O que deve permanecer)
COLUNAS_OFICIAIS = [
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
    "Situação do monitoramento", "Realizada no Prazo?", "Motivo da Retificação"
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
actions = ActionChains(driver)

# --- FUNÇÕES DE TRATAMENTO ---

def processar_e_mesclar_dados(diretorio, caminho_final):
    """Trata o CSV, limpa colunas extras, cria chave única e mescla."""
    time.sleep(5)
    arquivos_csv = glob.glob(os.path.join(diretorio, "*.csv"))
    if not arquivos_csv:
        print("❌ Nenhum CSV novo encontrado.")
        return
    
    ultimo_csv = max(arquivos_csv, key=os.path.getmtime)
    print(f"📄 Processando download: {os.path.basename(ultimo_csv)}")

    try:
        # Detectar Encoding
        with open(ultimo_csv, 'rb') as f:
            enc = chardet.detect(f.read(25000))['encoding']
        
        # Ler CSV
        df_novo = pd.read_csv(ultimo_csv, sep=None, engine='python', encoding=enc)

        # 1. FILTRAR APENAS AS COLUNAS DESEJADAS
        # Isso remove automaticamente o "0", "TRANSKING TRANSPORTES LTDA", "Unnamed: 4", etc.
        colunas_presentes = [col for col in COLUNAS_OFICIAIS if col in df_novo.columns]
        df_novo = df_novo[colunas_presentes].copy()
        print("🧹 Colunas extras e cabeçalhos de erro removidos.")

        # 2. Filtrar apenas TRANSKING
        if 'Transportador' in df_novo.columns:
            df_novo = df_novo[df_novo['Transportador'].str.contains('TRANSKING', case=False, na=False)].copy()

        # 3. Criar Coluna "Shc - Tipo"
        if 'Carga' in df_novo.columns and 'Tipo Parada' in df_novo.columns:
            df_novo['Shc - Tipo'] = df_novo['Carga'].astype(str) + "/" + df_novo['Tipo Parada'].astype(str)

        # 4. Mesclar com a base existente
        if os.path.exists(caminho_final):
            print("🔄 Mesclando com a base AutomacaoReckitt.xlsx...")
            df_antigo = pd.read_excel(caminho_final)
            df_consolidado = pd.concat([df_antigo, df_novo], ignore_index=True)
        else:
            print("🆕 Criando base inicial...")
            df_consolidado = df_novo

        # 5. Remover duplicados pela nova chave
        if 'Shc - Tipo' in df_consolidado.columns:
            total_antes = len(df_consolidado)
            df_consolidado = df_consolidado.drop_duplicates(subset=['Shc - Tipo'], keep='first').reset_index(drop=True)
            print(f"✨ Linhas únicas na base final: {len(df_consolidado)} (Removidas: {total_antes - len(df_consolidado)})")

        # Salvar
        df_consolidado.to_excel(caminho_final, index=False)
        print(f"✅ Sucesso! Planilha atualizada em: {caminho_final}")

        # Limpeza do CSV
        os.remove(ultimo_csv)

    except Exception as e:
        print(f"❌ Erro no tratamento dos dados: {e}")

def esperar_download(diretorio):
    print("⏳ Aguardando conclusão do download...")
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
        print(f"📅 Modo: Últimos 30 dias ({data_alvo})")
    else:
        data_alvo = f"01/01/{datetime.now().year}"
        print(f"📅 Modo: Ano atual completo ({data_alvo})")

    driver.get("https://reckitt.multitransportador.com.br/Login")
    wait.until(EC.presence_of_element_located((By.ID, "Usuario"))).send_keys("20932-ES")
    driver.find_element(By.ID, "Senha").send_keys("Tk@2025")
    driver.find_element(By.CSS_SELECTOR, "button[type='submit']").click()
    
    time.sleep(5)
    menu = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[text()='Relatórios']")))
    menu.click()
    wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "a[title='Paradas']"))).click()
    time.sleep(8)

    campo_data = wait.until(EC.presence_of_element_located((By.XPATH, "//input[contains(@data-bind, 'DataInicial')]")))
    campo_data.clear()
    campo_data.send_keys(data_alvo)
    campo_data.send_keys(Keys.TAB)
    time.sleep(2)
    driver.find_element(By.CLASS_NAME, "subheader-title").click()

    print("👁️ Abrindo Preview...")
    btn_preview = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Preview')]")))
    driver.execute_script("arguments[0].click();", btn_preview)
    time.sleep(25) 

    try:
        btn_todos = driver.find_element(By.XPATH, "//label[@for='marcar-desmarcar-colunas-gridPreviewRelatorio']")
        driver.execute_script("arguments[0].click();", btn_todos)
    except: pass

    print("📥 Solicitando Planilha...")
    btn_excel = wait.until(EC.presence_of_element_located((By.XPATH, "//button[contains(., 'Gerar Planilha Excel')]")))
    driver.execute_script("arguments[0].click();", btn_excel)
    
    if esperar_download(download_dir):
        processar_e_mesclar_dados(download_dir, arquivo_final_path)
    else:
        print("❌ Download não detectado.")

except Exception as e:
    print(f"💥 Erro: {e}")

finally:
    driver.quit()
    print("🏁 Finalizado.")