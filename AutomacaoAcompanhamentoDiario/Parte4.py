import subprocess
import os
import sys
import time

# --- CONFIGURAÇÕES ---
# Pega o caminho da pasta onde este script (Parte4) está salvo
PASTA_PROJETO = os.path.dirname(os.path.abspath(__file__))

# Lista dos scripts na ordem de execução
SCRIPTS = [
    "Parte1.py", # Baixa o arquivo
    "Parte2.py", # Trata os dados
    "Parte3.py"  # Envia o e-mail
]

def executar_automacao_completa():
    print("🚀 INICIANDO AUTOMAÇÃO COMPLETA - TRANSKING")
    print("="*50)
    
    inicio_geral = time.time()

    for script in SCRIPTS:
        caminho_script = os.path.join(PASTA_PROJETO, script)
        
        if not os.path.exists(caminho_script):
            print(f"❌ ERRO: O arquivo {script} não foi encontrado na pasta.")
            return

        print(f"▶️ Executando: {script}...")
        
        # Executa o script e espera ele terminar
        resultado = subprocess.run([sys.executable, caminho_script], capture_output=False)

        if resultado.returncode == 0:
            print(f"✅ {script} finalizado com sucesso.\n")
        else:
            print(f"🛑 ERRO em {script}. A automação foi interrompida.")
            return

    fim_geral = time.time()
    tempo_total = round((fim_geral - inicio_geral) / 60, 2)
    
    print("="*50)
    print(f"🎉 PROCESSO CONCLUÍDO EM {tempo_total} MINUTOS!")
    print("📂 Todos os relatórios foram gerados e o e-mail foi enviado.")

if __name__ == "__main__":
    executar_automacao_completa()