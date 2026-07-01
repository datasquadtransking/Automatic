import os
import subprocess
import sys
import time

# --- CONFIGURAÇÃO DE DIRETÓRIO ---
# Obtém a pasta onde este arquivo 'Parte 4' está salvo
DIRETORIO_ATUAL = os.path.dirname(os.path.abspath(__file__))

# Define o nome exato dos seus arquivos com base na sua imagem
SCRIPTS_PARA_RODAR = [
    "Parte1UnileverCancelamentos.py",
    "Parte2UnileverCancelamentos.py",
    "Parte3UnileverCancelamentos.py",
]


def executar_script_oculto(nome_script):
    caminho_completo = os.path.join(DIRETORIO_ATUAL, nome_script)

    # Verifica se o arquivo realmente existe na pasta antes de rodar
    if not os.path.exists(caminho_completo):
        print(f"❌ Erro: O arquivo {nome_script} não foi encontrado na pasta.")
        return False

    print(
        f"⏳ Inicializando: {nome_script} (Executando em segundo plano)..."
    )

    try:
        # Configuração para Windows não criar uma nova janela de terminal (Prompt de Comando)
        # O creationflags=0x08000000 (CREATE_NO_WINDOW) esconde a janela do terminal do Python
        processo = subprocess.Popen(
            [sys.executable, caminho_completo],
            creationflags=0x08000000,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
        )

        # Aguarda o script atual terminar completamente antes de ir para o próximo
        stdout, stderr = processo.communicate()

        if processo.returncode == 0:
            print(f"✅ Concluído com sucesso: {nome_script}")
            return True
        else:
            print(f"⚠️ {nome_script} terminou com aviso/erro de retorno.")
            if stderr:
                print(f"   Detalhe do erro interno:\n{stderr.strip()}")
            return False

    except Exception as e:
        print(f"❌ Falha grave ao tentar executar {nome_script}: {e}")
        return False


def rodar_fluxo_completo():
    print("==========================================================")
    print("🚀 INICIANDO ORQUESTRADOR GERAL - FLUXO UNILEVER (OCULTO)")
    print("==========================================================\n")

    tempo_inicio = time.time()

    for script in SCRIPTS_PARA_RODAR:
        sucesso = executar_script_oculto(script)

        # Se algum script falhar de forma crítica, você pode decidir se para o fluxo ou continua.
        # Aqui, se a Parte 2 falhar, não faz sentido a Parte 3 tentar ler o arquivo que não existe.
        if not sucesso and script == "Parte2UnileverCancelamentos.py":
            print(
                "\n🛑 Fluxo interrompido: A Parte 2 falhou, cancelando execução da Parte 3."
            )
            break

        print("-" * 58)
        time.sleep(2)  # Pausa de segurança de 2 segundos entre as fases

    tempo_total = time.time() - tempo_inicio
    print("\n==========================================================")
    print(
        f"🏁 FIM DO PROCESSO GERAL! Tempo total: {tempo_total:.2f} segundos."
    )
    print("==========================================================")


if __name__ == "__main__":
    rodar_fluxo_completo()