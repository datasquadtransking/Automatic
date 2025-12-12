# -*- coding: utf-8 -*-
import subprocess
import sys
import os

def executar_script(nome_script):
    """
    Função para executar um script Python e aguardar sua conclusão.
    """
    print(f"=====================================================")
    print(f" Iniciando execução do script: {nome_script}")
    print(f"=====================================================")
    
    # Usa 'sys.executable' para garantir que o mesmo interpretador Python 
    # que está rodando este script seja usado para rodar o subprocesso.
    try:
        # Executa o script e espera que ele termine (subprocess.run)
        # capture_output=True e text=True ajudam a capturar a saída, se necessário.
        resultado = subprocess.run([sys.executable, nome_script], 
                                   check=True,  # Levanta erro se o script falhar
                                   capture_output=True, 
                                   text=True)
        
        print(f"\n--- {nome_script} CONCLUÍDO com sucesso. ---")
        
        # Opcional: Imprimir a saída do script executado
        # print("\nSaída do Script:")
        # print(resultado.stdout)

    except subprocess.CalledProcessError as e:
        print(f"\nERRO: O script {nome_script} falhou com código de saída {e.returncode}.")
        print("Saída de Erro:")
        print(e.stderr)
        # Opcional: Parar a execução se um script falhar
        # sys.exit(1)
    except FileNotFoundError:
        print(f"\nERRO: O arquivo {nome_script} não foi encontrado na pasta.")
        # Opcional: Parar a execução se um script não for encontrado
        # sys.exit(1)
    except Exception as e:
        print(f"\nERRO inesperado ao executar {nome_script}: {e}")
        # Opcional: Parar a execução
        # sys.exit(1)


# =================================================================
# LISTA DE SCRIPTS A SEREM EXECUTADOS NA ORDEM
# =================================================================
scripts_a_rodar = [
    "UnileverBrlogColetas.py",
    "UnileverBrlogColetasEntregas.py"
]

if __name__ == "__main__":
    print(f"Iniciando a execução sequencial de {len(scripts_a_rodar)} scripts...")
    
    for script in scripts_a_rodar:
        # Verifica se o arquivo existe antes de tentar rodar
        if os.path.exists(script):
            executar_script(script)
        else:
            print(f"\nAVISO: Pulando a execução de '{script}'. Arquivo não encontrado.")

    print("\n\n--- TODOS OS SCRIPTS FINALIZADOS ---")