import subprocess
import sys
import os

def run_script(script_name):
    # Usamos sys.executable para garantir que estamos usando o mesmo interpretador python
    # O comando é passado como uma lista de argumentos
    print(f"--- Iniciando a execução de: {script_name} ---")
    
    # O comando shell=True é opcional, mas simplifica o caminho se houver espaços ou caracteres especiais (como °)
    try:
        # A função run() espera o script terminar
        result = subprocess.run(['python', script_name], check=True, cwd=os.path.dirname(os.path.abspath(__file__)))
        print(f"--- {script_name} concluído com sucesso. ---")
        return True
    except subprocess.CalledProcessError as e:
        print(f"ERRO: O script {script_name} falhou. Código de retorno: {e.returncode}")
        # Se um script falhar, podemos parar a sequência
        return False
    except FileNotFoundError:
        print(f"ERRO: O arquivo {script_name} não foi encontrado.")
        return False
        
# 1. Executa a primeira automação
if run_script("1°EtapaMeli.py"):
    # 2. Executa a segunda somente se a primeira foi bem-sucedida
    run_script("MatarColetasDedicado.py")