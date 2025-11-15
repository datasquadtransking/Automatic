import os
import shutil


PASTA_LIMPAR = r"C:\Users\flavi\OneDrive\001\BI"
PASTA_COLETAS = r"C:\Users\flavi\OneDrive\001\Coletas"
ARQUIVOS_EXCLUIR = [
    "Coleta_10_25.xlsx",  
    "Coleta_11_25.xlsx"   
]

# -----------------------------------------------------------

def excluir_conteudo_de_pasta(caminho_pasta):
    """Exclui todos os arquivos e subpastas DENTRO do caminho especificado."""
    print(f"\n--- Processando pasta: {caminho_pasta} ---")
    if not os.path.exists(caminho_pasta):
        print(f"⚠️ Erro: A pasta '{caminho_pasta}' não existe. Pulando.")
        return

    try:
       
        for item in os.listdir(caminho_pasta):
            caminho_completo = os.path.join(caminho_pasta, item)
            
            if os.path.isfile(caminho_completo) or os.path.islink(caminho_completo):
                os.unlink(caminho_completo)
                print(f"✅ Arquivo excluído: {item}")
            elif os.path.isdir(caminho_completo):
                shutil.rmtree(caminho_completo)
                print(f"✅ Subpasta excluída: {item}")
        if not os.listdir(caminho_pasta):
            print(f"🎉 Pasta '{caminho_pasta}' limpa com sucesso!")
        else:
             print(f"⚠️ Aviso: Pasta '{caminho_pasta}' pode conter arquivos protegidos/ocultos.")

    except Exception as e:
        print(f"❌ Ocorreu um erro ao limpar a pasta: {e}")


def excluir_arquivos_especificos(caminho_pasta, lista_arquivos):
    """Exclui arquivos específicos de uma pasta."""
    print(f"\n--- Processando arquivos em: {caminho_pasta} ---")
    if not os.path.exists(caminho_pasta):
        print(f"⚠️ Erro: A pasta '{caminho_pasta}' não existe. Pulando.")
        return

    for nome_arquivo in lista_arquivos:
        caminho_completo = os.path.join(caminho_pasta, nome_arquivo)
        
        if os.path.exists(caminho_completo):
            try:
                os.remove(caminho_completo)
                print(f"✅ Arquivo excluído com sucesso: {nome_arquivo}")
            except Exception as e:
                print(f"❌ Erro ao excluir o arquivo '{nome_arquivo}': {e}")
        else:
            print(f"ℹ️ Arquivo não encontrado (já excluído ou nome errado): {nome_arquivo}")

if __name__ == "__main__":
    
    print("\n" + "="*50)
    print("INICIANDO ROTINA DE EXCLUSÃO DE ARQUIVOS (v50.0)")
    print("="*50)
    excluir_conteudo_de_pasta(PASTA_LIMPAR)
    excluir_arquivos_especificos(PASTA_COLETAS, ARQUIVOS_EXCLUIR)
    
    print("\n" + "="*50)
    print("ROTINA DE EXCLUSÃO FINALIZADA.")
    print("="*50)
print("ROTINA DE EXCLUSÃO FINALIZADA.")
