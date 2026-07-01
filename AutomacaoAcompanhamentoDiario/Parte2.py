import pandas as pd
import numpy as np
import os

# --- CONFIGURAÇÕES DE CAMINHO ---
PASTA_PROJETO = r"C:\Users\flavi\Documents\GitHub\Automatic\AutomacaoAcompanhamentoDiario"
ARQUIVO_ENTRADA = os.path.join(PASTA_PROJETO, "RelatorioAtualizado.xlsx")

def tratar_base_acompanhamento():
    print(f"📂 Lendo arquivo: {ARQUIVO_ENTRADA}")

    try:
        # Lendo o Excel
        df = pd.read_excel(ARQUIVO_ENTRADA, engine='openpyxl')
        
        # Padroniza nomes: Maiúsculo e remove espaços nas pontas
        df.columns = df.columns.str.strip().str.upper()

        # 1. MAPEAMENTO DE STATUS (CONFORME O MENU LATERAL DA IMAGEM)
        mapeamento_status = {
            '1. AG. PLANEJAMENTO': 'COLETA',
            '2. EM GR': 'COLETA',
            '3. AG. COLETA': 'COLETA',
            'EM PESQUISA GR': 'COLETA',
            '4. AG. EMISSÃO': 'ENTREGA',
            '5. AG. A.E': 'ENTREGA',
            '6. EM ROTA': 'ENTREGA',
            '7. AG. DESCARGA': 'ENTREGA',
            '8. AG. COMPROVANTE': 'ENTREGA'
        }
        
        # O nome da coluna no seu BI parece ser apenas 'STATUS'
        col_status = [c for c in df.columns if 'STATUS' in c][0]
        df['TIPO'] = df[col_status].str.strip().map(mapeamento_status)

        # 2. BUSCA DAS DATAS (USANDO OS NOMES REAIS DA IMAGEM)
        # Na sua imagem os nomes são: PROGRAMAÇÃO. COLETA e PROGRAMAÇÃO. ENTREGA
        col_data_coleta = 'PROG. COLETA'
        col_data_entrega = 'PROG. ENTREGA'

        print(f"🕒 Extraindo horários de: {col_data_coleta} e {col_data_entrega}")

        # Condições para escolher a data na nova coluna Dock/Agendamento
        condicoes = [
            (df['TIPO'] == 'COLETA'),
            (df['TIPO'] == 'ENTREGA') | (df['TIPO'] == 'FINALIZADA')
        ]
        
        # Converte para datetime para garantir que o Pandas entenda como data/hora
        data_col = pd.to_datetime(df[col_data_coleta], errors='coerce')
        data_ent = pd.to_datetime(df[col_data_entrega], errors='coerce')
        
        escolhas = [data_col, data_ent]

        # 3. CRIAÇÃO E FORMATAÇÃO DA COLUNA DOCK/AGENDAMENTO
        df['DOCK/AGENDAMENTO'] = np.select(condicoes, escolhas, default=pd.NaT)
        
        # Formata para dd/mm/aaaa hh:mm (O formato que você pediu)
        df['DOCK/AGENDAMENTO'] = pd.to_datetime(df['DOCK/AGENDAMENTO']).dt.strftime('%d/%m/%Y %H:%M')

        # 4. CRIANDO A ROTA (ORIGEM - DESTINO)
        df['ROTA'] = df['ORIGEM'].astype(str) + " - " + df['DESTINO'].astype(str)

        # 5. REORDENANDO AS COLUNAS (CONFORME SUA NECESSIDADE DE COLAR)
        # ID, COLETA, COTACAO, ROTA, DOCK/AGENDAMENTO, TIPO
        # Note: Ajustei 'COTACAO' e 'COLETA' para os nomes que aparecem no seu cabeçalho
        colunas_finais = ['ID', 'COLETA', 'COTAÇÃO', 'ROTA', 'DOCK/AGENDAMENTO', 'TIPO']
        
        # Mantém outras colunas no final por segurança
        outras = [c for c in df.columns if c not in colunas_finais and c not in ['ORIGEM', 'DESTINO']]
        
        df = df[colunas_finais + outras]

        # 6. SALVAR O RESULTADO
        caminho_saida = os.path.join(PASTA_PROJETO, "Relatorio_Pronto_Para_Colar.xlsx")
        df.to_excel(caminho_saida, index=False)
        
        print("\n" + "="*50)
        print("✅ TRATAMENTO FINALIZADO COM SUCESSO!")
        print(f"📍 Arquivo disponível em: {caminho_saida}")
        print("="*50)

    except Exception as e:
        print(f"❌ Erro crítico no processamento: {e}")
        print(f"Colunas detectadas no arquivo: {df.columns.tolist() if 'df' in locals() else 'N/A'}")

if __name__ == "__main__":
    tratar_base_acompanhamento()