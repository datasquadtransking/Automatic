import pandas as pd
import os

caminho_arquivo = r'C:\Users\flavi\Desktop\Coletas.csv'

if os.path.exists(caminho_arquivo) and os.path.getsize(caminho_arquivo) > 0:
    try:
        # 1. LEITURA
        df = pd.read_csv(caminho_arquivo, sep=None, engine='python', encoding='utf-8-sig')

        # 2. LIMPEZA DOS CABEÇALHOS
        df.columns = [
            str(c).replace('ï»¿', '').replace('Ã', '').replace('Â', '').strip() 
            for c in df.columns
        ]

        # 3. LIMPEZA DAS CÉLULAS
        df = df.replace({'Ã': '', 'Â': '', ' ,': ','}, regex=True)

        # Configuração de nomes de colunas de origem
        col_ocorr = 'ULTIMA_OCORRENCIA'
        col_data = 'DATA_ULTIMA_OCORRENCIA'
        col_hora = 'HORA_ULTIMA_OCORRENCIA'
        
        # --- 4. LÓGICA DE NEGÓCIO ---

        #Ocorrência 79 (Com regra de limpeza pelo 97)
        col_nova_79 = '79.EntrandoEmGr'
        if col_ocorr in df.columns:
            if col_nova_79 not in df.columns: df[col_nova_79] = ""
            df[col_nova_79] = df[col_nova_79].fillna("").astype(str)
            
            mask_97 = df[col_ocorr].astype(str).str.contains('97', na=False)
            df.loc[mask_97, col_nova_79] = ""
            
            mask_79 = (df[col_ocorr].astype(str).str.contains('79', na=False)) & (df[col_nova_79].str.strip() == "")
            if mask_79.any():
                df.loc[mask_79, col_nova_79] = df.loc[mask_79, col_data].astype(str) + " " + df.loc[mask_79, col_hora].astype(str)


        # --- JUNÇÃO DAS COLUNAS DE COLETA ---
        col_col_data = 'COLETADO_DATA'
        col_col_hora = 'COLETADO_HORA'
        if col_col_data in df.columns and col_col_hora in df.columns:
            df['ColetadoDataHora'] = df[col_col_data].astype(str) + " " + df[col_col_hora].astype(str)
            df = df.drop(columns=[col_col_data, col_col_hora])

        # C) Ocorrência 34 (NOVO: Acumular a primeira vez)
        col_nova_34 = '34.EntrouEmRota'
        if col_ocorr in df.columns:
            if col_nova_34 not in df.columns: df[col_nova_34] = ""
            df[col_nova_34] = df[col_nova_34].fillna("").astype(str)
            
            # Filtra onde a ocorrência é 34 e a coluna nova ainda está vazia
            mask_34 = (df[col_ocorr].astype(str).str.contains('34', na=False)) & (df[col_nova_34].str.strip() == "")
            if mask_34.any():
                df.loc[mask_34, col_nova_34] = df.loc[mask_34, col_data].astype(str) + " " + df.loc[mask_34, col_hora].astype(str)

        #Ocorrência 17 (Acumular a primeira vez)
        col_nova_17 = '17.ChegadaNaEntrega'
        if col_ocorr in df.columns:
            if col_nova_17 not in df.columns: df[col_nova_17] = ""
            df[col_nova_17] = df[col_nova_17].fillna("").astype(str)
            
            mask_17 = (df[col_ocorr].astype(str).str.contains('17', na=False)) & (df[col_nova_17].str.strip() == "")
            if mask_17.any():
                df.loc[mask_17, col_nova_17] = df.loc[mask_17, col_data].astype(str) + " " + df.loc[mask_17, col_hora].astype(str)

        # 5. SALVAMENTO
        df.to_csv(caminho_arquivo, index=False, encoding='utf-8-sig', sep=';')
        print("Sucesso! Ocorrências 79, 17 e 34 processadas e colunas de coleta mescladas.")

    except Exception as e:
        print(f"Erro: {e}")
else:
    print("Arquivo vazio ou não encontrado.")