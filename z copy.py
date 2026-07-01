# import time
# import pyautogui as py

# for i in range(53):
#     time.sleep(10)
#     py.click(x=415, y=254)
#     time.sleep(2)
#     py.click(x=594, y=250)

import requests
from bs4 import BeautifulSoup
import pandas as pd

# URL do site
url = "https://www.fundamentus.com.br/fii_resultado.php"

headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
}

response = requests.get(url, headers=headers)

if response.status_code == 200:
    soup = BeautifulSoup(response.content, 'html.parser')
    table = soup.find('table', {'id': 'tabelaResultado'})
    
    headers = [header.text.strip() for header in table.find_all('th')]
    
    rows = []
    for row in table.find_all('tr')[1:]:
        cols = [col.text.strip() for col in row.find_all('td')]
        if cols:
            rows.append(cols)
    
    df = pd.DataFrame(rows, columns=headers)
    
    # Remove colunas indesejadas
    colunas_remover = ['Preço do m²', 'Aluguel por m²', 'Endereço']
    df = df.drop(columns=[col for col in colunas_remover if col in df.columns])
    
    # Conversões
    df['Cotação'] = df['Cotação'].str.replace('.', '', regex=False).str.replace(',', '.', regex=False).astype(float)
    df['P/VP'] = df['P/VP'].str.replace('.', '', regex=False).str.replace(',', '.', regex=False).astype(float)
    df['Vacância Média'] = df['Vacância Média'].str.replace('.', '', regex=False).str.replace(',', '.', regex=False).str.replace('%','', regex=False).astype(float)
    df['Dividend Yield'] = df['Dividend Yield'].str.replace('.', '', regex=False).str.replace(',', '.', regex=False).str.replace('%','', regex=False).astype(float) / 100
    
    for col in ['Liquidez', 'Valor de Mercado']:
        df[col] = df[col].str.replace('.', '', regex=False).str.replace(',', '.', regex=False).astype(float)
    
    # Filtros
    df = df[
        (df['P/VP'] >= 0.85) & (df['P/VP'] <= 1.15) & 
        (df['Vacância Média'] <= 10) & 
        (df['Liquidez'] >= 500000) & 
        (df['Valor de Mercado'] >= 500000000) & 
        (df['Dividend Yield'] >= 0.07) & (df['Dividend Yield'] <= 0.14)
    ]
    
    # Classificação por base
    def classificar_base(valor):
        if valor < 15:
            return 'Base 10'
        elif valor < 80:
            return 'Base 50'
        elif valor < 120:
            return 'Base 100'
        else:
            return 'Base 100+'
    
    df['Base'] = df['Cotação'].apply(classificar_base)
    
    # Criar coluna de pontuação
    df['Pontos'] = 0
    
    # Aplicar regras de pontuação por base
    for base, grupo in df.groupby('Base'):
        idx_valor = grupo.nlargest(3, 'Valor de Mercado').index
        idx_pvp = grupo.nsmallest(3, 'P/VP').index
        idx_yield = grupo.nlargest(3, 'Dividend Yield').index
        
        df.loc[idx_valor, 'Pontos'] += 1
        df.loc[idx_pvp, 'Pontos'] += 1
        df.loc[idx_yield, 'Pontos'] += 1
    
    # Mostrar top 5 de cada classe
    for base, grupo in df.groupby('Base'):
        print(f"\n=== {base} ===")
        top5 = grupo.sort_values(by='Pontos', ascending=False).head(5)
        print(top5[['Papel', 'Cotação', 'P/VP', 'Dividend Yield', 'Valor de Mercado', 'Pontos']])
    
    # Salvar no Excel também
    df.to_excel('fii_resultado_rank.xlsx', index=False)
    print("\nTabela salva como 'fii_resultado_rank.xlsx'.")
else:
    print(f"Erro ao acessar o site. Código de status: {response.status_code}")
