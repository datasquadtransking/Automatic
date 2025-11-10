import pandas as pd

# Caminho do arquivo original
caminho_arquivo = r"C:\Users\flavi\Desktop\UnileverPontualidade.xlsx"

# Leitura da planilha
df = pd.read_excel(caminho_arquivo)

# Seleção das colunas relevantes
colunas_desejadas = [
    "Shipment",
    "SAP ID",
    "Origem",
    "Dock",
    "CH",
    "Pontualidade",
    "Chegada BRLOG",
    "Pontualidade BRLOG"
]
df_filtrado = df[colunas_desejadas]

# Usa a coluna de data para agrupar (ajustável para "Dock", "CH" ou "Chegada BRLOG")
df_filtrado["Data"] = pd.to_datetime(df_filtrado["Dock"], errors="coerce").dt.date

# Remove linhas sem data válida
df_filtrado = df_filtrado.dropna(subset=["Data"])

# Agrupa por data e calcula os indicadores
tabela_detalhada = df_filtrado.groupby("Data").agg(
    No_Prazo=("Pontualidade", lambda x: (x == "ON TIME").sum()),
    Fora_do_Prazo=("Pontualidade", lambda x: (x == "NO SHOW").sum()),
    Total=("Pontualidade", "count")
).reset_index()

# Calcula SLA diário
tabela_detalhada["SLA"] = (tabela_detalhada["No_Prazo"] / tabela_detalhada["Total"])

# Adiciona coluna de contestação (sem critério, deixamos como 0)
tabela_detalhada["Contestação"] = 0

# Calcula SLA acumulado
tabela_detalhada["SLA Acumulado"] = (
    tabela_detalhada["No_Prazo"].cumsum() / tabela_detalhada["Total"].cumsum()
)

# Formata a data
tabela_detalhada["Data"] = pd.to_datetime(tabela_detalhada["Data"]).dt.strftime("%d/%m/%Y")

# Formata os percentuais como string com símbolo %
tabela_detalhada["SLA"] = (tabela_detalhada["SLA"] * 100).map("{:.2f}%".format)
tabela_detalhada["SLA Acumulado"] = (tabela_detalhada["SLA Acumulado"] * 100).map("{:.2f}%".format)

# Exporta para Excel com nova aba
with pd.ExcelWriter(r"C:\Users\flavi\Desktop\UnileverPontualidade_Com_Detalhamento.xlsx") as writer:
    df_filtrado.to_excel(writer, sheet_name="Base Original", index=False)
    tabela_detalhada.to_excel(writer, sheet_name="Resumo por Dia", index=False)

print("Arquivo exportado com percentuais formatados e guia de detalhamento por dia!")