import pandas as pd
import os

# ====================================================================
#   CAMINHOS DOS ARQUIVOS
# ====================================================================
csv_base = r"C:\Users\flavi\Desktop\103.CSV"
xlsx_final = r"C:\Users\flavi\Desktop\Baixar dedicado.xlsx"
PASTA_DOWNLOADS = r"C:\Users\flavi\Downloads"

# ====================================================================
#   CARREGAR CSV
# ====================================================================
df = pd.read_csv(csv_base, sep=";", dtype=str, encoding="latin1", engine="python")
df = df.fillna("")

# ====================================================================
#   FILTRAR SOMENTE OS PAGADORES SOLICITADOS
# ====================================================================
chaves_pagador = ["AMAZON", "BRSP02", "MERCADO ENVIOS", "SHPX"]
regex_pagadores = "|".join(chaves_pagador)

df = df[
    df["PAGADOR"].fillna("").str.contains(regex_pagadores, case=False, regex=True)
]

df = df[
    (df["ULTIMA_OCORRENCIA"] != "11 - VALIDACAO CONCLUIDA") &
    (df["ULTIMA_OCORRENCIA"] != "99 - COLETA CANCELADA")
]

df = df[["NUMERO_COLETA", "PAGADOR", "ULTIMA_OCORRENCIA", "OBSERVACAO_2"]]

# -------------------------------------------------------------------------
# TRATAMENTO DA COLUNA OBSERVACAO_2 (DIVIDIR EM COLETA E ENTREGA)
# -------------------------------------------------------------------------

df[["COLETA_RAW", "ENTREGA_RAW"]] = df["OBSERVACAO_2"].str.split(" - ", n=1, expand=True)

df["COLETA_RAW"] = df["COLETA_RAW"].fillna("")
df["ENTREGA_RAW"] = df["ENTREGA_RAW"].fillna("")

df["COLETA_RAW"] = df["COLETA_RAW"].str.replace("COLETA: ", "", regex=False)
df["ENTREGA_RAW"] = df["ENTREGA_RAW"].str.replace("ENTREGA: ", "", regex=False)

# -------------------------------------------------------------------------
# EXTRAÇÃO DAS DATAS
# -------------------------------------------------------------------------

df["DATA_COLETA_DT"] = df["COLETA_RAW"].str[:10]
df["DATA_ENTREGA_DT"] = df["ENTREGA_RAW"].str[:10]

# -------------------------------------------------------------------------
# CONVERTER DATAS PARA DATETIME
# -------------------------------------------------------------------------

df["DATA_COLETA_DT"] = pd.to_datetime(
    df["DATA_COLETA_DT"], format="%d/%m/%Y", errors="coerce"
)

df["DATA_ENTREGA_DT"] = pd.to_datetime(
    df["DATA_ENTREGA_DT"], format="%d/%m/%Y", errors="coerce"
)

# -------------------------------------------------------------------------
# CÁLCULO DIAS PARA ENTREGA
# -------------------------------------------------------------------------

hoje = pd.Timestamp.today().normalize()

df["DIAS_PARA_ENTREGA"] = (df["DATA_ENTREGA_DT"] - hoje).dt.days
df["DIAS_PARA_ENTREGA"] = df["DIAS_PARA_ENTREGA"].fillna(-999).astype(int)

# ====================================================================
#   🎯 FILTRO: MANTER APENAS DIAS_PARA_ENTREGA <= -8
# ====================================================================
df = df[df["DIAS_PARA_ENTREGA"] <= -8]

# ====================================================================
#   SALVAR XLSX
# ====================================================================
df.to_excel(xlsx_final, index=False)
print(f"💾 Arquivo final salvo em: {xlsx_final}")

# ====================================================================
#   REMOVER TODOS .sswweb DA PASTA DOWNLOADS
# ====================================================================
try:
    for arq in os.listdir(PASTA_DOWNLOADS):
        if arq.endswith(".sswweb"):
            os.remove(os.path.join(PASTA_DOWNLOADS, arq))
    print("🗑 Todos os arquivos .sswweb foram removidos da pasta Downloads.")
except Exception as e:
    print("⚠ Erro ao remover arquivos .sswweb:", e)

print("\n✅ PROCESSO FINALIZADO COM SUCESSO!")
input("Pressione ENTER para sair...")
