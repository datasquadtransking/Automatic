import os
import pandas as pd


def processar_desvios():
    # ================= Mapeamento de Configurações =================
    caminho_origem = (
        r"C:\Users\flavi\OneDrive\Weslley\AcompanhamentoWeslley.xlsx"
    )
    caminho_destino = (
        r"C:\Users\flavi\OneDrive\Weslley\Desvios_Validados.xlsx"
    )

    nome_da_aba = "Acompanhamento"

    # Nomes das colunas (sempre em letras minúsculas aqui)
    coluna_coleta = "coleta"  # Mude para 'colclien' se for o caso
    coluna_cotacao = "cotacao"
    coluna_tipo = "tipo"
    coluna_sla = "sla"
    coluna_pagador = "pagador"  # Nova coluna adicionada para o filtro

    # Lista de pagadores que devem ser EXCLUÍDOS da exportação
    # (Escreva tudo em minúsculo aqui para garantir a validação)
    pagadores_para_excluir = [
        "agv logistica s.a - vinhedo",
        "antilhas",
        "antilhas embalagens editora e grafica ltda",
        "HISENSE GORENJE DO BRASIL IMP",
        "aska comercial ltda",
        "bdf nivea ltda",
        "click rodo entregas ltda",
        "comercial de alimentos esthamp",
        "etic",
        "fini",
        "haleon brasil distri. ltda.",
        "haleon do brasil distribuidora ltda",
        "health",
        "hisense gorenje do brasis imp",
        "ingleza",
        "jahu borrachas e autopecas ltd",
        "jv fritais industria e comerci",
        "lab pierre fabre do brasil",
        "lg",
        "lg electronics do brasil ltda.",
        "madeiramadeira comercio eletro",
        "magazine luiza s/a",
        "marluvas",
        "mary kay do brasil ltda",
        "mili s/a",
        "mimo importacao e exportacao s",
        "mk log ltda",
        "nelida do brasil com. e imp. l",
        "saga medicao ltda",
        "samsung sds latin america tecn",
        "supporte armazenagem vendas e",
        "unicharm do brasil ind. e com.",
        "union medic comercio e represe",
        "via varejo s/a",
    ]
    # ===============================================================

    print("Iniciando o processamento da base...")

    if not os.path.exists(caminho_origem):
        print(f"Erro: Arquivo original não encontrado em: {caminho_origem}")
        return

    try:
        df = pd.read_excel(caminho_origem, sheet_name=nome_da_aba, header=4)
    except PermissionError:
        print(
            "\n[ERRO]: O Excel está aberto! Feche o arquivo antes de rodar.\n"
        )
        return
    except Exception as e:
        print(f"Erro ao ler o arquivo: {e}")
        return

    # Padroniza os cabeçalhos
    df.columns = df.columns.str.strip().str.lower()

    # Validação das colunas obrigatórias
    colunas_necessarias = [
        coluna_sla,
        coluna_coleta,
        coluna_cotacao,
        coluna_tipo,
        coluna_pagador,
    ]
    for col in colunas_necessarias:
        if col not in df.columns:
            print(
                f"\n[ERRO DE COLUNA]: A coluna '{col}' não foi encontrada na linha 5."
            )
            print(f"Colunas identificadas: {list(df.columns[:15])}...")
            return

    # 1. Filtrar onde SLA é igual a 0
    df[coluna_sla] = pd.to_numeric(df[coluna_sla], errors="coerce")
    df_filtrado = df[df[coluna_sla] == 0].copy()

    if df_filtrado.empty:
        print("Nenhum registro com SLA = 0 foi encontrado.")
        return

    # 2. NOVO FILTRO: Remover os pagadores da lista de restrição
    # Converte temporariamente os valores da coluna para minúsculo e remove espaços extras
    pagador_limpo = (
        df_filtrado[coluna_pagador].astype(str).str.strip().str.lower()
    )

    # O caractere '~' faz a inversão (ou seja: mantém apenas quem NÃO está na lista)
    df_filtrado = df_filtrado[~pagador_limpo.isin(pagadores_para_excluir)].copy()

    if df_filtrado.empty:
        print(
            "Após filtrar os pagadores restritos, nenhuma linha sobrou para exportação."
        )
        return

    # 3. Criar chave de validação de duplicidade (Coleta - COTACAO - TIPO)
    df_filtrado["chave_duplicidade"] = (
        df_filtrado[coluna_coleta].astype(str).str.strip()
        + "-"
        + df_filtrado[coluna_cotacao].astype(str).str.strip()
        + "-"
        + df_filtrado[coluna_tipo].astype(str).str.strip()
    )

    # 4. Remover duplicados internos da própria planilha do dia
    df_filtrado.drop_duplicates(
        subset=["chave_duplicidade"], keep="first", inplace=True
    )

    # 5. Validar contra a base de destino histórica
    if os.path.exists(caminho_destino):
        print("Cruzando dados com a base histórica...")
        try:
            df_historico = pd.read_excel(caminho_destino)
            df_historico.columns = df_historico.columns.str.strip().str.lower()

            df_historico["chave_duplicidade"] = (
                df_historico[coluna_coleta].astype(str).str.strip()
                + "-"
                + df_historico[coluna_cotacao].astype(str).str.strip()
                + "-"
                + df_historico[coluna_tipo].astype(str).str.strip()
            )

            chaves_existentes = set(df_historico["chave_duplicidade"].dropna())
            df_novos = df_filtrado[
                ~df_filtrado["chave_duplicidade"].isin(chaves_existentes)
            ].copy()

            if df_novos.empty:
                print("Todos os desvios coletados já existem no histórico.")
                return

            df_final = pd.concat([df_historico, df_novos], ignore_index=True)
            df_final.drop(columns=["chave_duplicidade"], inplace=True)

        except Exception as e:
            print(
                f"Aviso ao ler histórico ({e}). Criando um novo arquivo de destino."
            )
            df_final = df_filtrado.drop(columns=["chave_duplicidade"])
    else:
        print("Criando uma nova base histórica de desvios...")
        df_final = df_filtrado.drop(columns=["chave_duplicidade"])

    # 6. Exportar resultado final
    try:
        df_final.to_excel(caminho_destino, index=False)
        print(f"\n=> SUCESSO! Dados exportados para: {caminho_destino}")
        print(f"=> Total de registros na base consolidada: {len(df_final)}")
    except Exception as e:
        print(f"Erro ao salvar arquivo de destino: {e}")


if __name__ == "__main__":
    processar_desvios()