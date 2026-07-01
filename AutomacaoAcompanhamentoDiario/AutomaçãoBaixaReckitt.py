import os
import asyncio
from playwright.async_api import async_playwright
from datetime import datetime
import pandas as pd

# --- CONFIGURAÇÕES DE ACESSO E CAMINHOS ---
USER_PBI = "danilo.lacerda@transking.com.br"
PASS_PBI = "QwErTyAsDfgH!2#4%6"
REPORT_URL = "https://app.powerbi.com/groups/59efcb19-4c50-47fc-8dd1-cdcee2357723/reports/e1a37785-9bef-43e0-8354-8392ecd3bf6f/bba973db55ef86a780da?experience=power-bi&clientSideAuth=0"

PASTA_DESTINO = r"C:\Users\flavi\OneDrive - transking.com.br\PROCESSOS\01 - Equipe\04 - Weslley"

async def exportar_tabela_powerbi():
    pasta_do_script = os.path.dirname(os.path.abspath(__file__))
    
    async with async_playwright() as p:
        print(f"🚀 Iniciando navegador...")
        
        browser = await p.chromium.launch(
            headless=False,
            args=["--start-maximized", "--window-position=0,0"]
        )

        context = await browser.new_context(
            accept_downloads=True,
            viewport=None 
        )
        
        page = await context.new_page()

        print(f"🔗 Acessando Power BI...")
        await page.goto(REPORT_URL)

        try:
            # --- LOGIN MICROSOFT ---
            print("🔑 Preenchendo e-mail...")
            await page.wait_for_selector('#email', timeout=20000)
            await page.fill('#email', USER_PBI)
            await page.press('#email', 'Enter')

            print("🔑 Preenchendo senha...")
            await page.wait_for_selector('input[type="password"]', timeout=15000)
            await page.fill('input[type="password"]', PASS_PBI)
            await page.click('input[type="submit"]')

            try:
                await page.wait_for_selector('#idSIButton9', timeout=8000)
                await page.click('#idSIButton9')
                print("✅ Login efetuado com sucesso!")
            except:
                print("⚠️ Botão de 'Manter conectado' não apareceu, prosseguindo...")

            # --- INTERVALO PARA CARREGAMENTO ---
            print("⏳ Aguardando 5 segundos para carregar o Power BI...")
            await asyncio.sleep(5) 
            
            print(f"🔄 Recarregando a página do relatório...")
            await page.goto(REPORT_URL)
            
            print("⏳ Aguardando 5 segundos adicionais para renderização...")
            await asyncio.sleep(5)

            await page.bring_to_front()
            
            # --- FLUXO DE CLIQUES ---
            print("🔍 Procurando a coluna 'ID' (COD AUXILIAR)...")
            seletor_coluna_id = 'div[role="columnheader"][data-query-ref="Cotacao.COD AUXILIAR"]'
            coluna_id = await page.wait_for_selector(seletor_coluna_id, timeout=60000)
            
            print("Main-cell 🖱️ Clicando na coluna 'ID' para ativar o visual...")
            await coluna_id.hover()
            await page.wait_for_timeout(1000)
            await coluna_id.click()
            await page.wait_for_timeout(2000)

            print("🖱️ Clicando nos 3 pontinhos (...)")
            seletor_tres_pontos = 'button[data-testid="visual-more-options-btn"]'
            await page.wait_for_selector(seletor_tres_pontos, state="visible", timeout=20000)
            await page.click(seletor_tres_pontos)
            
            print("📂 Selecionando 'Exportar dados'...")
            seletor_menu_exportar = 'button[data-testid="pbimenu-item.Exportar dados"]'
            await page.wait_for_selector(seletor_menu_exportar, state="visible", timeout=15000)
            await page.click(seletor_menu_exportar)

            print("⏳ Aguardando janela de confirmação de exportação...")
            seletor_botao_confirmar = 'button[data-testid="export-btn"]'
            await page.wait_for_selector(seletor_botao_confirmar, state="visible", timeout=15000)

            print("🔘 Clicando em Exportar e baixando o arquivo...")
            async with page.expect_download() as download_info:
                await page.click(seletor_botao_confirmar)
            
            download = await download_info.value
            
            if not os.path.exists(PASTA_DESTINO):
                os.makedirs(PASTA_DESTINO)

            caminho_temp = os.path.join(pasta_do_script, "temp_pbi.xlsx")
            caminho_final = os.path.join(PASTA_DESTINO, "FINALIZARRECKITT.xlsx")
            
            await download.save_as(caminho_temp)
            print(f"📥 Dados novos baixados temporariamente.")

            # --- TRATAMENTO DOS DADOS (PANDAS) ---
            print("🧹 Processando dados novos e aplicando formatações...")
            
            df_novo = pd.read_excel(caminho_temp)
            df_novo.columns = df_novo.columns.str.strip()

            # 1. Filtrar Clientes (Reckitt e Vestacy)
            if 'CLIENTE' in df_novo.columns:
                df_novo['CLIENTE'] = df_novo['CLIENTE'].astype(str).str.strip()
                df_novo = df_novo[df_novo['CLIENTE'].str.upper().isin(['VESTACY', 'RECKITT'])]
            
            # 2. Concatenar ORIGEM e DESTINO
            if 'ORIGEM' in df_novo.columns and 'DESTINO' in df_novo.columns:
                df_novo['ORIGEM - DESTINO'] = df_novo['ORIGEM'].astype(str).str.strip() + ' - ' + df_novo['DESTINO'].astype(str).str.strip()

            # 3. Formatar as colunas de data/hora
            colunas_data = ['PROG. COLETA', 'PROG. ENTREGA']
            for col in colunas_data:
                if col in df_novo.columns:
                    df_novo[col] = pd.to_datetime(df_novo[col], errors='coerce')
                    df_novo[col] = df_novo[col].dt.strftime('%d/%m/%y %H:%M').fillna('')

            # 4. Mapear e Renomear colunas essenciais
            if 'MERCADORIA' in df_novo.columns:
                df_novo = df_novo.rename(columns={'MERCADORIA': 'COTAÇÃO'})
                
            coluna_id_real = 'ID' if 'ID' in df_novo.columns else ('COD AUXILIAR' if 'COD AUXILIAR' in df_novo.columns else None)
            if coluna_id_real and coluna_id_real != 'ID':
                df_novo = df_novo.rename(columns={coluna_id_real: 'ID'})

            # 5. 🛠️ AJUSTADO: LÓGICA DA MÁSCARA DO ID (PADRÃO SHC-XXXXXX-XXXX) 🛠️
            if 'ID' in df_novo.columns:
                def tratar_e_mascarar_id(val):
                    if pd.isna(val):
                        return ""
                    
                    # Converte para string e limpa decimais flutuantes do Excel
                    txt = str(val).strip().split('.')[0]
                    
                    # Verifica se contém apenas números
                    if txt.isdigit():
                        # Se já tiver entre 1 e 10 dígitos, normaliza para 10 dígitos com zeros à esquerda
                        if 0 < len(txt) <= 10 and int(txt) > 0:
                            txt = txt.zfill(10)
                            # Retorna no padrão pedido: Bloco de 6 dígitos - Bloco de 4 dígitos
                            return f"SHC-{txt[:6]}-{txt[6:]}"
                    
                    # Se não cair nas regras acima (ou se não for numérico), mantém o valor original
                    return val

                df_novo['ID'] = df_novo['ID'].apply(tratar_e_mascarar_id)

            # Adiciona a coluna de controle vazia para novos registros antes de filtrar
            df_novo['STATUS ANALISTA'] = ""

            # --- DEFINIÇÃO DA ORDEM FINAL DAS COLUNAS ---
            ordem_colunas_final = [
                'COLETA', 'COTAÇÃO', 'ID', 'PROG. COLETA', 'PROG. ENTREGA', 
                'ORIGEM - DESTINO', 'DESTINATARIO', 'MOTORISTA', 'PLACA', 'STATUS ANALISTA'
            ]

            # Garantir que as colunas existam no dataframe novo antes de reorganizar
            for col in ordem_colunas_final:
                if col not in df_novo.columns:
                    df_novo[col] = ""

            # Filtra e ordena o dataframe novo baseado na estrutura desejada
            df_novo = df_novo[ordem_colunas_final]

            # --- VERIFICAÇÃO DE DUPLICIDADE ---
            if os.path.exists(caminho_final):
                print("💾 Planilha existente encontrada no OneDrive. Lendo dados antigos...")
                df_antigo = pd.read_excel(caminho_final)
                df_antigo.columns = df_antigo.columns.str.strip()
                
                if 'ID' in df_antigo.columns:
                    df_antigo['ID'] = df_antigo['ID'].astype(str).str.strip()
                    
                    ids_antigos = df_antigo['ID'].unique()
                    df_filtrado_novos = df_novo[~df_novo['ID'].isin(ids_antigos)]
                    
                    if not df_filtrado_novos.empty:
                        print(f"➕ Encontradas {len(df_filtrado_novos)} novas linhas para inserir.")
                        df_resultado = pd.concat([df_antigo, df_filtrado_novos], ignore_index=True)
                    else:
                        print("ℹ️ Nenhuma linha nova encontrada no Power BI desta vez.")
                        df_resultado = df_antigo
                else:
                    print("⚠️ Coluna 'ID' não encontrada na planilha antiga. Sobrescrevendo por segurança.")
                    df_resultado = df_novo
            else:
                print("🆕 Criando planilha inicial diretamente no OneDrive (primeira execução).")
                df_resultado = df_novo

            # --- AJUSTES FINAIS DE FORMATAÇÃO E REORDENAÇÃO COMPLETA ---
            if 'COTAÇÃO' in df_resultado.columns:
                df_resultado['COTAÇÃO'] = pd.to_numeric(df_resultado['COTAÇÃO'], errors='coerce').astype('Int64')

            colunas_finais_validas = [col for col in ordem_colunas_final if col in df_resultado.columns]
            df_resultado = df_resultado[colunas_finais_validas]

            # Salva atualizado na nuvem do Weslley
            df_resultado.to_excel(caminho_final, index=False)
            
            if os.path.exists(caminho_temp):
                os.remove(caminho_temp)

            print("\n" + "="*60)
            print(f"🎉 PROCESSO CONCLUÍDO COM SUCESSO!")
            print(f"📍 BASE ATUALIZADA E FORMATADA: {caminho_final}")
            print("="*60 + "\n")

        except Exception as e:
            print(f"❌ Erro durante o processo: {e}")
            try:
                caminho_erro = os.path.join(pasta_do_script, "erro_captura.png")
                await page.screenshot(path=caminho_erro)
                print(f"📸 Evidência de erro gravada em: {caminho_erro}")
            except:
                pass
        
        finally:
            print("🧹 Encerrando a instância do navegador...")
            await page.wait_for_timeout(4000)
            await browser.close()

if __name__ == "__main__":
    asyncio.run(exportar_tabela_powerbi())