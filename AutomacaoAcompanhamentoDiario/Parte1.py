import os
import asyncio
from playwright.async_api import async_playwright
from datetime import datetime # Para colocar a data no nome, se quiser

# --- CONFIGURAÇÕES DE ACESSO ---
USER_PBI = "danilo.lacerda@transking.com.br"
PASS_PBI = "80w3R81TK4@N#2026"
REPORT_URL = "https://app.powerbi.com/groups/59efcb19-4c50-47fc-8dd1-cdcee2357723/reports/e9b6fa02-7de5-4a0c-9c45-05617863d94d/ReportSectionc7a85f51131aa7fd1d6f?experience=power-bi"

async def exportar_tabela_powerbi():
    # --- DESCOBRINDO A PASTA REAL DO SCRIPT ---
    # Isso garante que o arquivo salve onde o seu .py está, não importa o terminal
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
            # --- LOGIN ---
            await page.wait_for_selector('#email', timeout=20000)
            await page.fill('#email', USER_PBI)
            await page.press('#email', 'Enter')

            await page.wait_for_selector('input[type="password"]', timeout=15000)
            await page.fill('input[type="password"]', PASS_PBI)
            await page.click('input[type="submit"]')

            try:
                await page.wait_for_selector('#idSIButton9', timeout=8000)
                await page.click('#idSIButton9')
            except:
                pass

            # --- NAVEGAÇÃO E CLIQUES ---
            print("⏳ Localizando tabela e abrindo menus...")
            seletor_destino = 'div[data-query-ref="fColetas.DESTINOE/UF"]'
            coluna_destino = await page.wait_for_selector(seletor_destino, timeout=60000)
            
            await page.bring_to_front()
            await coluna_destino.hover()
            await page.wait_for_timeout(2000)

            await page.click('button[data-testid="visual-more-options-btn"]')
            
            seletor_menu_exportar = 'button[data-testid="pbimenu-item.Exportar dados"]'
            await page.wait_for_selector(seletor_menu_exportar, state="visible")
            await page.click(seletor_menu_exportar)

            # --- DOWNLOAD ---
            seletor_botao_confirmar = 'button[data-testid="export-btn"]'
            await page.wait_for_selector(seletor_botao_confirmar, state="visible")

            print("🔘 Clicando em Exportar e capturando download...")
            async with page.expect_download() as download_info:
                await page.click(seletor_botao_confirmar)
            
            download = await download_info.value
            
            # --- SALVAMENTO GARANTIDO ---
            # Nome do arquivo com a data de hoje para facilitar
            data_hoje = datetime.now().strftime("%Y-%m-%d_%H-%M")
            nome_arquivo = f"RelatorioAtualizado.xlsx"
            
            # Montamos o caminho completo
            caminho_final = os.path.join(pasta_do_script, nome_arquivo)
            
            # Salva o arquivo
            await download.save_as(caminho_final)
            
            print("\n" + "="*50)
            print(f"✅ SUCESSO! ARQUIVO SALVO!")
            print(f"📂 PASTA: {pasta_do_script}")
            print(f"📄 ARQUIVO: {nome_arquivo}")
            print(f"📍 CAMINHO COMPLETO: {caminho_final}")
            print("="*50 + "\n")

        except Exception as e:
            print(f"❌ Erro: {e}")
        
        finally:
            print("🧹 Fechando navegador...")
            await page.wait_for_timeout(3000)
            await browser.close()

if __name__ == "__main__":
    asyncio.run(exportar_tabela_powerbi())