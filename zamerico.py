import os
import asyncio
from playwright.async_api import async_playwright

# --- CONFIGURAÇÕES DE ACESSO ---
USER_PBI = "danilo.lacerda@transking.com.br"
PASS_PBI = "@tkiSLI887654498***"
REPORT_URL = "https://app.powerbi.com/groups/59efcb19-4c50-47fc-8dd1-cdcee2357723/reports/e9b6fa02-7de5-4a0c-9c45-05617863d94d/ReportSectionc7a85f51131aa7fd1d6f?experience=power-bi"

async def exportar_tabela_powerbi():
    async with async_playwright() as p:
        # Lançando o navegador (headless=False permite ver a execução)
        browser = await p.chromium.launch(headless=False)
        context = await browser.new_context(accept_downloads=True)
        page = await context.new_page()

        print(f"🔗 Acessando: {REPORT_URL}")
        await page.goto(REPORT_URL)

        # --- 1. PROCESSO DE LOGIN ---
        try:
            print("📧 Inserindo email...")
            await page.wait_for_selector('#email', timeout=20000)
            await page.fill('#email', USER_PBI)
            await page.press('#email', 'Enter')

            print("🔑 Aguardando campo de senha...")
            await page.wait_for_selector('input[type="password"]', timeout=15000)
            await page.fill('input[type="password"]', PASS_PBI)
            await page.click('input[type="submit"]')

            # Lidar com "Manter conectado?"
            try:
                await page.wait_for_selector('#idSIButton9', timeout=5000)
                await page.click('#idSIButton9')
            except:
                pass
        except Exception as e:
            print(f"❌ Erro durante o login: {e}")
            await page.screenshot(path="erro_login.png")
            await browser.close()
            return

        # --- 2. CARREGAMENTO E LOCALIZAÇÃO DA TABELA ---
        print("📊 Aguardando carregamento dos dados (visualização)...")
        try:
            # Esperamos o texto de cabeçalho aparecer para garantir que a tabela renderizou
            await page.wait_for_selector('text=COTACAO', timeout=60000)
            
            # Localizamos o container da tabela "Operação"
            tabela_container = page.locator(".visual-container", has_text="COTACAO").first
            
            # Movemos o mouse para a tabela para o botão "Mais opções" aparecer
            await tabela_container.hover()
            await asyncio.sleep(1) # Pequena pausa para garantir a renderização do botão
            
            # --- 3. CLICANDO NO BOTÃO QUE VOCÊ IDENTIFICOU ---
            print("🔘 Clicando no botão 'Mais opções' (visual-more-options-btn)...")
            btn_mais_opcoes = tabela_container.locator('button[data-testid="visual-more-options-btn"]')
            await btn_mais_opcoes.wait_for(state="visible", timeout=10000)
            await btn_mais_opcoes.click()

            # --- 4. EXPORTAÇÃO ---
            print("📂 Selecionando 'Exportar dados'...")
            await page.wait_for_selector('text="Exportar dados"', state="visible")
            await page.click('text="Exportar dados"')

            # Lidar com o modal de confirmação do Power BI
            print("💾 Confirmando download no modal...")
            async with page.expect_download() as download_info:
                # Clicamos no botão de confirmação que finaliza o processo
                await page.click('button:has-text("Exportar")')
            
            download = await download_info.value
            nome_final = f"./dados_operacao_{download.suggested_filename}"
            await download.save_as(nome_final)
            
            print(f"✅ Arquivo baixado com sucesso: {nome_final}")

        except Exception as e:
            print(f"❌ Erro na interação com a tabela: {e}")
            await page.screenshot(path="erro_tabela.png")

        # Mantém aberto por alguns segundos para conferência antes de fechar
        await asyncio.sleep(5)
        await browser.close()
        print("🏁 Processo finalizado.")

if __name__ == "__main__":
    asyncio.run(exportar_tabela_powerbi())