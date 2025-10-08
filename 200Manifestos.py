import shutil
import time
import calendar
from datetime import datetime
from dateutil.relativedelta import relativedelta
import os
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException


# Função Igor Relatório 200
def baixar_relatorios_200_em_lote(
    navegador,
    usuario_ssw="automac",
    senha_ssw="A@tki123",
    empresa="TKI",
    cnpj="10620126663",
    pasta_destino=r"C:\Users\flavi\OneDrive\002\04-MANIFESTOS",
    meses_retroativos=1
):
    def gerar_periodos(inicio, fim):
        datas = []
        atual = inicio
        hoje = datetime.today()

        while atual <= fim:
            primeiro_dia = atual.replace(day=1)

            if atual.year == hoje.year and atual.month == hoje.month:
                ultimo_dia = hoje
            else:
                ultimo_dia = atual.replace(day=calendar.monthrange(atual.year, atual.month)[1])

            nome_mes = atual.strftime("%b").lower() + str(atual.year)
            datas.append((
                primeiro_dia.strftime("%d%m%y"),
                ultimo_dia.strftime("%d%m%y"),
                nome_mes
            ))

            atual += relativedelta(months=1)

        return datas

    def esperar_relatorio_disponivel(navegador, usuario_busca="automac", segundos_tolerancia=300):
        linhas_raw = navegador.find_elements(By.XPATH, "//tr[@class='srtr2']")
        candidato = None

        for linha in linhas_raw:
            try:
                nome_relatorio = linha.find_element(By.XPATH, ".//td[2]/div").text
                usuario = linha.find_element(By.XPATH, ".//td[4]/div").text
                status = linha.find_element(By.XPATH, ".//td[7]/div").text
                data_str = linha.find_element(By.XPATH, ".//td[3]/div").text
                data_linha = datetime.strptime(data_str, "%d/%m/%y %H:%M:%S")

                if "200" in nome_relatorio and usuario_busca in usuario and status != "Concluído":
                    candidato = {
                        "data": data_str,
                        "linha": linha,
                        "data_obj": data_linha
                    }
                    break
            except:
                continue

        if not candidato:
            return

        inicio_espera = time.time()
        while time.time() - inicio_espera < segundos_tolerancia:
            try:
                navegador.find_element(By.LINK_TEXT, "Atualizar").click()
                time.sleep(7)
                linhas_raw = navegador.find_elements(By.XPATH, "//tr[@class='srtr2']")
                for linha in linhas_raw:
                    try:
                        data_str = linha.find_element(By.XPATH, ".//td[3]/div").text
                        status = linha.find_element(By.XPATH, ".//td[7]/div").text
                        usuario = linha.find_element(By.XPATH, ".//td[4]/div").text
                        nome_relatorio = linha.find_element(By.XPATH, ".//td[2]/div").text

                        if (data_str == candidato["data"]
                            and "200" in nome_relatorio
                            and usuario_busca in usuario
                            and "Concluído" in status):

                            botao_baixar = linha.find_element(By.XPATH, ".//td[9]//a[@class='sra']")
                            navegador.execute_script("arguments[0].scrollIntoView(true);", botao_baixar)
                            time.sleep(2)
                            botao_baixar.click()
                            return
                    except:
                        continue
            except:
                pass
            time.sleep(5)

    # === Início ===
    pasta_downloads = os.path.join(os.path.expanduser("~"), "Downloads")
    tempo_inicio = time.time()

    # Acessa a tela de login
    navegador.get("https://sistema.ssw.inf.br/bin/ssw0422")
    time.sleep(3)

    # Login
    campos_login = navegador.find_elements(By.XPATH, "/html/body/form/input")

    if len(campos_login) >= 4:
        campos_login[0].clear()
        campos_login[0].send_keys(empresa)

        campos_login[1].clear()
        campos_login[1].send_keys(cnpj)

        campos_login[2].clear()
        campos_login[2].send_keys(usuario_ssw + Keys.TAB)  # <-- aqui o TAB garante que o foco vai para o próximo campo

        campos_login[3].clear()
        campos_login[3].send_keys(senha_ssw)

        navegador.find_element(By.XPATH, "/html/body/form/a").click()
    else:
        print("Erro: campos de login não encontrados.")
        return

    inicio = (datetime.today().replace(day=1) - relativedelta(months=meses_retroativos))
    fim = datetime.today()
    periodos = gerar_periodos(inicio, fim)

    for data_inicio, data_fim, nome_mes in periodos:
        print(f"\n>> Iniciando mês: {nome_mes} ({data_inicio} a {data_fim})")

        # Fecha janelas anteriores
        principal = navegador.window_handles[0]
        for h in navegador.window_handles:
            if h != principal:
                navegador.switch_to.window(h)
                navegador.close()
        navegador.switch_to.window(principal)

        # Abre relatório 200
        time.sleep(5)
        # campo_relatorio = navegador.find_element(By.XPATH, "/html/body/form/input[2]")
        campo_relatorio = WebDriverWait(navegador, 15).until(
            EC.visibility_of_element_located((By.XPATH, "/html/body/form/input[2]"))
        )
        campo_relatorio.clear()
        campo_relatorio.send_keys("200")
        time.sleep(3)
        WebDriverWait(navegador, 10).until(lambda d: len(d.window_handles) > 1)
        navegador.switch_to.window(navegador.window_handles[-1])
        WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.ID, "1")))

        navegador.find_element(By.ID, "1").clear()
        navegador.find_element(By.ID, "1").send_keys(data_inicio, data_fim)
        navegador.find_element(By.ID, "11").clear()
        navegador.find_element(By.ID, "11").send_keys("E")
        time.sleep(5)

        # Entra na fila (se existir)
        WebDriverWait(navegador, 10).until(lambda d: len(d.window_handles) > 1)
        navegador.switch_to.window(navegador.window_handles[-1])
        time.sleep(5)

        try:
            navegador.find_element(By.XPATH, "/html/body/div[3]/a[1]").click()
            WebDriverWait(navegador, 10).until(lambda d: len(d.window_handles) > 1)
            navegador.switch_to.window(navegador.window_handles[-1])
            navegador.find_element(By.LINK_TEXT, "Atualizar").click()
            time.sleep(7)
            esperar_relatorio_disponivel(navegador, usuario_busca=usuario_ssw)
        except NoSuchElementException:
            print("Relatório baixado automaticamente, sem fila.")

        # Aguarda o download
        arquivo_baixado = None
        inicio_espera = time.time()
        while time.time() - inicio_espera < 60:
            arquivos = os.listdir(pasta_downloads)
            arquivos_csv = [f for f in arquivos if f.lower().endswith(".csv") and not f.endswith(".crdownload")]
            if arquivos_csv:
                arquivos_csv.sort(key=lambda f: os.path.getmtime(os.path.join(pasta_downloads, f)), reverse=True)
                arquivo_baixado = os.path.join(pasta_downloads, arquivos_csv[0])
                break
            time.sleep(1)

        if arquivo_baixado:
            mes_ano_str = datetime.strptime(data_inicio, "%d%m%y").strftime("%m%Y")
            novo_nome = f"mani-{mes_ano_str}.csv"
            novo_arquivo = os.path.join(pasta_downloads, novo_nome)
            os.rename(arquivo_baixado, novo_arquivo)

            destino = os.path.join(pasta_destino, novo_nome)
            if os.path.exists(destino):
                os.remove(destino)
            shutil.move(novo_arquivo, destino)
            print(f"Arquivo {novo_nome} movido com sucesso!")
        else:
            print(f"Arquivo não encontrado para o mês {nome_mes}.")

    duracao = time.time() - tempo_inicio
    minutos, segundos = divmod(duracao, 60)
    print(f"Manifestos atualizado com sucesso. \nTempo total de execução: {int(minutos)}min {int(segundos)}s")
if __name__ == "__main__":
    from selenium import webdriver



    # Aqui chamamos a função para rodar e esta certinho
    # Aqui você cria o navegador (exemplo com Chrome)
    navegador = webdriver.Chrome()  
    navegador.maximize_window()


    # Chama a função passando o navegador
    baixar_relatorios_200_em_lote(navegador)

    # Fecha o navegador no final
    navegador.quit()
