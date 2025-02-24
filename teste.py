from selenium import webdriver
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime, timedelta
import pandas as pd
import os
import time
import schedule

# Importa as funções do arquivo alimentacao_planilha
# from subdiretorio.alimentacao_planilha import encontrar_ultimo_arquivo_swwweb, processar_arquivo_swwweb

# Caminho para a pasta de downloads desejada
download_folder = os.path.expanduser('G:\\.shortcut-targets-by-id\\1BbEijfOOPBwgJuz8LJhqn9OtOIAaEdeO\\Logdi\\Relatório e Dashboards\\teste_auto_455')

# Caminho para a planilha de destino
planilha_destino = os.path.join(download_folder, "dados_copiados.xlsx")

# Configurações do Edge
edge_options = Options()
edge_options.add_experimental_option('prefs', {
    "download.default_directory": download_folder,  # Define o diretório de download
    "download.prompt_for_download": False,           # Não solicitar confirmação para download
    "download.directory_upgrade": True,               # Permitir a atualização do diretório
    "safebrowsing.enabled": True                       # Habilitar navegação segura
})

# Função para calcular o primeiro e o último dia do bimestre
def calcular_datas_bimestre(bimestre, ano):
    if bimestre == 1:
        inicio = datetime(ano, 1, 1)
        fim = datetime(ano, 2, 28) if not (ano % 4 == 0 and (ano % 100 != 0 or ano % 400 == 0)) else datetime(ano, 2, 29)
    elif bimestre == 2:
        inicio = datetime(ano, 3, 1)
        fim = datetime(ano, 4, 30)
    elif bimestre == 3:
        inicio = datetime(ano, 5, 1)
        fim = datetime(ano, 6, 30)
    elif bimestre == 4:
        inicio = datetime(ano, 7, 1)
        fim = datetime(ano, 8, 31)
    elif bimestre == 5:
        inicio = datetime(ano, 9, 1)
        fim = datetime(ano, 10, 31)
    elif bimestre == 6:
        inicio = datetime(ano, 11, 1)
        fim = datetime(ano, 12, 31)
    return inicio.strftime('%d%m%y'), fim.strftime('%d%m%y')

def realizar_login(driver):
    # Navega até a página do formulário
    driver.get("https://sistema.ssw.inf.br/bin/ssw0422")  # Substitua pela URL do seu formulário

    # Atraso para garantir que a página carregue completamente
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, "f1")))

    # Preenche os campos de login
    driver.find_element(By.NAME, "f1").send_keys("LDI")
    driver.find_element(By.NAME, "f2").send_keys("12373493977")  # Exemplo de CPF
    driver.find_element(By.NAME, "f3").send_keys("gustavo")
    driver.find_element(By.NAME, "f4").send_keys("12032006")

    # Clica no botão de login diretamente
    login_button = driver.find_element(By.ID, "5")
    driver.execute_script("arguments[0].click();", login_button)
    time.sleep(5)

def preencher_formulario(driver, bimestre, ano):
    # Calcula as datas do bimestre
    inicio_bimestre, fim_bimestre = calcular_datas_bimestre(bimestre, ano)

    # Preenche os campos de Unidade e Opção
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, "f2")))
    driver.find_element(By.NAME, "f3").send_keys("455")
    time.sleep(7)  # Atraso para carregar a nova aba

    abas = driver.window_handles  # Lista o número de abas abertas.
    driver.switch_to.window(abas[-1])  # Muda o foco para a última aba (a nova aba)

    # Preenche o campo com a data final do bimestre
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, "f12")))
    input_element = driver.find_element(By.ID, "11")
    driver.execute_script("arguments[0].value = '';", input_element)
    time.sleep(0.3)
    driver.find_element(By.ID, "11").send_keys(inicio_bimestre) # Data ínicio do bimestre
    time.sleep(0.3)
    input_element = driver.find_element(By.ID, "12")
    driver.execute_script("arguments[0].value = '';", input_element)
    driver.find_element(By.ID, "12").send_keys(fim_bimestre)  # Data final do bimestre
    time.sleep(0.3)
    driver.find_element(By.NAME, "f21").clear()
    time.sleep(0.3)
    driver.find_element(By.NAME, "f21").send_keys("t")
    time.sleep(0.3)
    driver.find_element(By.NAME, "f35").clear()
    time.sleep(0.3)
    driver.find_element(By.NAME, "f35").send_keys("e")
    time.sleep(0.3)
    input_element = driver.find_element(By.NAME, "f37")
    driver.execute_script("arguments[0].value = '';", input_element)
    driver.find_element(By.NAME, "f37").send_keys("b")
    time.sleep(0.3)

    # Clica no botão de play
    login_button = driver.find_element(By.ID, "40")
    driver.execute_script("arguments[0].click();", login_button)
    time.sleep(0.8)

    # Usando ActionChains para enviar a tecla "1"
    actions = ActionChains(driver)
    actions.send_keys("1").perform()

    time.sleep(5)
    abas = driver.window_handles  # Lista o número de abas abertas.

    # Muda o foco para a última aba (a nova aba)
    driver.switch_to.window(abas[-1])
    time.sleep(2)

def main():
    # Inicializa o WebDriver
    driver = webdriver.Edge(options=edge_options)

    try:
        # Realiza o login
        realizar_login(driver)

        # Determina o bimestre atual e o ano
        hoje = datetime.now()
        mes_atual = hoje.month
        ano_atual = hoje.year

        # Determina o bimestre atual
        if mes_atual in [1, 2]:
            bimestre_atual = 1
        elif mes_atual in [3, 4]:
            bimestre_atual = 2
        elif mes_atual in [5, 6]:
            bimestre_atual = 3
        elif mes_atual in [7, 8]:
            bimestre_atual = 4
        elif mes_atual in [9, 10]:
            bimestre_atual = 5
        else:
            bimestre_atual = 6

        # Preenche o formulário com as datas do bimestre atual
        preencher_formulario(driver, bimestre_atual, ano_atual)

        print("Processo concluído com sucesso!")

    except Exception as e:
        print(f"Ocorreu um erro durante a execução: {e}")

    finally:
        # Fecha o navegador
        driver.quit()

# Executa a função principal para testar o código
if __name__ == "__main__":
    main()

