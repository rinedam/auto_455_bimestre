from selenium import webdriver
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime, timedelta
import os
import time
import schedule
import json

# Caminho para a pasta de downloads desejada
download_folder = os.path.expanduser('I:\\.shortcut-targets-by-id\\1BbEijfOOPBwgJuz8LJhqn9OtOIAaEdeO\\Logdi\\Relatório e Dashboards\\teste_auto_455')

# Configurações do Edge
edge_options = Options()
edge_options.add_experimental_option('prefs', {
    "download.default_directory": download_folder,  # Define o diretório de download
    "download.prompt_for_download": False,           # Não solicitar confirmação para download
    "download.directory_upgrade": True,               # Permitir a atualização do diretório
    "safebrowsing.enabled": True                       # Habilitar navegação segura
})

# Dicionário para armazenar o estado das execuções
execucoes = {}

# Função para carregar o estado das execuções de um arquivo JSON
def carregar_execucoes():
    try:
        with open("execucoes.json", "r") as arquivo:
            return json.load(arquivo)
    except FileNotFoundError:
        return {}

# Função para salvar o estado das execuções em um arquivo JSON
def salvar_execucoes(execucoes):
    with open("execucoes.json", "w") as arquivo:
        json.dump(execucoes, arquivo)

# Carregar o estado das execuções ao iniciar o script
execucoes = carregar_execucoes()

# Função para calcular o primeiro e o último dia do mês
def calcular_datas_mes(ano, mes):
    inicio = datetime(ano, mes, 1)
    if mes == 12:
        fim = datetime(ano, mes, 31)
    else:
        fim = datetime(ano, mes + 1, 1) - timedelta(days=1)
    return inicio.strftime('%d%m%y'), fim.strftime('%d%m%y')

# Função para determinar o bimestre atual, anterior e o penúltimo
def determinar_bimestres():
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

    # Determina o bimestre anterior
    if bimestre_atual == 1:
        bimestre_anterior = 6
        ano_anterior = ano_atual - 1
    else:
        bimestre_anterior = bimestre_atual - 1
        ano_anterior = ano_atual

    # Determina o penúltimo bimestre
    if bimestre_anterior == 1:
        penultimo_bimestre = 6
        ano_penultimo = ano_anterior - 1
    else:
        penultimo_bimestre = bimestre_anterior - 1
        ano_penultimo = ano_anterior

    return bimestre_atual, bimestre_anterior, penultimo_bimestre, ano_atual, ano_anterior, ano_penultimo

# Função para verificar se é dia útil
def eh_dia_util():
    hoje = datetime.now()
    # Verifica se é dia útil (segunda a sexta)
    return hoje.weekday() < 5  # 0 = segunda, 4 = sexta

# Função para verificar se o relatório do penúltimo bimestre já foi extraído na semana
def ja_executado_na_semana(mes, ano):
    chave = f"{mes}/{ano}"
    if chave in execucoes:
        ultima_execucao = datetime.strptime(execucoes[chave], '%Y-%m-%d %H:%M:%S')
        hoje = datetime.now()
        # Verifica se a última execução foi na mesma semana
        if (hoje - ultima_execucao).days < 7:
            return True
    return False

# Função para atualizar o estado de execução
def atualizar_execucao(mes, ano):
    chave = f"{mes}/{ano}"
    execucoes[chave] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    salvar_execucoes(execucoes)

# Função para realizar o login
def realizar_login(driver):
    driver.get("https://sistema.ssw.inf.br/bin/ssw0422")
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, "f1")))
    driver.find_element(By.NAME, "f1").send_keys("LDI")
    driver.find_element(By.NAME, "f2").send_keys("12373493977")
    driver.find_element(By.NAME, "f3").send_keys("gustavo")
    driver.find_element(By.NAME, "f4").send_keys("12032006")
    login_button = driver.find_element(By.ID, "5")
    driver.execute_script("arguments[0].click();", login_button)
    time.sleep(5)

# Função para preencher o formulário e extrair o relatório
def preencher_formulario(driver, ano, mes):
    inicio_mes, fim_mes = calcular_datas_mes(ano, mes)
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, "f2")))
    driver.find_element(By.NAME, "f3").send_keys("455")
    time.sleep(7)
    abas = driver.window_handles
    driver.switch_to.window(abas[-1])
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, "f12")))
    input_element = driver.find_element(By.ID, "11")
    driver.execute_script("arguments[0].value = '';", input_element)
    driver.find_element(By.ID, "11").send_keys(inicio_mes)
    time.sleep(0.3)
    input_element = driver.find_element(By.ID, "12")
    driver.execute_script("arguments[0].value = '';", input_element)
    driver.find_element(By.ID, "12").send_keys(fim_mes)
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
    login_button = driver.find_element(By.ID, "40")
    driver.execute_script("arguments[0].click();", login_button)
    time.sleep(0.8)
    actions = ActionChains(driver)
    actions.send_keys("1").perform()
    time.sleep(5)
    abas = driver.window_handles
    driver.switch_to.window(abas[-1])
    time.sleep(2)

def capturar_seq(driver):
    # Captura o valor do seq da segunda linha da primeira coluna da tabela com o ID 'tblsr'
    tabela = driver.find_element(By.ID, "tblsr")
    linhas = tabela.find_elements(By.TAG_NAME, "tr")

    # Verifica se há pelo menos duas linhas (cabeçalho + dados)
    if len(linhas) > 1:
        seq_da_requisicao = linhas[1].find_element(By.TAG_NAME, "td").text  # Captura o valor da primeira coluna da segunda linha
        print(f"Seq da requisição: {seq_da_requisicao}")
        return seq_da_requisicao
    else:
        print("Não há linhas suficientes na tabela para capturar o seq.")
        return None

def atualizar_relatorio(driver, seq_da_requisicao):
    # Clica no botão de atualização
    update_button = driver.find_element(By.ID, "2")
    time.sleep(60)
    driver.execute_script("arguments[0].click();", update_button)
    time.sleep(5)

    # Após a atualização, busca novamente a linha correspondente ao seq
    relatorios_atualizados = driver.find_elements(By.CSS_SELECTOR, "table#tblsr tr")
    relatorio_encontrado = None

    for relatorio in relatorios_atualizados[1:]:  # Ignora o cabeçalho da tabela
        seq_atual = relatorio.find_element(By.TAG_NAME, "td").text  # Captura o seq da primeira coluna
        if seq_atual == seq_da_requisicao:
            relatorio_encontrado = relatorio
            break

    # Se um relatório correspondente foi encontrado, clique no link (substituído por <u>) na mesma linha
    if relatorio_encontrado:
        link = relatorio_encontrado.find_element(By.TAG_NAME, "u")  # Supondo que o link esteja dentro de uma tag <u>
        driver.execute_script("arguments[0].click();", link)
        print("Clicou no link da requisição correspondente.")
    else:
        print("Nenhum relatório encontrado após a atualização.")

# Função para executar o processo de extração
def executar_processo(bimestre, ano, mes, penultimo_bimestre):
    if not eh_dia_util():
        print(f"Ignorando execução: hoje não é dia útil.")
        return

    # Verifica se é o penúltimo bimestre e se já foi executado na semana
    if bimestre == penultimo_bimestre and ja_executado_na_semana(mes, ano):
        print(f"Relatório do mês {mes}/{ano} (penúltimo bimestre) já foi extraído esta semana.")
        return

    print(f"Iniciando extração do relatório para o mês {mes}/{ano} do bimestre {bimestre}...")
    driver = webdriver.Edge(options=edge_options)
    try:
        realizar_login(driver)
        preencher_formulario(driver, ano, mes)
        seq_da_requisicao = capturar_seq(driver)
        if seq_da_requisicao:
            atualizar_relatorio(driver, seq_da_requisicao)
            time.sleep(10)  # Atraso para garantir que o download seja iniciado

        print(f"Relatório extraído com sucesso para o mês {mes}/{ano} do bimestre {bimestre}.")
        # Atualiza o estado de execução apenas para o penúltimo bimestre
        if bimestre == penultimo_bimestre:
            atualizar_execucao(mes, ano)
    except Exception as e:
        print(f"Erro ao extrair relatório: {e}")
    finally:
        driver.quit()

# Função para gerar horários únicos entre 09:00 e 18:00
def gerar_horarios_unicos(quantidade):
    # Horários fixos para o penúltimo bimestre
    if quantidade == 2:
        return ["10:00", "10:30"]  # Horários específicos para o penúltimo bimestre
    else:
        horarios = set()
        while len(horarios) < quantidade:
            hora = f"{9 + len(horarios):02d}:00"  # Garante que a hora tenha dois dígitos (ex: 09:00)
            horarios.add(hora)
        return sorted(horarios)

# Função para agendar as tarefas para a semana atual
def agendar_tarefas_semana():
    hoje = datetime.now()
    mes_atual = hoje.month
    ano_atual = hoje.year

    # Determina os bimestres
    bimestre_atual, bimestre_anterior, penultimo_bimestre, ano_atual, ano_anterior, ano_penultimo = determinar_bimestres()

    print(f"Agendando tarefas para a semana atual...")

    # Lista de meses para cada bimestre
    meses_bimestre_atual = list(range((bimestre_atual - 1) * 2 + 1, (bimestre_atual - 1) * 2 + 3))
    meses_bimestre_anterior = list(range((bimestre_anterior - 1) * 2 + 1, (bimestre_anterior - 1) * 2 + 3))
    meses_penultimo_bimestre = list(range((penultimo_bimestre - 1) * 2 + 1, (penultimo_bimestre - 1) * 2 + 3))

    # Horários fixos para o bimestre atual (cada mês com horários diferentes)
    horarios_bimestre_atual = {
        meses_bimestre_atual[0]: ["09:00", "12:00", "14:00", "17:00"],  # Mês 1
        meses_bimestre_atual[1]: ["09:10", "12:10", "14:10", "17:10"],  # Mês 2
    }

    # Horários fixos para o penúltimo bimestre (10:00 e 10:30)
    horarios_penultimo_bimestre = ["10:00", "10:30"]

    # Agenda as tarefas para cada dia útil da semana
    for dia in range(hoje.day, hoje.day + 7):  # Agenda para os próximos 7 dias
        try:
            data = datetime(ano_atual, mes_atual, dia)
            if data.weekday() < 5:  # Apenas dias úteis (segunda a sexta)
                # Agenda a extração para o bimestre atual
                for mes in meses_bimestre_atual:
                    for horario in horarios_bimestre_atual[mes]:
                        schedule.every().day.at(horario).do(executar_processo, bimestre_atual, ano_atual, mes, penultimo_bimestre).tag("bimestre_atual")
                        print(f"Agendando extração para o mês {mes}/{ano_atual} (bimestre atual) às {horario}")

                # Agenda a extração para o bimestre anterior
                for mes in meses_bimestre_anterior:
                    horario = f"09:{20 + meses_bimestre_anterior.index(mes) * 10}"  # Horários dinâmicos
                    schedule.every().day.at(horario).do(executar_processo, bimestre_anterior, ano_anterior, mes, penultimo_bimestre).tag("bimestre_anterior")
                    print(f"Agendando extração para o mês {mes}/{ano_anterior} (bimestre anterior) às {horario}")

                # Agenda a extração para o penúltimo bimestre (1 vez por semana)
                if data.weekday() == 0:  # Segunda-feira
                    for i, mes in enumerate(meses_penultimo_bimestre):
                        # Atribui horários específicos para cada mês do penúltimo bimestre
                        horario = horarios_penultimo_bimestre[i]
                        schedule.every().monday.at(horario).do(executar_processo, penultimo_bimestre, ano_penultimo, mes, penultimo_bimestre).tag("penultimo_bimestre")
                        print(f"Agendando extração para o mês {mes}/{ano_penultimo} (penúltimo bimestre) às {horario} (1 vez por semana)")
        except ValueError:
            # Ignora dias que não existem no mês (ex: 31 de fevereiro)
            pass

# Loop principal
while True:
    print("Iniciando agendamento para a semana atual...")
    agendar_tarefas_semana()

    # Executa as tarefas agendadas durante a semana
    while datetime.now().weekday() < 5:  # Enquanto for dia útil (segunda a sexta)
        schedule.run_pending()
        time.sleep(1)

    print("Semana concluída. Aguardando próxima semana...")
    time.sleep(86400)  # Aguarda 1 dia antes de reiniciar o agendamento