# -*- coding: utf-8 -*-
import os
import time
import shutil
from datetime import datetime
from dateutil.relativedelta import relativedelta
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

import config

def login_e_download():
    """Orquestra a automação web, com período de datas customizado."""
    options = webdriver.ChromeOptions()
    # options.add_argument("--headless")
    options.add_argument("--start-maximized")
    driver = webdriver.Chrome(options=options)
    
    periodo_relatorio_str = ""
    
    try:
        wait = WebDriverWait(driver, 20)
        print("Acessando o site...")
        driver.get(config.SITE_URL)

        try:
            print("Procurando por campos de login...")
            username_field = wait.until(EC.presence_of_element_located((By.NAME, "LoginName")))
            password_field = driver.find_element(By.NAME, "Password")
            print("Realizando login...")
            username_field.send_keys(config.SITE_LOGIN)
            password_field.send_keys(config.SITE_SENHA)
            driver.find_element(By.ID, "button_processLogin").click()
        except TimeoutException:
            print("Campos de login não encontrados. Presumindo que já está logado.")
        
        print("Navegando para 'Resumo de Horas por Profissional'...")
        wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Resumo de Horas por Profissional"))).click()

        print("Aguardando a página de filtros carregar...")
        wait.until(EC.number_of_windows_to_be(2))
        driver.switch_to.window(driver.window_handles[1])

        print("Calculando o período de datas (15/mês anterior a 14/mês atual)...")
        hoje = datetime.now()
        data_fim = hoje.replace(day=14)
        data_inicio = (hoje - relativedelta(months=1)).replace(day=15)
        
        data_inicio_str = data_inicio.strftime("%d/%m/%Y")
        data_fim_str = data_fim.strftime("%d/%m/%Y")
        periodo_relatorio_str = f"{data_inicio_str} a {data_fim_str}"
        print(f"Período selecionado: {periodo_relatorio_str}")

        print("Aplicando filtro de data customizado...")
        wait.until(EC.element_to_be_clickable((By.ID, "P_DATA_div"))).click()
        
        campo_data_inicio = wait.until(EC.visibility_of_element_located((By.NAME, "daterangepicker_start")))
        campo_data_inicio.clear()
        campo_data_inicio.send_keys(data_inicio_str)
        
        campo_data_fim = driver.find_element(By.NAME, "daterangepicker_end")
        campo_data_fim.clear()
        campo_data_fim.send_keys(data_fim_str)
        
        driver.find_element(By.CLASS_NAME, "applyBtn").click()
        time.sleep(1)

        print("Executando relatório...")
        driver.find_element(By.ID, "button_Execute").click()

        print("Aguardando o carregamento do relatório...")
        time.sleep(10)
        
        arquivos_antes = set(os.listdir(config.PASTA_DOWNLOADS))
        
        print("Exportando para Excel...")
        driver.find_element(By.ID, "button_ExecuteXSL").click()
        print("Aguardando 10 segundos para o início do download...")
        time.sleep(10)

        print(f"Monitorando a pasta {config.PASTA_DOWNLOADS} por um novo arquivo...")
        start_time = time.time()
        caminho_arquivo_baixado = None
        
        while time.time() - start_time < 60:
            arquivos_depois = set(os.listdir(config.PASTA_DOWNLOADS))
            novos_arquivos = arquivos_depois - arquivos_antes
            
            if novos_arquivos and not list(novos_arquivos)[0].endswith('.crdownload'):
                nome_do_arquivo = novos_arquivos.pop()
                caminho_arquivo_baixado = os.path.join(config.PASTA_DOWNLOADS, nome_do_arquivo)
                print(f"Novo arquivo detectado: {nome_do_arquivo}")
                break
            time.sleep(1)

        if not caminho_arquivo_baixado:
            raise TimeoutException("O download do arquivo demorou mais de 60 segundos ou não foi detectado.")
        
        time.sleep(2)
        print("Download concluído.")
        
        caminho_destino = os.path.join(os.getcwd(), os.path.basename(caminho_arquivo_baixado))
        shutil.move(caminho_arquivo_baixado, caminho_destino)
        print(f"Arquivo movido para a pasta do projeto: {caminho_destino}")
        
        return caminho_destino, periodo_relatorio_str

    finally:
        if driver:
            for handle in driver.window_handles:
                driver.switch_to.window(handle)
                driver.close()