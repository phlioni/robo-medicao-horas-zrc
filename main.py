# -*- coding: utf-8 -*-
import os
import shutil
from datetime import datetime
import time

import config
import automacao_web
import processamento_dados
import envio_email
import excel_handler

def arquivar_relatorio(caminho_arquivo):
    """Renomeia e move o arquivo principal baixado para a pasta de relatórios."""
    print("Arquivando o relatório principal...")
    if not os.path.exists(config.PASTA_RELATORIOS_FINAL):
        os.makedirs(config.PASTA_RELATORIOS_FINAL)
        print(f"Pasta '{config.PASTA_RELATORIOS_FINAL}' criada.")
        
    timestamp = datetime.now().strftime("%d-%m-%Y_%H-%M-%S")
    _, extensao = os.path.splitext(caminho_arquivo)
    novo_nome = f"relatorio-horas-ZRC_{timestamp}{extensao}"
    
    novo_caminho = os.path.join(config.PASTA_RELATORIOS_FINAL, novo_nome)
    
    shutil.move(caminho_arquivo, novo_caminho)
    print(f"Relatório arquivado como: '{novo_caminho}'")

def run():
    """Função principal que executa o processo para o projeto ZRC."""
    timing_report = {}
    status = "SUCESSO"
    error_message = ""
    caminho_arquivo_processado = None
    
    start_total_time = time.time()

    try:
        start_step_time = time.time()
        caminho_arquivo_processado, periodo_relatorio = automacao_web.login_e_download()
        timing_report["1. Download do Relatório"] = time.time() - start_step_time
        
        start_step_time = time.time()
        excel_handler.unprotect_and_save(caminho_arquivo_processado)
        timing_report["2. Desbloqueio do Arquivo Excel"] = time.time() - start_step_time
        
        start_step_time = time.time()
        df_processado, mes_ano_relatorio = processamento_dados.processar_planilha(caminho_arquivo_processado)
        timing_report["3. Processamento (Filtro ZRC)"] = time.time() - start_step_time
        
        if df_processado is not None and not df_processado.empty:
            start_step_time = time.time()
            relatorios_por_cliente = processamento_dados.gerar_planilhas_por_projeto(df_processado, mes_ano_relatorio)
            timing_report["4. Geração de Planilha Detalhada"] = time.time() - start_step_time
            
            start_step_time = time.time()
            tabela_html_zrc = processamento_dados.criar_tabela_html(df_processado)
            
            # Pega a lista de arquivos para o cliente 'ZRC'
            arquivos_para_anexar = relatorios_por_cliente.get('ZRC', [])
            
            envio_email.enviar_email_zrc(tabela_html_zrc, periodo_relatorio, arquivos_para_anexar)
            timing_report["5. Envio de E-mail ZRC"] = time.time() - start_step_time

            start_step_time = time.time()
            arquivar_relatorio(caminho_arquivo_processado)
            timing_report["6. Arquivamento do Relatório"] = time.time() - start_step_time
        else:
            print("Nenhum dado para processar para o projeto ZRC após a filtragem.")
            if caminho_arquivo_processado and os.path.exists(caminho_arquivo_processado):
                os.remove(caminho_arquivo_processado)

    except Exception as e:
        status = "FALHA"
        error_message = str(e)
        print(f"\nOcorreu um erro crítico durante a execução: {e}")
        if caminho_arquivo_processado and os.path.exists(caminho_arquivo_processado):
            os.remove(caminho_arquivo_processado)
    finally:
        timing_report["Tempo Total de Execução"] = time.time() - start_total_time
        print("\nProcesso finalizado.")
        envio_email.enviar_email_status(timing_report, status, error_message)

if __name__ == "__main__":
    run()