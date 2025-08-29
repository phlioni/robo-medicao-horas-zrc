# -*- coding: utf-8 -*-
import pandas as pd
import os
import re
from openpyxl.styles import Font

def processar_planilha(caminho_arquivo):
    """
    Lê o relatório, extrai o mês/ano, filtra para o projeto ZRC e prepara os dados.
    """
    print("Processando a planilha...")
    mes_ano_formatado = "do período"
    try:
        df = pd.read_excel(caminho_arquivo, skiprows=18, engine='openpyxl')
        df.columns = df.columns.str.strip()

        df['Data'] = pd.to_datetime(df['Data'], format='%d/%m/%Y', errors='coerce')
        df.dropna(subset=['Data'], inplace=True)

        # Adicionada a extração de mês/ano para nomear arquivos/pastas
        if not df.empty:
            primeira_data_valida = df['Data'].iloc[0]
            mes_num = str(primeira_data_valida.month).zfill(2)
            ano = str(primeira_data_valida.year)
            meses_pt = {"01":"Janeiro", "02":"Fevereiro", "03":"Março", "04":"Abril", "05":"Maio", "06":"Junho", "07":"Julho", "08":"Agosto", "09":"Setembro", "10":"Outubro", "11":"Novembro", "12":"Dezembro"}
            mes_nome = meses_pt.get(mes_num, "")
            mes_ano_formatado = f"<b>{mes_nome}</b>/<b>{ano}</b>"
            print(f"Mês de referência encontrado na coluna 'Data': {mes_nome}/{ano}")

        df_filtrado = df[df['Projeto'].str.startswith('ZRC', na=False)].copy()

        if df_filtrado.empty:
            print("AVISO: Nenhum lançamento para o projeto ZRC encontrado no relatório.")
            return None, mes_ano_formatado

        df_filtrado = df_filtrado[df_filtrado['Situação'] == 'Aprovado'].copy()
        df_filtrado['Comentários'] = df_filtrado['Comentários'].astype(str)
        df_filtrado = df_filtrado[df_filtrado['Comentários'].str.len() >= 5]
        
        print("Planilha do projeto ZRC processada com sucesso.")
        return df_filtrado, mes_ano_formatado

    except Exception as e:
        print(f"Ocorreu um erro inesperado ao processar a planilha: {e}")
        return None, mes_ano_formatado

def criar_tabela_html(df_zrc):
    """Cria a tabela HTML detalhada (com profissionais) para o projeto ZRC."""
    # (Esta função permanece a mesma)
    grouped = df_zrc.groupby(['Projeto', 'Profissional'])['Horas'].sum().reset_index()
    html_output = ""
    total_geral = 0
    projetos = grouped['Projeto'].unique()
    for projeto in projetos:
        df_projeto = grouped[grouped['Projeto'] == projeto]
        total_projeto = df_projeto['Horas'].sum()
        total_geral += total_projeto
        total_projeto_str = f"{total_projeto:,.2f}".replace(',', '#').replace('.', ',').replace('#', '.')
        html_output += f"""<tr style="background-color: #DDEBF7; font-weight: bold;"><td style="padding: 8px; border: 1px solid #dddddd;">{projeto}</td><td style="padding: 8px; border: 1px solid #dddddd; text-align: right;">{total_projeto_str}</td></tr>"""
        for _, row in df_projeto.iterrows():
            horas_str = f"{row['Horas']:,.2f}".replace(',', '#').replace('.', ',').replace('#', '.')
            html_output += f"""<tr><td style="padding: 8px; border: 1px solid #dddddd; padding-left: 25px;">{row['Profissional']}</td><td style="padding: 8px; border: 1px solid #dddddd; text-align: right;">{horas_str}</td></tr>"""
    total_geral_str = f"{total_geral:,.2f}".replace(',', '#').replace('.', ',').replace('#', '.')
    html_final = f"""<table style="width: 600px; border-collapse: collapse; font-family: Calibri, sans-serif; font-size: 11pt;"><thead style="background-color: #4472C4; color: white;"><tr><th style="padding: 8px; border: 1px solid #dddddd; text-align: left;">Rótulos de Linha</th><th style="padding: 8px; border: 1px solid #dddddd; text-align: right;">Soma de Horas</th></tr></thead><tbody>{html_output}</tbody><tfoot><tr style="background-color: #DDEBF7; font-weight: bold;"><td style="padding: 8px; border: 1px solid #dddddd;">Total Geral</td><td style="padding: 8px; border: 1px solid #dddddd; text-align: right;">{total_geral_str}</td></tr></tfoot></table>"""
    return html_final

# --- (FUNÇÃO ADICIONADA DO OUTRO ROBÔ) ---
def gerar_planilhas_por_projeto(df_completo, mes_ano_relatorio):
    """
    Cria um arquivo Excel para cada projeto, contendo abas para cada profissional
    com uma linha de total de horas no topo.
    """
    print("\nIniciando geração de planilhas individuais por projeto...")
    
    mes_ano_pasta = re.sub('<[^<]+?>', '', mes_ano_relatorio).replace('/', '-')
    
    caminho_saida = os.path.join("relatorios", "medicao", mes_ano_pasta)
    os.makedirs(caminho_saida, exist_ok=True)
    print(f"Salvando relatórios individuais em: {caminho_saida}")

    projetos_unicos = df_completo['Projeto'].unique()
    relatorios_por_cliente = {}

    for projeto in projetos_unicos:
        df_projeto = df_completo[df_completo['Projeto'] == projeto]
        profissionais_unicos = df_projeto['Profissional'].unique()
        
        nome_arquivo = f"{projeto}-relatorio-horas-{mes_ano_pasta}.xlsx"
        caminho_completo_arquivo = os.path.join(caminho_saida, nome_arquivo)
        
        # O código do cliente para ZRC será 'ZRC'
        cod_cliente = projeto[:3]
        if cod_cliente not in relatorios_por_cliente:
            relatorios_por_cliente[cod_cliente] = []
        relatorios_por_cliente[cod_cliente].append(caminho_completo_arquivo)
        
        with pd.ExcelWriter(caminho_completo_arquivo, engine='openpyxl') as writer:
            for profissional in profissionais_unicos:
                df_aba = df_projeto[df_projeto['Profissional'] == profissional]
                
                total_horas_profissional = df_aba['Horas'].sum()

                df_aba_final = df_aba[['Profissional', 'Projeto', 'Data', 'Horas', 'Comentários']].copy()
                df_aba_final['Data'] = df_aba_final['Data'].dt.strftime('%d/%m/%Y')
                df_aba_final['Horas'] = df_aba_final['Horas'].apply(lambda x: str(f'{x:.2f}').replace('.', ','))
                
                nome_aba = profissional[:31]
                
                df_aba_final.to_excel(writer, sheet_name=nome_aba, index=False, startrow=1)
                
                worksheet = writer.sheets[nome_aba]
                
                worksheet['A1'] = 'Total de Horas'
                worksheet['D1'] = total_horas_profissional
                
                bold_font = Font(bold=True)
                worksheet['A1'].font = bold_font
                worksheet['D1'].font = bold_font
                worksheet['D1'].number_format = '#,##0.00'

        print(f" -> Arquivo '{nome_arquivo}' gerado com {len(profissionais_unicos)} aba(s).")
        
    return relatorios_por_cliente