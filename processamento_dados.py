# -*- coding: utf-8 -*-
import pandas as pd

def processar_planilha(caminho_arquivo):
    """
    Lê o relatório, filtra para o projeto ZRC e prepara os dados.
    Retorna apenas o DataFrame filtrado.
    """
    print("Processando a planilha...")
    try:
        df = pd.read_excel(caminho_arquivo, skiprows=18, engine='openpyxl')
        df.columns = df.columns.str.strip()
        
        # --- FILTRO ESPECÍFICO PARA O PROJETO ZRC ---
        df_filtrado = df[df['Projeto'].str.startswith('ZRC', na=False)].copy()

        if df_filtrado.empty:
            print("AVISO: Nenhum lançamento para o projeto ZRC encontrado no relatório.")
            return None

        # Aplica os filtros de situação e comentários
        #df_filtrado = df_filtrado[df_filtrado['Situação'] == 'Aprovado'].copy()
        #df_filtrado['Comentários'] = df_filtrado['Comentários'].astype(str)
        #df_filtrado = df_filtrado[df_filtrado['Comentários'].str.len() >= 5]
        
        print("Planilha do projeto ZRC processada com sucesso.")
        return df_filtrado

    except Exception as e:
        print(f"Ocorreu um erro inesperado ao processar a planilha: {e}")
        return None

def criar_tabela_html(df_zrc):
    """Cria a tabela HTML detalhada (com profissionais) para o projeto ZRC."""
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