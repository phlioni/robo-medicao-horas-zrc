# -*- coding: utf-8 -*-
import pandas as pd
import os
import re
from openpyxl.styles import Font
import calendar

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
    """
    Cria a tabela HTML detalhada para o projeto ZRC, com as horas divididas
    em dois períodos mensais e uma coluna de total.
    """
    if df_zrc is None or df_zrc.empty:
        return "<p>Não há dados para exibir.</p>"

    # Identificar os meses e anos presentes nos dados
    df_zrc['MesAno'] = df_zrc['Data'].dt.to_period('M')
    periodos = sorted(df_zrc['MesAno'].unique())

    # Garantir que temos no máximo 2 períodos para as colunas
    if len(periodos) == 0:
        return "<p>Não foi possível determinar o período de datas.</p>"
        
    periodo1 = periodos[0]
    periodo2 = periodos[1] if len(periodos) > 1 else None

    # Agrupar por projeto, profissional e período
    grouped = df_zrc.groupby(['Projeto', 'Profissional', 'MesAno'])['Horas'].sum().unstack(fill_value=0)
    
    # Adicionar coluna de total
    grouped['Total'] = grouped.sum(axis=1)
    
    # Resetar o índice para transformar multi-índice em colunas
    grouped = grouped.reset_index()

    # Formatação dos nomes dos meses para o cabeçalho
    meses_pt = {1:"Jan", 2:"Fev", 3:"Mar", 4:"Abr", 5:"Mai", 6:"Jun", 7:"Jul", 8:"Ago", 9:"Set", 10:"Out", 11:"Nov", 12:"Dez"}
    
    # Cabeçalho Período 1
    ultimo_dia_p1 = calendar.monthrange(periodo1.year, periodo1.month)[1]
    header_p1 = f"15/{periodo1.month:02d} - {ultimo_dia_p1}/{periodo1.month:02d}"

    # Cabeçalho Período 2
    header_p2 = f"01/{periodo2.month:02d} - 14/{periodo2.month:02d}" if periodo2 else ""

    # Construção do HTML
    html_output = ""
    total_geral_p1 = 0
    total_geral_p2 = 0
    total_geral_final = 0

    projetos = grouped['Projeto'].unique()
    for projeto in projetos:
        df_projeto = grouped[grouped['Projeto'] == projeto]
        
        # Totais por projeto
        total_projeto_p1 = df_projeto[periodo1].sum() if periodo1 in df_projeto else 0
        total_projeto_p2 = df_projeto[periodo2].sum() if periodo2 and periodo2 in df_projeto else 0
        total_projeto_final = df_projeto['Total'].sum()

        total_geral_p1 += total_projeto_p1
        total_geral_p2 += total_projeto_p2
        total_geral_final += total_projeto_final

        # Função para formatar números
        def format_num(n):
            return f"{n:,.2f}".replace(',', '#').replace('.', ',').replace('#', '.')

        # Linha de total do projeto
        html_output += f"""<tr style="background-color: #DDEBF7; font-weight: bold;">
                             <td style="padding: 8px; border: 1px solid #dddddd;">{projeto}</td>
                             <td style="padding: 8px; border: 1px solid #dddddd; text-align: right;">{format_num(total_projeto_p1)}</td>"""
        if periodo2:
            html_output += f'<td style="padding: 8px; border: 1px solid #dddddd; text-align: right;">{format_num(total_projeto_p2)}</td>'
        html_output += f'<td style="padding: 8px; border: 1px solid #dddddd; text-align: right;">{format_num(total_projeto_final)}</td></tr>'

        # Linhas dos profissionais
        for _, row in df_projeto.iterrows():
            horas_p1 = row.get(periodo1, 0)
            horas_p2 = row.get(periodo2, 0) if periodo2 else 0
            
            html_output += f"""<tr>
                                 <td style="padding: 8px; border: 1px solid #dddddd; padding-left: 25px;">{row['Profissional']}</td>
                                 <td style="padding: 8px; border: 1px solid #dddddd; text-align: right;">{format_num(horas_p1)}</td>"""
            if periodo2:
                html_output += f'<td style="padding: 8px; border: 1px solid #dddddd; text-align: right;">{format_num(horas_p2)}</td>'
            html_output += f'<td style="padding: 8px; border: 1px solid #dddddd; text-align: right;">{format_num(row["Total"])}</td></tr>'

    # Construção da tabela final com cabeçalhos dinâmicos
    header_html = f"""<th style="padding: 8px; border: 1px solid #dddddd; text-align: left;">Rótulos de Linha</th>
                      <th style="padding: 8px; border: 1px solid #dddddd; text-align: right;">{header_p1}</th>"""
    if periodo2:
        header_html += f'<th style="padding: 8px; border: 1px solid #dddddd; text-align: right;">{header_p2}</th>'
    header_html += '<th style="padding: 8px; border: 1px solid #dddddd; text-align: right;">Total de Horas</th>'

    # Rodapé com totais gerais
    footer_html = f"""<tr style="background-color: #DDEBF7; font-weight: bold;">
                        <td style="padding: 8px; border: 1px solid #dddddd;">Total Geral</td>
                        <td style="padding: 8px; border: 1px solid #dddddd; text-align: right;">{format_num(total_geral_p1)}</td>"""
    if periodo2:
        footer_html += f'<td style="padding: 8px; border: 1px solid #dddddd; text-align: right;">{format_num(total_geral_p2)}</td>'
    footer_html += f'<td style="padding: 8px; border: 1px solid #dddddd; text-align: right;">{format_num(total_geral_final)}</td></tr>'
    
    html_final = f"""<table style="width: 800px; border-collapse: collapse; font-family: Calibri, sans-serif; font-size: 11pt;">
                       <thead style="background-color: #4472C4; color: white;">
                         <tr>{header_html}</tr>
                       </thead>
                       <tbody>{html_output}</tbody>
                       <tfoot>{footer_html}</tfoot>
                     </table>"""
    return html_final

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