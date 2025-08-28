# -*- coding: utf-8 -*-
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from datetime import datetime

import config

def enviar_email_zrc(tabela_html, periodo_str):
    """Monta e envia o e-mail específico para o projeto ZRC."""
    print("Preparando e-mail para o projeto ZRC...")

    msg = MIMEMultipart('related')
    msg['Subject'] = f"Medição de Horas - Projeto ZRC - Período {periodo_str}"
    msg['From'] = f"{config.ASSINATURA_NOME} <{config.EMAIL_REMETENTE}>"
    msg['To'] = config.ZRC_DESTINATARIO_PRINCIPAL['email']
    if config.ZRC_DESTINATARIOS_COPIA:
        msg['Cc'] = ", ".join(config.ZRC_DESTINATARIOS_COPIA)
    
    destinatarios_lista = [config.ZRC_DESTINATARIO_PRINCIPAL['email']] + config.ZRC_DESTINATARIOS_COPIA

    corpo_html = f"""
    <html><head></head><body style="font-family: Calibri, sans-serif; font-size: 11pt;">
        <p>Olá {config.ZRC_DESTINATARIO_PRINCIPAL['nome']},</p>
        <p>Segue medição de horas para o projeto ZRC, referente ao período de {periodo_str}.</p><br>
        {tabela_html}<br>
        <p>Qualquer dúvida, estou à disposição.</p>
        <br>
        
        <p style="margin: 0; font-size: 11pt; color: black;"><b style="font-size: 12pt;">{config.ASSINATURA_NOME}</b><br>{config.ASSINATURA_CARGO}</p>
        <hr size="1" width="250" align="left" color="#333333">
        <p style="margin: 0; font-size: 11pt; color: black;"><a href="mailto:{config.ASSINATURA_EMAIL}" style="color: #007bff; text-decoration: none;">{config.ASSINATURA_EMAIL}</a><br>{config.ASSINATURA_TELEFONE1} | {config.ASSINATURA_TELEFONE2}<br><a href="http://{config.ASSINATURA_SITE}" style="color: #007bff; text-decoration: none;">{config.ASSINATURA_SITE}</a></p>
        <p style="margin-top: 10px;"><img src="cid:logo_mosten" height="50">&nbsp;&nbsp;<img src="cid:logo_selos" height="50"></p>
    </body></html>"""
    msg.attach(MIMEText(corpo_html, 'html'))
    
    try:
        with open(config.CAMINHO_LOGO_MOSTEN, 'rb') as f:
            img_mosten = MIMEImage(f.read()); img_mosten.add_header('Content-ID', '<logo_mosten>'); msg.attach(img_mosten)
        with open(config.CAMINHO_LOGO_SELOS, 'rb') as f:
            img_selos = MIMEImage(f.read()); img_selos.add_header('Content-ID', '<logo_selos>'); msg.attach(img_selos)
    except FileNotFoundError as e:
        print(f"AVISO: Não foi possível encontrar a imagem da assinatura: {e}. O e-mail será enviado sem ela.")

    try:
        server = smtplib.SMTP(config.SMTP_SERVER, config.SMTP_PORT)
        server.starttls()
        server.login(config.EMAIL_REMETENTE, config.EMAIL_SENHA)
        server.sendmail(config.EMAIL_REMETENTE, destinatarios_lista, msg.as_string())
        server.quit()
        print(f"E-mail do projeto ZRC enviado com sucesso.")
    except Exception as e:
        print(f"ERRO ao enviar e-mail do ZRC: {e}")

def enviar_email_status(timing_report, status, erro_msg=""):
    """Envia um e-mail final com o status da execução e os tempos de cada etapa."""
    print("Preparando e-mail de status da execução...")
    
    if not config.STATUS_EMAIL_DESTINATARIO:
        print("Destinatário do e-mail de status não configurado. Pulando envio.")
        return

    msg = MIMEMultipart()
    
    status_cor = "green" if status_final == "SUCESSO" else "red"
    msg['Subject'] = f"Status do Robô ZRC: {status_final}"
    msg['From'] = f"Robô ZRC <{config.EMAIL_REMETENTE}>"
    msg['To'] = config.STATUS_EMAIL_DESTINATARIO

    tabela_html_linhas = ""
    for etapa, tempo in timing_report.items():
        tabela_html_linhas += f"<tr><td style='padding: 8px; border: 1px solid #dddddd;'>{etapa}</td><td style='padding: 8px; border: 1px solid #dddddd; text-align: right;'>{tempo:.2f} segundos</td></tr>"

    corpo_html = f"""
    <html><head></head><body style="font-family: Calibri, sans-serif; font-size: 11pt;">
        <h2>Relatório de Execução do Robô ZRC</h2>
        <p><b>Status Final:</b> <span style="color: {status_cor}; font-weight: bold;">{status_final}</span></p>
        <p><b>Data e Hora:</b> {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}</p>
        <h3>Tempos de Execução por Etapa:</h3>
        <table style="width: 600px; border-collapse: collapse;">
            <thead style="background-color: #4472C4; color: white;">
                <tr><th style="padding: 8px; border: 1px solid #dddddd; text-align: left;">Etapa</th><th style="padding: 8px; border: 1px solid #dddddd; text-align: right;">Duração</th></tr>
            </thead>
            <tbody>{tabela_html_linhas}</tbody>
        </table>
    """

    if status_final == "FALHA":
        corpo_html += f"""<h3 style="color: red;">Detalhes do Erro:</h3><p style="font-family: 'Courier New', monospace; background-color: #f5f5f5; padding: 10px;">{erro_msg}</p>"""

    corpo_html += "</body></html>"
    msg.attach(MIMEText(corpo_html, 'html'))

    try:
        server = smtplib.SMTP(config.SMTP_SERVER, config.SMTP_PORT)
        server.starttls()
        server.login(config.EMAIL_REMETENTE, config.EMAIL_SENHA)
        server.sendmail(config.EMAIL_REMETENTE, config.STATUS_EMAIL_DESTINATARIO, msg.as_string())
        server.quit()
        print(f"E-mail de status enviado com sucesso.")
    except Exception as e:
        print(f"ERRO ao enviar e-mail de status: {e}")