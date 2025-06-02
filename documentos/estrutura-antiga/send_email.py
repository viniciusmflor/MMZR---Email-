import os
import numpy as np
import pandas as pd
from datetime import date, datetime, timedelta

import win32com.client as win32
from pretty_html_table import build_table

def make_initial_html_body(data_ref):
    html_body = """<!DOCTYPE html>
    <html>
    <head>
        <meta charset="utf-8">
    </head>
    """
    html_sec = f'''<div>
        <h2 style="font-family: 'Titillium Web'; color: #061844;"> Acompanhamento das Aplicações e Resgates do Allocation</h2>
        <ul style="margin-bottom: 2px;"">
    '''
    html_sec += f'''Data Ref: {data_ref}<br><br>'''
    html_sec += f'''Legenda das colunas:
    <ul style="margin-bottom: 2px;"">
    <li>Nos últimos 15 dias</li>
    <li>No mês atual</li>
    <li>No ano atual</li>
    <li>Nos últimos 12 meses</li>
    </ul>
    '''
    html_body += html_sec

    return html_body

def make_section(df_alvo:pd.DataFrame, nome_seção:str) -> str:
    html_body = ""
    html_sec = f'''<div>
        <h2 style="font-family: 'Titillium Web'; color: #061844;">{nome_seção}</h2>
        <ul style="margin-bottom: 2px;"">
    '''
    html_body += html_sec

    html_table_blue_light = build_table(df_alvo, 'blue_light')
    html_body += html_table_blue_light \
                    .replace('<tr>', '<tr align="center">') \
                    .replace('"text-align: right;">', '"text-align: center;">') \
                    .replace('width: auto"', '"')
    
    if nome_seção == "Resgates Solicitados (Soma)":
        html_body += "<br>* Os resgates que aconteceram nos últimos 15 dias ainda podem estar em período de liquidação."
        html_body += "<br>* O valor aqui registrado é sempre dos resgates solicitados (data de cotização do resgate), não necessariamente já retirado do fundo para pagamento ao cotista."


    html_body += ''' </ul> '''
    html_body += ''' </div> '''
    return html_body

def send_email(html_body, destination, emails_to_cc=None):
    print("Abrindo Outlook")
    outlook = win32.Dispatch('Outlook.Application') # create outlook object
    mail = outlook.CreateItem(0) # create e-mail message object
    mail.To = ";".join(destination)
    if emails_to_cc:
        mail.CC = ";".join(emails_to_cc)

    mail.Subject = f'Relatório Aplicações e Resgates Allocation {date.today().strftime("%d/%m/%Y")}' # e-mail subject
    mail.HTMLBody = html_body
    mail.Display(True)
    # mail.send() # send e-mail

def main(data_ref, df_captacao, df_resgate, df_PL, df_cotistas, destination, emails_to_cc=None):
    html_body = make_initial_html_body(data_ref)
    
    html_body += make_section(df_captacao, "Aplicações (Soma)")
    html_body += make_section(df_resgate, "Resgates Solicitados (Soma)")
    html_body += make_section(df_PL, "PL (Diferença entre PL ajustado pelas quotas + Captação - Resgate)")
    html_body += make_section(df_cotistas, "Cotistas (Delta)")

    html_body += ''' </html> '''

    send_email(html_body, destination, emails_to_cc=emails_to_cc)