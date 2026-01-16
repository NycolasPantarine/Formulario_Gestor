import streamlit as st
import pandas as pd
from datetime import datetime, date
import os
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart


def enviar_email(dados):
    smtp_host = st.secrets["SMTP_HOST"]
    smtp_port = int(st.secrets["SMTP_PORT"])
    smtp_user = st.secrets["SMTP_USER"]
    smtp_pass = st.secrets["SMTP_PASS"]
    smtp_from = st.secrets["SMTP_FROM"]

    destino = "nycolas.pantarine@futtorh.com.br"

    assunto = f"📥 Nova solicitação de admissão - {dados['empresa']}"

    corpo = f"""
Nova solicitação de admissão recebida.

Protocolo: {dados['id_solicitacao']}

Empresa: {dados['empresa']}
CNPJ: {dados['cnpj']}

Gestor:
- Nome: {dados['gestor_nome']}
- E-mail: {dados['gestor_email']}

Colaborador:
- Nome: {dados['colaborador_nome']}
- E-mail: {dados['colaborador_email']}

Admissão:
- Cargo: {dados['cargo']}
- Salário: R$ {dados['salario']:.2f}
- Data de admissão: {dados['data_admissao']}

Data da solicitação: {dados['data_solicitacao']}
"""

    msg = MIMEMultipart()
    msg["From"] = smtp_from
    msg["To"] = destino
    msg["Subject"] = assunto
    msg.attach(MIMEText(corpo, "plain"))

    with smtplib.SMTP_SSL(smtp_host, smtp_port) as server:
        server.login(smtp_user, smtp_pass)
        server.send_message(msg)
