import streamlit as st
import pandas as pd
from datetime import datetime, date, timedelta
import os
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# ================= CONFIGURAÇÃO DA PÁGINA =================
st.set_page_config(
    page_title="Solicitação de Admissão",
    page_icon="📝",
    layout="centered"
)

st.title("📝 Solicitação de Admissão")
st.write("Preencha corretamente os dados abaixo para solicitar uma nova admissão.")

ARQUIVO = "solicitacoes_admissao.xlsx"

# ================= FUNÇÕES =================

def gerar_id():
    ano = datetime.now().year
    if os.path.exists(ARQUIVO):
        df = pd.read_excel(ARQUIVO)
        sequencial = len(df) + 1
    else:
        sequencial = 1
    return f"ADM-{ano}-{str(sequencial).zfill(5)}"


def enviar_email(dados):
    try:
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

        with smtplib.SMTP_SSL(smtp_host, smtp_port, timeout=20) as server:
            server.login(smtp_user, smtp_pass)
            server.send_message(msg)

    except Exception as e:
        st.error("❌ Erro ao enviar o e-mail.")
        st.exception(e)

# ================= FORMULÁRIO =================

with st.form("form_admissao"):
    st.subheader("👤 Identificação do Gestor")
    gestor_nome = st.text_input("Nome do gestor")
    gestor_email = st.text_input("E-mail do gestor")
    empresa = st.text_input("Empresa")
    cnpj = st.text_input("CNPJ da empresa")

    st.subheader("👨‍💼 Dados do Colaborador")
    colaborador_nome = st.text_input("Nome do colaborador")
    colaborador_email = st.text_input("E-mail do colaborador")

    st.subheader("📅 Dados da Admissão")
    data_admissao = st.date_input(
        "Data de admissão",
        min_value=date.today() + timedelta(days=1)
    )

    cargo = st.text_input("Cargo")
    salario = st.number_input(
        "Salário fixo mensal (R$)",
        min_value=0.0,
        step=100.0,
        format="%.2f"
    )

    enviar = st.form_submit_button("Enviar solicitação")

# ================= PROCESSAMENTO =================

if enviar:
    if not all([
        gestor_nome, gestor_email, empresa, cnpj,
        colaborador_nome, colaborador_email,
        cargo, salario > 0
    ]):
        st.error("❌ Preencha todos os campos obrigatórios.")
    else:
        dados = {
            "id_solicitacao": gerar_id(),
            "empresa": empresa,
            "cnpj": cnpj,
            "gestor_nome": gestor_nome,
            "gestor_email": gestor_email,
            "colaborador_nome": colaborador_nome,
            "colaborador_email": colaborador_email,
            "cargo": cargo,
            "salario": salario,
            "data_admissao": data_admissao,
            "data_solicitacao": datetime.now()
        }

        if os.path.exists(ARQUIVO):
            df = pd.read_excel(ARQUIVO)
            df = pd.concat([df, pd.DataFrame([dados])], ignore_index=True)
        else:
            df = pd.DataFrame([dados])

        df.to_excel(ARQUIVO, index=False)

        enviar_email(dados)

        st.success("✅ Solicitação enviada com sucesso!")
        st.info(f"📌 Protocolo da solicitação: **{dados['id_solicitacao']}**")
