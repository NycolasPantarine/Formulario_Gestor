import streamlit as st
import pandas as pd
from datetime import datetime, date, timedelta
import os
import smtplib
import json
import gspread
from google.oauth2.service_account import Credentials
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication

# ======================================================
# CONFIG
# ======================================================
st.set_page_config(page_title="Solicitação de Admissão", page_icon="📝", layout="centered")

ARQ_SOLICITACOES = "solicitacoes_admissao.xlsx"

# ======================================================
# GOOGLE SHEETS
# ======================================================
@st.cache_resource
def get_gsheet_client():
    service_info = json.loads(st.secrets["GOOGLE_SERVICE_ACCOUNT_JSON"])
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_info(service_info, scopes=scopes)
    return gspread.authorize(creds)

def get_worksheet():
    gc = get_gsheet_client()
    sh = gc.open_by_key(st.secrets["GSHEETS_SPREADSHEET_ID"])
    return sh.worksheet(st.secrets.get("GSHEETS_WORKSHEET", "gestores"))

def to_bool(valor):
    # Aceita booleano real do Sheets
    if isinstance(valor, bool):
        return valor

    s = str(valor).strip().lower()

    return s in (
        "true",
        "1",
        "sim",
        "yes",
        "y",
        "verdadeiro",
        "verdade",
        "v"
    )

def carregar_gestores():
    ws = get_worksheet()
    dados = ws.get_all_values()

    if len(dados) <= 1:
        return pd.DataFrame(columns=["email_gestor", "nome_gestor", "empresa", "cnpj", "ativo"])

    cabecalho = dados[0]
    linhas = dados[1:]

    df = pd.DataFrame(linhas, columns=cabecalho)

    # Normalizações
    df["email_gestor"] = df["email_gestor"].astype(str).str.strip().str.lower()
    df["ativo"] = df["ativo"].apply(to_bool)

    return df

def identificar_gestor(email):
    email = email.strip().lower()
    df = carregar_gestores()

    gestor = df[
        (df["email_gestor"] == email) &
        (df["ativo"] == True)
    ]

    if gestor.empty:
        return None

    return gestor.iloc[0].to_dict()

# ======================================================
# AUX
# ======================================================
def gerar_id():
    ano = datetime.now().year
    if os.path.exists(ARQ_SOLICITACOES):
        df = pd.read_excel(ARQ_SOLICITACOES)
        seq = len(df) + 1
    else:
        seq = 1
    return f"ADM-{ano}-{str(seq).zfill(5)}"

def enviar_email(dados, arquivo):
    smtp_host = st.secrets["SMTP_HOST"]
    smtp_port = int(st.secrets["SMTP_PORT"])
    smtp_user = st.secrets["SMTP_USER"]
    smtp_pass = st.secrets["SMTP_PASS"]
    smtp_from = st.secrets["SMTP_FROM"]

    destino = "nycolas.pantarine@futtorh.com.br"
    assunto = f"📥 Nova Solicitação de Admissão – {dados['empresa']}"

    html = f"""
    <html>
    <body style="font-family:Arial;">
        <h2>Nova Solicitação de Admissão</h2>
        <p><b>Protocolo:</b> {dados['id']}</p>
        <hr>
        <p><b>Empresa:</b> {dados['empresa']}<br>
        <b>CNPJ:</b> {dados['cnpj']}</p>

        <p><b>Gestor:</b> {dados['gestor_nome']} ({dados['gestor_email']})</p>
        <p><b>Colaborador:</b> {dados['colaborador_nome']} ({dados['colaborador_email']})</p>

        <p><b>Cargo:</b> {dados['cargo']}<br>
        <b>Salário:</b> R$ {dados['salario']:.2f}<br>
        <b>Data de admissão:</b> {dados['data_admissao']}</p>
    </body>
    </html>
    """

    msg = MIMEMultipart()
    msg["From"] = smtp_from
    msg["To"] = destino
    msg["Subject"] = assunto
    msg.attach(MIMEText(html, "html"))

    with open(arquivo, "rb") as f:
        anexo = MIMEApplication(f.read(), _subtype="xlsx")
        anexo.add_header("Content-Disposition", "attachment", filename=arquivo)
        msg.attach(anexo)

    with smtplib.SMTP_SSL(smtp_host, smtp_port) as server:
        server.login(smtp_user, smtp_pass)
        server.send_message(msg)

# ======================================================
# APP
# ======================================================
st.title("📝 Solicitação de Admissão")

email_gestor = st.text_input("📧 E-mail do gestor")

gestor = None
if email_gestor:
    gestor = identificar_gestor(email_gestor)

if email_gestor and not gestor:
    st.error("❌ E-mail não autorizado ou gestor inativo.")

if gestor:
    st.success("Gestor identificado")

    st.text_input("Empresa", gestor["empresa"], disabled=True)
    st.text_input("CNPJ", gestor["cnpj"], disabled=True)

    with st.form("form_admissao"):
        colaborador_nome = st.text_input("Nome do colaborador")
        colaborador_email = st.text_input("E-mail do colaborador")

        data_admissao = st.date_input(
            "Data de admissão",
            min_value=date.today() + timedelta(days=1)
        )

        cargo = st.text_input("Cargo")
        salario = st.number_input("Salário fixo", min_value=0.0, step=100.0)

        enviar = st.form_submit_button("Enviar solicitação")

    if enviar:
        dados = {
            "id": gerar_id(),
            "empresa": gestor["empresa"],
            "cnpj": gestor["cnpj"],
            "gestor_nome": gestor["nome_gestor"],
            "gestor_email": email_gestor,
            "colaborador_nome": colaborador_nome,
            "colaborador_email": colaborador_email,
            "cargo": cargo,
            "salario": salario,
            "data_admissao": data_admissao,
            "data_envio": datetime.now()
        }

        df = pd.read_excel(ARQ_SOLICITACOES) if os.path.exists(ARQ_SOLICITACOES) else pd.DataFrame()
        df = pd.concat([df, pd.DataFrame([dados])], ignore_index=True)
        df.to_excel(ARQ_SOLICITACOES, index=False)

        enviar_email(dados, ARQ_SOLICITACOES)

        st.success("✅ Solicitação enviada com sucesso!")
        st.info(f"Protocolo: {dados['id']}")
