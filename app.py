import streamlit as st
import pandas as pd
from datetime import datetime, date, timedelta
import os
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication

# ======================================================
# CONFIGURAÇÕES
# ======================================================
st.set_page_config(page_title="Solicitação de Admissão", page_icon="📝", layout="centered")

ARQ_GESTORES = "gestores.xlsx"
ARQ_SOLICITACOES = "solicitacoes_admissao.xlsx"

# ======================================================
# FUNÇÕES BASE
# ======================================================
def inicializar_gestores():
    if not os.path.exists(ARQ_GESTORES):
        df = pd.DataFrame(columns=[
            "email_gestor", "nome_gestor", "empresa", "cnpj", "ativo"
        ])
        df.to_excel(ARQ_GESTORES, index=False)

def carregar_gestores():
    inicializar_gestores()
    return pd.read_excel(ARQ_GESTORES)

def salvar_gestores(df):
    df.to_excel(ARQ_GESTORES, index=False)

def identificar_gestor(email):
    df = carregar_gestores()
    gestor = df[
        (df["email_gestor"].str.lower() == email.lower()) &
        (df["ativo"] == True)
    ]
    if gestor.empty:
        return None
    return gestor.iloc[0].to_dict()

def gerar_id():
    ano = datetime.now().year
    if os.path.exists(ARQ_SOLICITACOES):
        df = pd.read_excel(ARQ_SOLICITACOES)
        seq = len(df) + 1
    else:
        seq = 1
    return f"ADM-{ano}-{str(seq).zfill(5)}"

# ======================================================
# EMAIL
# ======================================================
def enviar_email_html(dados, arquivo_excel):
    smtp_host = st.secrets["SMTP_HOST"]
    smtp_port = int(st.secrets["SMTP_PORT"])
    smtp_user = st.secrets["SMTP_USER"]
    smtp_pass = st.secrets["SMTP_PASS"]
    smtp_from = st.secrets["SMTP_FROM"]

    destino = "nycolas.pantarine@futtorh.com.br"

    assunto = f"📥 Nova Solicitação de Admissão – {dados['empresa']}"

    html = f"""
    <html>
    <body style="font-family:Arial">
        <h2>Nova Solicitação de Admissão</h2>
        <p><b>Protocolo:</b> {dados['id_solicitacao']}</p>
        <hr>
        <p><b>Empresa:</b> {dados['empresa']}<br>
           <b>CNPJ:</b> {dados['cnpj']}</p>
        <p><b>Gestor:</b> {dados['gestor_nome']} ({dados['gestor_email']})</p>
        <p><b>Colaborador:</b> {dados['colaborador_nome']} ({dados['colaborador_email']})</p>
        <p><b>Cargo:</b> {dados['cargo']}<br>
           <b>Salário:</b> R$ {dados['salario']:.2f}<br>
           <b>Admissão:</b> {dados['data_admissao']}</p>
    </body>
    </html>
    """

    msg = MIMEMultipart()
    msg["From"] = smtp_from
    msg["To"] = destino
    msg["Subject"] = assunto
    msg.attach(MIMEText(html, "html"))

    with open(arquivo_excel, "rb") as f:
        anexo = MIMEApplication(f.read(), _subtype="xlsx")
        anexo.add_header("Content-Disposition", "attachment", filename=os.path.basename(arquivo_excel))
        msg.attach(anexo)

    with smtplib.SMTP_SSL(smtp_host, smtp_port) as server:
        server.login(smtp_user, smtp_pass)
        server.send_message(msg)

# ======================================================
# PAINEL ADMIN
# ======================================================
st.sidebar.title("🔐 Painel Interno")
senha = st.sidebar.text_input("Senha administrativa", type="password")

if senha == st.secrets["ADMIN_PASSWORD"]:
    st.sidebar.success("Acesso liberado")

    st.header("👤 Cadastro de Gestores")

    df_gestores = carregar_gestores()

    email = st.text_input("E-mail do gestor").lower()
    nome = st.text_input("Nome do gestor")
    empresa = st.text_input("Empresa")
    cnpj = st.text_input("CNPJ")
    ativo = st.checkbox("Ativo", value=True)

    if st.button("💾 Salvar gestor"):
        df_gestores = df_gestores[df_gestores["email_gestor"] != email]
        novo = pd.DataFrame([{
            "email_gestor": email,
            "nome_gestor": nome,
            "empresa": empresa,
            "cnpj": cnpj,
            "ativo": ativo
        }])
        df_gestores = pd.concat([df_gestores, novo], ignore_index=True)
        salvar_gestores(df_gestores)
        st.success("Gestor salvo com sucesso")

    st.subheader("📋 Gestores cadastrados")
    st.dataframe(df_gestores)

st.divider()

# ======================================================
# FORMULÁRIO PÚBLICO
# ======================================================
st.title("📝 Solicitação de Admissão")
email_input = st.text_input("📧 E-mail do gestor")

gestor = identificar_gestor(email_input) if email_input else None

if email_input and not gestor:
    st.error("❌ E-mail não autorizado")

if gestor:
    st.success("Gestor identificado")
    st.text_input("Empresa", gestor["empresa"], disabled=True)
    st.text_input("CNPJ", gestor["cnpj"], disabled=True)

    with st.form("form_admissao"):
        colaborador_nome = st.text_input("Nome do colaborador")
        colaborador_email = st.text_input("E-mail do colaborador")
        data_admissao = st.date_input("Data de admissão", min_value=date.today() + timedelta(days=1))
        cargo = st.text_input("Cargo")
        salario = st.number_input("Salário", min_value=0.0, step=100.0)

        enviar = st.form_submit_button("Enviar solicitação")

    if enviar:
        dados = {
            "id_solicitacao": gerar_id(),
            "empresa": gestor["empresa"],
            "cnpj": gestor["cnpj"],
            "gestor_nome": gestor["nome_gestor"],
            "gestor_email": email_input,
            "colaborador_nome": colaborador_nome,
            "colaborador_email": colaborador_email,
            "cargo": cargo,
            "salario": salario,
            "data_admissao": data_admissao,
            "data_solicitacao": datetime.now()
        }

        df = pd.read_excel(ARQ_SOLICITACOES) if os.path.exists(ARQ_SOLICITACOES) else pd.DataFrame()
        df = pd.concat([df, pd.DataFrame([dados])], ignore_index=True)
        df.to_excel(ARQ_SOLICITACOES, index=False)

        enviar_email_html(dados, ARQ_SOLICITACOES)
        st.success("✅ Solicitação enviada com sucesso")
