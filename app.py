import streamlit as st
import pandas as pd
from datetime import datetime, date, timedelta
import os
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication

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


def enviar_email_html(dados, arquivo_excel):
    try:
        smtp_host = st.secrets["SMTP_HOST"]
        smtp_port = int(st.secrets["SMTP_PORT"])
        smtp_user = st.secrets["SMTP_USER"]
        smtp_pass = st.secrets["SMTP_PASS"]
        smtp_from = st.secrets["SMTP_FROM"]

        destino = "nycolas.pantarine@futtorh.com.br"

        assunto = f"📥 Nova Solicitação de Admissão – {dados['empresa']}"

        html = f"""
        <html>
        <body style="font-family: Arial, sans-serif; background-color: #f5f5f5; padding: 20px;">
            <div style="max-width: 600px; margin: auto; background: #ffffff; padding: 20px; border-radius: 6px;">
                
                <h2 style="color:#1f2937;">Nova Solicitação de Admissão</h2>

                <p><strong>Protocolo:</strong> {dados['id_solicitacao']}</p>

                <hr>

                <h3>🏢 Empresa</h3>
                <p>
                    <strong>Nome:</strong> {dados['empresa']}<br>
                    <strong>CNPJ:</strong> {dados['cnpj']}
                </p>

                <h3>👤 Gestor</h3>
                <p>
                    <strong>Nome:</strong> {dados['gestor_nome']}<br>
                    <strong>E-mail:</strong> {dados['gestor_email']}
                </p>

                <h3>👨‍💼 Colaborador</h3>
                <p>
                    <strong>Nome:</strong> {dados['colaborador_nome']}<br>
                    <strong>E-mail:</strong> {dados['colaborador_email']}
                </p>

                <h3>📅 Dados da Admissão</h3>
                <p>
                    <strong>Cargo:</strong> {dados['cargo']}<br>
                    <strong>Salário:</strong> R$ {dados['salario']:.2f}<br>
                    <strong>Data de admissão:</strong> {dados['data_admissao']}
                </p>

                <hr>

                <p style="font-size: 12px; color: #6b7280;">
                    Solicitação enviada em {dados['data_solicitacao']}
                </p>

            </div>
        </body>
        </html>
        """

        msg = MIMEMultipart()
        msg["From"] = smtp_from
        msg["To"] = destino
        msg["Subject"] = assunto

        msg.attach(MIMEText(html, "html"))

        # ===== ANEXO EXCEL =====
        with open(arquivo_excel, "rb") as f:
            anexo = MIMEApplication(f.read(), _subtype="xlsx")
            anexo.add_header(
                "Content-Disposition",
                "attachment",
                filename=os.path.basename(arquivo_excel)
            )
            msg.attach(anexo)

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

        enviar_email_html(dados, ARQUIVO)

        st.success("✅ Solicitação enviada com sucesso!")
        st.info(f"📌 Protocolo da solicitação: **{dados['id_solicitacao']}**")
