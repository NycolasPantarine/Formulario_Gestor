import streamlit as st
import pandas as pd
from datetime import datetime, date
import os

st.set_page_config(
    page_title="Solicitação de Admissão",
    page_icon="📝",
    layout="centered"
)

st.title("📝 Solicitação de Admissão")
st.write("Preencha corretamente os dados abaixo para solicitar uma nova admissão.")

ARQUIVO = "solicitacoes_admissao.xlsx"

def gerar_id():
    ano = datetime.now().year
    if os.path.exists(ARQUIVO):
        df = pd.read_excel(ARQUIVO)
        sequencial = len(df) + 1
    else:
        sequencial = 1
    return f"ADM-{ano}-{str(sequencial).zfill(5)}"

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
        min_value=date.today() + pd.Timedelta(days=1)
    )
    cargo = st.text_input("Cargo")
    salario = st.number_input(
        "Salário fixo mensal (R$)",
        min_value=0.0,
        step=100.0,
        format="%.2f"
    )

    enviar = st.form_submit_button("Enviar solicitação")

if enviar:
    if not all([
        gestor_nome, gestor_email, empresa, cnpj,
        colaborador_nome, colaborador_email,
        cargo, salario > 0
    ]):
        st.error("❌ Preencha todos os campos obrigatórios.")
    else:
        nova_linha = {
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
            df = pd.concat([df, pd.DataFrame([nova_linha])], ignore_index=True)
        else:
            df = pd.DataFrame([nova_linha])

        df.to_excel(ARQUIVO, index=False)

        st.success("✅ Solicitação enviada com sucesso!")
        st.info(f"📌 Protocolo da solicitação: **{nova_linha['id_solicitacao']}**")
