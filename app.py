import streamlit as st
import subprocess
import os

st.set_page_config(
    page_title="Automação SIGA"
)

st.write('## Seja bem-vindo(a)!')
st.write('### Essa ferramenta irá baixar os relatórios financeiros do sistema automaticamente. Isso vai te poupar tempo e energia!')

arquivo_excel = st.file_uploader("Por gentileza, faça o upload do arquivo .xlsx que contém os dados para busca:", type="xlsx")

if arquivo_excel is not None:
    with st.spinner("Baixando os planos financeiros..."):
        # Salvar o arquivo upload temporariamente
        caminho_arquivo = f"temp_{arquivo_excel.name}"
        with open(caminho_arquivo, "wb") as f:
            f.write(arquivo_excel.getbuffer())
        
        # Chamar o script de automação como um subprocesso
        result = subprocess.run(["python", "automacao.py", caminho_arquivo], capture_output=True, text=True)

        if result.returncode == 0:
            st.success('Arquivos baixados com sucesso!')
        else:
            st.error(f'Erro na automação: {result.stderr}')