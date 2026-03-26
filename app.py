import streamlit as st
import pandas as pd
import io
import urllib.parse
import datetime

# Configuração da página
st.set_page_config(page_title="Validador DR vs PR", layout="wide")

st.title("Validador de Demanda e Produção v3.0 🚀 (RECONSTRUÍDO)")

# --- REGRAS DE NEGÓCIO ---
produtos_alvo = ['TA', 'PA', 'PU', 'CO']

de_para_mercados = {
    'MERCADO INTERNO': 'BRA',
    'EXPORTAÇÃO AMERICA DO SUL': 'OSA',
    'ARGENTINA': 'ARG'
}

de_para_marcas = {
    'FE': 'FT'
}

email_padrao_teams = "ana.teste@outlook.com"

# --- FUNÇÃO GERADORA DO LINK COM MENSAGEM ---
def gerar_link_teams(email, marca, mercado, produto, diferenca):
    mensagem = f"Olá! Identificamos uma diferença no Validador DR vs PR.\n\n📍 Marca: {marca}\n🌍 Mercado: {mercado}\n🚜 Produto: {produto}\n⚠️ Diferença: {diferenca} unidades.\n\nPor favor, poderia verificar?"
    mensagem_codificada = urllib.parse.quote(mensagem)
    return f"https://teams.microsoft.com/l/chat/0/0?users={email}&message={mensagem_codificada}"

aba1, aba2, aba3 = st.tabs(["1. Arquivo Base (DR)", "2. Arquivos de Produção (PR)", "3. Resumo de Diferenças"])

meses_comparacao = ['Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez']

with aba1:
    st.subheader("Etapa 1: Leitura e Resumo do DR")
    arquivo_dr = st.file_uploader("Upload do arquivo Excel (DR)", type=["xlsx", "xls", "xlsm"], key="upload_dr")
    
    if arquivo_dr is not None:
        try:
            xls_dr = pd.ExcelFile(arquivo_dr)
            aba_selecionada = st.selectbox("Selecione a aba desejada:", xls_dr.sheet_names)
            
            if aba_selecionada:
                df_dr_raw = pd.read_excel(xls_dr, sheet_name=aba_selecionada, skiprows=3, header=None)
                
                colunas_base = {10: 'Mercado', 11: 'Marca', 12: 'Produto', 13: 'Série', 14: 'Planta'}
                meses = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez', 'Total MRP']
                colunas_mrp = {i + 41: meses[i] for i in range(13)}
                
                colunas_selecionadas = {**colunas_base, **colunas_mrp}
                df_dr = df_dr_raw[list(colunas_selecionadas.keys())].copy()
                df_dr.rename(columns=colunas_selecionadas, inplace=True)
                
                df_dr.dropna(subset=['Mercado', 'Marca', 'Produto', 'Série'], how='all', inplace=True)
                
                df_dr['Planta'] = df_dr['Planta'].astype(str).str.strip().str.upper()
                df_dr = df_dr[df_dr['Planta'].str.startswith('BRA')]
                
                df_dr['Produto'] = df_dr['Produto'].astype(str).str.strip().str.upper()
                df_dr = df_dr[df_dr['Produto'].isin(produtos_alvo)]
                
                df_dr['Mercado'] = df_dr['Mercado'].astype(str).str.strip().str.upper().replace(de_para_mercados)
                df_dr['Marca'] = df_dr['Marca'].astype(str).str.strip().str.upper().replace(de_para_marcas)
                
                for mes in meses:
                    df_dr[mes] = pd.to_numeric(df_dr[mes], errors='coerce').fillna(0)
                
                chaves_agrupamento = ['Marca', 'Mercado', 'Produto', 'Série']
                df_dr[chaves_agrupamento] = df_dr[chaves_agrupamento].fillna('N/A')
                
                st.session_state['df_dr'] = df_dr.groupby(chaves_agrupamento)[meses].sum().reset_index()
                
                st.success(f"Aba '{aba_selecionada}' processada com sucesso!")
                st.dataframe(st.session_state['df_dr'])
                
        except Exception as e:
            st.error(f"Ocorreu um erro na Etapa 1: {e}")

with aba2:
    st.subheader("
