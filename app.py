import streamlit as st
import pandas as pd

# Configuração da página
st.set_page_config(page_title="Validador MRP", layout="wide")

st.title("Comparador de Arquivos")
st.subheader("Etapa 1: Leitura e Resumo do MRP")

# Upload do arquivo
arquivo_1 = st.file_uploader("Faça o upload do seu arquivo Excel base", type=["xlsx", "xls"])

if arquivo_1 is not None:
    try:
        # Carrega o arquivo Excel para extrair os nomes das abas
        xls = pd.ExcelFile(arquivo_1)
        
        # Cria um menu dropdown para selecionar a aba
        aba_selecionada = st.selectbox("Selecione a aba desejada:", xls.sheet_names)
        
        if aba_selecionada:
            # Lê os dados pulando as 3 primeiras linhas (cabeçalhos bagunçados)
            df = pd.read_excel(xls, sheet_name=aba_selecionada, skiprows=3, header=None)
            
            # Mapeamento das colunas base (K=10, L=11, M=12, N=13)
            colunas_base = {10: 'Mercado', 11: 'Marca', 12: 'Produto', 13: 'Série'}
            
            # Mapeamento das colunas MRP (AP=41 até BB=53)
            meses = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez', 'Total MRP']
            colunas_mrp = {i + 41: meses[i] for i in range(13)}
            
            # Une os dois dicionários de colunas
            colunas_selecionadas = {**colunas_base, **colunas_mrp}
            
            # Filtra o dataframe para trazer apenas os índices que queremos
            df_mrp = df[list(colunas_selecionadas.keys())].copy()
            
            # Renomeia as colunas numéricas para os nomes reais
            df_mrp.rename(columns=colunas_selecionadas, inplace=True)
            
            # Limpeza: remove linhas onde Mercado, Marca, Produto e Série estejam totalmente vazios
            df_mrp.dropna(subset=['Mercado', 'Marca', 'Produto', 'Série'], how='all', inplace=True)
            
            # Preenche possíveis vazios nos meses com 0
            df_mrp[meses] = df_mrp[meses].fillna(0)
            
            # Cria o resumo somando os valores mensais agrupados pela chave
            df_resumo = df_mrp.groupby(['Mercado', 'Marca', 'Produto', 'Série'])[meses].sum().reset_index()
            
            st.success(f"Aba '{aba_selecionada}' processada com sucesso! Confira o resumo do MRP abaixo:")
            
            # Exibe a tabela final
            st.dataframe(df_resumo)
            
    except Exception as e:
        st.error(f"Ocorreu um erro ao processar o arquivo: {e}")