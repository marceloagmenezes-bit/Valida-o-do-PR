import streamlit as st
import pandas as pd
import io # Nova biblioteca nativa para gerenciar arquivos na memória

# Configuração da página
st.set_page_config(page_title="Validador MRP", layout="wide")

st.title("Comparador de Arquivos")
st.subheader("Etapa 1: Leitura e Resumo do MRP")

# Upload do arquivo (Agora aceita .xlsm também!)
arquivo_1 = st.file_uploader("Faça o upload do seu arquivo Excel base", type=["xlsx", "xls", "xlsm"])

if arquivo_1 is not None:
    try:
        # Carrega o arquivo Excel para extrair os nomes das abas
        xls = pd.ExcelFile(arquivo_1)
        
        # Cria um menu dropdown para selecionar a aba
        aba_selecionada = st.selectbox("Selecione a aba desejada:", xls.sheet_names)
        
        if aba_selecionada:
            # Lê os dados pulando as 3 primeiras linhas
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
            
            # Limpeza: remove linhas vazias nas chaves principais
            df_mrp.dropna(subset=['Mercado', 'Marca', 'Produto', 'Série'], how='all', inplace=True)
            
            # Preenche possíveis vazios nos meses com 0
            df_mrp[meses] = df_mrp[meses].fillna(0)
            
            # Cria o resumo agrupado
            df_resumo = df_mrp.groupby(['Mercado', 'Marca', 'Produto', 'Série'])[meses].sum().reset_index()
            
            st.success(f"Aba '{aba_selecionada}' processada com sucesso! Confira o resumo do MRP abaixo:")
            st.dataframe(df_resumo)
            
            # --- NOVA SESSÃO: DOWNLOAD EM EXCEL ---
            st.markdown("---")
            st.subheader("Validação dos Dados")
            st.write("Baixe o resumo gerado acima para validar os números no seu computador.")
            
            # Cria um arquivo Excel na memória do servidor
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                df_resumo.to_excel(writer, index=False, sheet_name='Resumo_Gerado')
            
            # Cria o botão de download
            st.download_button(
                label="📥 Baixar Resumo em Excel",
                data=buffer.getvalue(),
                file_name="Resumo_MRP_Validacao.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
    except Exception as e:
        st.error(f"Ocorreu um erro ao processar o arquivo: {e}")
