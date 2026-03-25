import streamlit as st
import pandas as pd
import io

# Configuração da página
st.set_page_config(page_title="Validador DR vs PR", layout="wide")

st.title("Validador de Demanda e Produção")

# Dicionário de conversão de Mercados (De -> Para)
de_para_mercados = {
    'MERCADO INTERNO': 'BRA',
    'EXPORTAÇÃO AMERICA DO SUL': 'OSA',
    'ARGENTINA': 'ARG'
}

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
                
                colunas_base = {10: 'Mercado', 11: 'Marca', 12: 'Produto', 13: 'Série'}
                meses = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez', 'Total MRP']
                colunas_mrp = {i + 41: meses[i] for i in range(13)}
                
                colunas_selecionadas = {**colunas_base, **colunas_mrp}
                df_dr = df_dr_raw[list(colunas_selecionadas.keys())].copy()
                df_dr.rename(columns=colunas_selecionadas, inplace=True)
                
                df_dr.dropna(subset=['Mercado', 'Marca', 'Produto', 'Série'], how='all', inplace=True)
                
                df_dr['Mercado'] = df_dr['Mercado'].astype(str).str.strip().str.upper()
                df_dr['Mercado'] = df_dr['Mercado'].replace(de_para_mercados)
                
                # --- NOVO: Força a conversão de todos os meses do DR para números puros ---
                for mes in meses:
                    df_dr[mes] = pd.to_numeric(df_dr[mes], errors='coerce').fillna(0)
                
                st.session_state['df_dr'] = df_dr.groupby(['Marca', 'Mercado', 'Produto', 'Série'])[meses].sum().reset_index()
                
                st.success(f"Aba '{aba_selecionada}' lida com sucesso! Mercados padronizados e números validados.")
                st.dataframe(st.session_state['df_dr'])
                
        except Exception as e:
            st.error(f"Ocorreu um erro na Etapa 1: {e}")

with aba2:
    st.subheader("Etapa 2: Consolidação dos Arquivos PR")
    st.write("Faça o upload dos 6 arquivos juntos. Os dados de julho em diante serão consolidados.")
    
    arquivos_pr = st.file_uploader("Upload dos arquivos Excel (PR)", type=["xlsx", "xls", "xlsm"], accept_multiple_files=True, key="upload_pr")
    
    if arquivos_pr:
        lista_pr = []
        erros_pr = []
        
        for arq in arquivos_pr:
            try:
                xls_pr = pd.ExcelFile(arq)
                abas_pr = xls_pr.sheet_names
                
                aba_alvo = None
                for aba in abas_pr:
                    if "production request" in aba.lower():
                        aba_alvo = aba
                        break
                        
                if not aba_alvo:
                    aba_alvo = abas_pr[0]
                    st.warning(f"Aviso: Aba 'Production Request' não encontrada com o nome exato no arquivo '{arq.name}'. Lendo a primeira aba: '{aba_alvo}'.")

                df_tmp = pd.read_excel(xls_pr, sheet_name=aba_alvo, skiprows=3)
                
                df_tmp = df_tmp.rename(columns={
                    df_tmp.columns[5]: 'Marca',
                    df_tmp.columns[6]: 'Mercado',
                    df_tmp.columns[7]: 'Produto',
                    df_tmp.columns[8]: 'Série'
                })
                
                df_tmp['Mercado'] = df_tmp['Mercado'].astype(str).str.strip().str.upper()
                df_tmp['Mercado'] = df_tmp['Mercado'].replace(de_para_mercados)
                
                colunas_chave = ['Marca', 'Mercado', 'Produto', 'Série']
                colunas_meses_pr = []
                
                for mes in meses_comparacao:
                    col_encontrada = [c for c in df_tmp.columns if str(c).strip().lower().startswith(mes.lower())]
                    if col_encontrada:
                        df_tmp = df_tmp.rename(columns={col_encontrada[0]: mes})
                        colunas_meses_pr.append(mes)
                    else:
                        df_tmp[mes] = 0
                        colunas_meses_pr.append(mes)
                
                df_limpo = df_tmp[colunas_chave + colunas_meses_pr]
                lista_pr.append(df_limpo)
                
            except Exception as e:
                erros_pr.append(f"Erro no arquivo {arq.name}: {e}")
        
        if erros_pr:
            for erro in erros_pr:
                st.error(erro)
                
        if lista_pr:
            df_pr_full = pd.concat(lista_pr, ignore_index=True)
            df_pr_full.dropna(subset=['Marca', 'Mercado', 'Produto', 'Série'], how='all', inplace=True)
            
            # --- NOVO: Força a conversão de todos os meses do PR para números puros ---
            for mes in meses_comparacao:
                df_pr_full[mes] = pd.to_numeric(df_pr_full[mes], errors='coerce').fillna(0)
            
            df_pr_resumo = df_pr_full.groupby(['Marca', 'Mercado', 'Produto', 'Série'])[meses_comparacao].sum().reset_index()
            df_pr_resumo['Total PR'] = df_pr_resumo[meses_comparacao].sum(axis=1)
            
            st.session_state['df_pr'] = df_pr_resumo
            
            st.success(f"{len(lista_pr)} arquivos consolidados com sucesso!")
            st.dataframe(st.session_state['df_pr'])

with aba3:
    st.subheader("Etapa 3: Resultado da Comparação")
    
    if 'df_dr' in st.session_state and 'df_pr' in st.session_state:
        df_dr_final = st.session_state['df_dr']
        df_pr_final = st.session_state['df_pr']
        
        dr_subset = df_dr_final[['Marca', 'Mercado', 'Produto', 'Série'] + meses_comparacao].copy()
        dr_subset['Total DR'] = dr_subset[meses_comparacao].sum(axis=1)
        
        df_merge = pd.merge(dr_subset, df_pr_final, on=['Marca', 'Mercado', 'Produto', 'Série'], how='outer', suffixes=('_DR', '_PR')).fillna(0)
        
        colunas_diferenca = []
        for mes in meses_comparacao:
            df_merge[f'Dif_{mes}'] = df_merge[f'{mes}_PR'] - df_merge[f'{mes}_DR']
            colunas_diferenca.append(f'Dif_{mes}')
            
        df_merge['Dif_Total'] = df_merge['Total PR'] - df_merge['Total DR']
        colunas_diferenca.append('Dif_Total')
        
        df_dif_marca_mercado = df_merge.groupby(['Marca', 'Mercado'])[colunas_diferenca].sum().reset_index()
        df_dif_detalhada = df_merge[['Marca', 'Mercado', 'Produto', 'Série'] + colunas_diferenca]
        
        if df_merge['Dif_Total'].abs().sum() == 0 and sum(df_merge[col].abs().sum() for col in colunas_diferenca[:-1]) == 0:
            st.success("🎉 TUDO OK! Os números batem perfeitamente. Nenhuma diferença encontrada de Julho a Dezembro.")
        else:
            st.warning("Atenção: Diferenças encontradas entre a demanda (DR) e a produção (PR).")
            
            st.markdown("#### Resumo da Diferença (Marca e Mercado)")
            st.dataframe(df_dif_marca_mercado[df_dif_marca_mercado['Dif_Total'] != 0])
            
            st.markdown("#### Detalhe Aberto por Série")
            st.dataframe(df_dif_detalhada[df_dif_detalhada['Dif_Total'] != 0])
            
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df_dr_final.to_excel(writer, index=False, sheet_name='DR')
            df_pr_final.to_excel(writer, index=False, sheet_name='PR_Consolidado')
            df_dif_marca_mercado.to_excel(writer, index=False, sheet_name='Dif_Marca_Mercado')
            df_dif_detalhada.to_excel(writer, index=False, sheet_name='Dif_Detalhada')
            
        st.markdown("---")
        st.download_button(
            label="📥 Baixar Consolidação e Diferenças em Excel",
            data=buffer.getvalue(),
            file_name="Analise_DR_vs_PR.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.info("👈 Por favor, carregue os arquivos nas Abas 1 (DR) e 2 (PR) para visualizar o cruzamento.")
