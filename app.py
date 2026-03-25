import streamlit as st
import pandas as pd
import io

# Configuração da página
st.set_page_config(page_title="Validador DR vs PR", layout="wide")

st.title("Validador de Demanda e Produção")

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
                
                # --- NOVO: Mapeamento da Coluna O (14) como Planta ---
                colunas_base = {10: 'Mercado', 11: 'Marca', 12: 'Produto', 13: 'Série', 14: 'Planta'}
                meses = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez', 'Total MRP']
                colunas_mrp = {i + 41: meses[i] for i in range(13)}
                
                colunas_selecionadas = {**colunas_base, **colunas_mrp}
                df_dr = df_dr_raw[list(colunas_selecionadas.keys())].copy()
                df_dr.rename(columns=colunas_selecionadas, inplace=True)
                
                df_dr.dropna(subset=['Mercado', 'Marca', 'Produto', 'Série'], how='all', inplace=True)
                
                # --- NOVO: FILTRO DA PLANTA (Começa com BRA) ---
                df_dr['Planta'] = df_dr['Planta'].astype(str).str.strip().str.upper()
                df_dr = df_dr[df_dr['Planta'].str.startswith('BRA')]
                
                # FILTRO DE PRODUTOS DR
                df_dr['Produto'] = df_dr['Produto'].astype(str).str.strip().str.upper()
                df_dr = df_dr[df_dr['Produto'].isin(produtos_alvo)]
                
                # Limpezas
                df_dr['Mercado'] = df_dr['Mercado'].astype(str).str.strip().str.upper()
                df_dr['Mercado'] = df_dr['Mercado'].replace(de_para_mercados)
                
                df_dr['Marca'] = df_dr['Marca'].astype(str).str.strip().str.upper()
                df_dr['Marca'] = df_dr['Marca'].replace(de_para_marcas)
                
                for mes in meses:
                    df_dr[mes] = pd.to_numeric(df_dr[mes], errors='coerce').fillna(0)
                
                st.session_state['df_dr'] = df_dr.groupby(['Marca', 'Mercado', 'Produto', 'Série'])[meses].sum().reset_index()
                
                st.success(f"Aba '{aba_selecionada}' processada com sucesso! Filtros aplicados: Plantas 'BRA' e Produtos Alvo.")
                st.dataframe(st.session_state['df_dr'])
                
        except Exception as e:
            st.error(f"Ocorreu um erro na Etapa 1: {e}")

with aba2:
    st.subheader("Etapa 2: Consolidação dos Arquivos PR")
    st.write("Faça o upload dos 6 arquivos juntos.")
    
    arquivos_pr = st.file_uploader("Upload dos arquivos Excel (PR)", type=["xlsx", "xls", "xlsm"], accept_multiple_files=True, key="upload_pr")
    
    if arquivos_pr:
        lista_pr_resumo = []
        lista_pr_bruto = []
        erros_pr = []
        
        for arq in arquivos_pr:
            try:
                xls_pr = pd.ExcelFile(arq)
                abas_pr = xls_pr.sheet_names
                
                aba_alvo = next((aba for aba in abas_pr if "production request" in aba.lower()), abas_pr[0])

                df_tmp_raw = pd.read_excel(xls_pr, sheet_name=aba_alvo, skiprows=3, header=None)
                
                if df_tmp_raw.empty:
                    continue

                # CRIANDO A MÁSCARA DE FILTRO DOS PRODUTOS
                produtos_raw = df_tmp_raw[7].iloc[1:].astype(str).str.strip().str.upper()
                mask_produtos = produtos_raw.isin(produtos_alvo)

                # 1. ESTEIRA DO ARQUIVO BRUTO
                df_bruto = df_tmp_raw.iloc[1:, 1:23].copy()
                df_bruto = df_bruto[mask_produtos]
                
                raw_cols = df_tmp_raw.iloc[0, 1:23].astype(str).tolist()
                unique_cols = []
                seen = set()
                
                for col in raw_cols:
                    new_col = col
                    counter = 1
                    while new_col in seen:
                        new_col = f"{col}_{counter}"
                        counter += 1
                    seen.add(new_col)
                    unique_cols.append(new_col)
                
                df_bruto.columns = unique_cols
                df_bruto['Arquivo_Origem'] = arq.name
                df_bruto.dropna(how='all', inplace=True)
                lista_pr_bruto.append(df_bruto)

                # 2. ESTEIRA DO RESUMO E COMPARAÇÃO
                df_resumo_temp = pd.DataFrame()
                
                df_resumo_temp['Marca'] = df_tmp_raw[5].iloc[1:]
                df_resumo_temp['Mercado'] = df_tmp_raw[6].iloc[1:]
                df_resumo_temp['Produto'] = df_tmp_raw[7].iloc[1:]
                df_resumo_temp['Série'] = df_tmp_raw[8].iloc[1:]
                
                headers = df_tmp_raw.iloc[0]
                meses_indices = {}
                
                for idx in range(9, min(23, len(headers))):
                    val = headers[idx]
                    if pd.isna(val): continue
                    
                    if isinstance(val, pd.Timestamp):
                        m = val.month
                        if m == 7: meses_indices['Jul'] = idx
                        elif m == 8: meses_indices['Ago'] = idx
                        elif m == 9: meses_indices['Set'] = idx
                        elif m == 10: meses_indices['Out'] = idx
                        elif m == 11: meses_indices['Nov'] = idx
                        elif m == 12: meses_indices['Dez'] = idx
                    else:
                        val_str = str(val).lower()
                        if 'jul' in val_str or '07' in val_str: meses_indices['Jul'] = idx
                        elif 'ago' in val_str or 'aug' in val_str or '08' in val_str: meses_indices['Ago'] = idx
                        elif 'set' in val_str or 'sep' in val_str or '09' in val_str: meses_indices['Set'] = idx
                        elif 'out' in val_str or 'oct' in val_str or '10' in val_str: meses_indices['Out'] = idx
                        elif 'nov' in val_str or '11' in val_str: meses_indices['Nov'] = idx
                        elif 'dez' in val_str or 'dec' in val_str or '12' in val_str: meses_indices['Dez'] = idx

                for mes in meses_comparacao:
                    if mes in meses_indices:
                        df_resumo_temp[mes] = df_tmp_raw[meses_indices[mes]].iloc[1:]
                    else:
                        df_resumo_temp[mes] = 0
                
                df_resumo_temp = df_resumo_temp[mask_produtos]
                lista_pr_resumo.append(df_resumo_temp)
                
            except Exception as e:
                erros_pr.append(f"Erro no arquivo {arq.name}: {e}")
        
        if erros_pr:
            for erro in erros_pr:
                st.error(erro)
                
        if lista_pr_resumo:
            df_pr_full = pd.concat(lista_pr_resumo, ignore_index=True)
            df_pr_full.dropna(subset=['Marca', 'Mercado', 'Produto', 'Série'], how='all', inplace=True)
            
            df_pr_full['Mercado'] = df_pr_full['Mercado'].astype(str).str.strip().str.upper().replace(de_para_mercados)
            df_pr_full['Marca'] = df_pr_full['Marca'].astype(str).str.strip().str.upper().replace(de_para_marcas)
            
            for mes in meses_comparacao:
                df_pr_full[mes] = pd.to_numeric(df_pr_full[mes], errors='coerce').fillna(0)
            
            df_pr_resumo_final = df_pr_full.groupby(['Marca', 'Mercado', 'Produto', 'Série'])[meses_comparacao].sum().reset_index()
            df_pr_resumo_final['Total PR'] = df_pr_resumo_final[meses_comparacao].sum(axis=1)
            
            if lista_pr_bruto:
                df_pr_bruto_final = pd.concat(lista_pr_bruto, ignore_index=True)
                df_pr_bruto_final.dropna(how='all', inplace=True)
                st.session_state['df_pr_bruto'] = df_pr_bruto_final 
            
            st.session_state['df_pr'] = df_pr_resumo_final
            
            st.success(f"{len(lista_pr_resumo)} arquivos consolidados com sucesso! Apenas TA, PA, PU e CO.")
            st.dataframe(st.session_state['df_pr'])

with aba3:
    st.subheader("Etapa 3: Resultado da Comparação")
    
    if 'df_dr' in st.session_state and 'df_pr' in st.session_state:
        df_dr_final = st.session_state['df_dr']
        df_pr_final = st.session_state['df_pr']
        df_pr_bruto_export = st.session_state.get('df_pr_bruto', pd.DataFrame())
        
        dr_subset = df_dr_final[['Marca', 'Mercado', 'Produto', 'Série'] + meses_comparacao].copy()
        dr_subset['Total DR'] = dr_subset[meses_comparacao].sum(axis=1)
        
        df_merge = pd.merge(dr_subset, df_pr_final, on=['Marca', 'Mercado', 'Produto', 'Série'], how='outer', suffixes=('_DR', '_PR')).fillna(0)
        
        colunas_diferenca = []
        for mes in meses_comparacao:
            df_merge[f'Dif_{mes}'] = df_merge[f'{mes}_PR'] - df_merge[f'{mes}_DR']
            colunas_diferenca.append(f'Dif_{mes}')
            
        df_merge['Dif_Total'] = df_merge['Total PR'] - df_merge['Total DR']
        colunas_diferenca.append('Dif_Total')
        
        df_dif_resumo = df_merge.groupby(['Marca', 'Mercado', 'Produto'])[colunas_diferenca].sum().reset_index()
        df_dif_detalhada = df_merge[['Marca', 'Mercado', 'Produto', 'Série'] + colunas_diferenca]
        
        if df_merge['Dif_Total'].abs().sum() == 0 and sum(df_merge[col].abs().sum() for col in colunas_diferenca[:-1]) == 0:
            st.success("🎉 TUDO OK! Os números batem perfeitamente. Nenhuma diferença encontrada de Julho a Dezembro para os produtos selecionados.")
        else:
            st.warning("Atenção: Diferenças encontradas entre a demanda (DR) e a produção (PR).")
            
            st.markdown("#### Resumo da Diferença (Marca, Mercado e Produto)")
            st.dataframe(df_dif_resumo[df_dif_resumo['Dif_Total'] != 0])
            
            st.markdown("#### Detalhe Aberto por Série")
            st.dataframe(df_dif_detalhada[df_dif_detalhada['Dif_Total'] != 0])
            
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df_dr_final.to_excel(writer, index=False, sheet_name='DR')
            
            if not df_pr_bruto_export.empty:
                df_pr_bruto_export.to_excel(writer, index=False, sheet_name='PR_Bruto_Completo')
                
            df_pr_final.to_excel(writer, index=False, sheet_name='PR_Resumo_Consolidado')
            
            df_dif_resumo.to_excel(writer, index=False, sheet_name='Dif_Resumo')
            df_dif_detalhada.to_excel(writer, index=False, sheet_name='Dif_Detalhada')
            
        st.markdown("---")
        st.download_button(
            label="📥 Baixar Consolidação Total (Com Aba Bruta) em Excel",
            data=buffer.getvalue(),
            file_name="Analise_DR_vs_PR_Completo.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.info("👈 Por favor, carregue os arquivos nas Abas 1 (DR) e 2 (PR) para visualizar o cruzamento.")
