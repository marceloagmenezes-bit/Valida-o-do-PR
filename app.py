import streamlit as st
import pandas as pd
import io
import urllib.parse

# Configuração da página
st.set_page_config(page_title="Validador DR vs PR", layout="wide")

st.title("Validador de Demanda e Produção v4.1 🚀 (Cópia Literal)")

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
                st.session_state['df_dr'] = df_dr.groupby(chaves_agrupamento, dropna=False)[meses].sum().reset_index()
                
                st.success(f"Aba '{aba_selecionada}' processada com sucesso!")
                st.dataframe(st.session_state['df_dr'])
                
        except Exception as e:
            st.error(f"Ocorreu um erro na Etapa 1: {e}")

with aba2:
    st.subheader("Etapa 2: Consolidação Bruta e Direta do PR")
    st.write("Faça o upload dos arquivos juntos.")
    
    arquivos_pr = st.file_uploader("Upload dos arquivos Excel (PR)", type=["xlsx", "xls", "xlsm"], accept_multiple_files=True, key="upload_pr")
    
    if arquivos_pr:
        lista_pr_bruto = []
        lista_pr_resumo = []
        erros_pr = []
        
        # A VARIÁVEL MÁGICA: Vai guardar o cabeçalho do 1º arquivo
        cabecalho_oficial = None 
        
        for arq in arquivos_pr:
            try:
                xls_pr = pd.ExcelFile(arq)
                abas_pr = xls_pr.sheet_names
                aba_alvo = next((aba for aba in abas_pr if "production request" in aba.lower()), abas_pr[0])
                
                # Lê de B até W, considerando a linha 4 do Excel (header=3 no Python) como cabeçalho
                df_raw = pd.read_excel(xls_pr, sheet_name=aba_alvo, header=3, usecols="B:W")
                
                # Definição do Range: Apaga as linhas do final onde as colunas B até G estão vazias
                df_raw.dropna(subset=df_raw.columns[0:6], how='all', inplace=True)
                
                # --- A LÓGICA DO CABEÇALHO MESTRE ---
                if cabecalho_oficial is None:
                    # Se for o primeiro arquivo, salva os nomes das colunas como oficiais
                    # Tratamento rápido apenas para evitar colunas com o exato mesmo nome
                    cols_limpas = []
                    vistos = set()
                    for c in df_raw.columns.astype(str):
                        novo_c = c.strip()
                        contador = 1
                        while novo_c in vistos:
                            novo_c = f"{c.strip()}_{contador}"
                            contador += 1
                        vistos.add(novo_c)
                        cols_limpas.append(novo_c)
                    
                    cabecalho_oficial = cols_limpas
                
                # FORÇA O ARQUIVO A USAR O CABEÇALHO DO PRIMEIRO (Cola como valor embaixo)
                df_raw.columns = cabecalho_oficial
                
                # Adiciona a origem
                df_raw['Arquivo_Origem'] = arq.name
                lista_pr_bruto.append(df_raw.copy())
                
                # --- EXTRAÇÃO PARA COMPARAÇÃO PELA POSIÇÃO FÍSICA ---
                # Como forçamos todos os cabeçalhos a serem iguais, a posição não muda.
                df_resumo_temp = pd.DataFrame()
                df_resumo_temp['Planta']  = df_raw.iloc[:, 1]  # Coluna C
                df_resumo_temp['Marca']   = df_raw.iloc[:, 4]  # Coluna F
                df_resumo_temp['Mercado'] = df_raw.iloc[:, 5]  # Coluna G
                df_resumo_temp['Produto'] = df_raw.iloc[:, 6]  # Coluna H
                df_resumo_temp['Série']   = df_raw.iloc[:, 7]  # Coluna I
                
                # Meses (Q=15, R=16, S=17, T=18, U=19, V=20)
                df_resumo_temp['Jul'] = df_raw.iloc[:, 15]
                df_resumo_temp['Ago'] = df_raw.iloc[:, 16]
                df_resumo_temp['Set'] = df_raw.iloc[:, 17]
                df_resumo_temp['Out'] = df_raw.iloc[:, 18]
                df_resumo_temp['Nov'] = df_raw.iloc[:, 19]
                df_resumo_temp['Dez'] = df_raw.iloc[:, 20]
                
                lista_pr_resumo.append(df_resumo_temp)
                
            except Exception as e:
                erros_pr.append(f"Erro no arquivo {arq.name}: {e}")
        
        if erros_pr:
            for erro in erros_pr:
                st.error(erro)
                
        if lista_pr_resumo:
            # --- SALVA O BRUTO COMPLETO ALINHADO ---
            df_pr_bruto_final = pd.concat(lista_pr_bruto, ignore_index=True)
            st.session_state['df_pr_bruto'] = df_pr_bruto_final 
            
            # --- PROCESSA O RESUMO ---
            df_pr_full = pd.concat(lista_pr_resumo, ignore_index=True)
            
            df_pr_full['Produto'] = df_pr_full['Produto'].astype(str).str.strip().str.upper()
            df_pr_full['Planta'] = df_pr_full['Planta'].astype(str).str.strip().str.upper()
            
            mask_produtos = df_pr_full['Produto'].isin(produtos_alvo)
            mask_planta = df_pr_full['Planta'] != 'GENERAL RODRIGUEZ'
            
            df_pr_filtrado = df_pr_full[mask_produtos & mask_planta].copy()
            
            df_pr_filtrado['Mercado'] = df_pr_filtrado['Mercado'].astype(str).str.strip().str.upper().replace(de_para_mercados)
            df_pr_filtrado['Marca'] = df_pr_filtrado['Marca'].astype(str).str.strip().str.upper().replace(de_para_marcas)
            
            for mes in meses_comparacao:
                df_pr_filtrado[mes] = pd.to_numeric(df_pr_filtrado[mes], errors='coerce').fillna(0)
            
            chaves = ['Marca', 'Mercado', 'Produto', 'Série']
            df_pr_resumo_final = df_pr_filtrado.groupby(chaves, dropna=False)[meses_comparacao].sum().reset_index()
            df_pr_resumo_final['Total PR'] = df_pr_resumo_final[meses_comparacao].sum(axis=1)
            
            st.session_state['df_pr'] = df_pr_resumo_final
            
            st.success(f"{len(lista_pr_bruto)} arquivos consolidados e perfeitamente empilhados!")
            st.dataframe(st.session_state['df_pr'])

with aba3:
    st.subheader("Etapa 3: Resultado da Comparação (Visão Executiva)")
    
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
        
        df_dif_resumo_excel = df_merge.groupby(['Marca', 'Mercado', 'Produto'], dropna=False)[colunas_diferenca].sum().reset_index()
        df_dif_detalhada_excel = df_merge[['Marca', 'Mercado', 'Produto', 'Série'] + colunas_diferenca]
        
        df_dif_tela = df_merge.groupby(['Marca', 'Mercado', 'Produto'], dropna=False)['Dif_Total'].sum().reset_index()
        df_dif_tela = df_dif_tela[df_dif_tela['Dif_Total'] != 0].copy()
        
        if df_dif_tela.empty:
            st.success("🎉 TUDO OK! Os números batem perfeitamente. Nenhuma diferença encontrada no Total de Julho a Dezembro.")
        else:
            st.warning("Atenção: Diferenças encontradas. Baixe o Excel para o detalhamento.")
            
            df_dif_tela['Email'] = email_padrao_teams 
            df_dif_tela['Follow-up'] = df_dif_tela.apply(
                lambda row: gerar_link_teams(row['Email'], row['Marca'], row['Mercado'], row['Produto'], row['Dif_Total']), 
                axis=1
            )
            
            st.dataframe(
                df_dif_tela, 
                hide_index=True,
                column_config={
                    "Email": None,
                    "Follow-up": st.column_config.LinkColumn("Ação Sugerida", display_text="💬 Enviar Alerta")
                }
            )
            
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df_dr_final.to_excel(writer, index=False, sheet_name='DR')
            
            if not df_pr_bruto_export.empty:
                df_pr_bruto_export.to_excel(writer, index=False, sheet_name='PR_Bruto_Completo')
                
            df_pr_final.to_excel(writer, index=False, sheet_name='PR_Resumo_Consolidado')
            df_dif_resumo_excel.to_excel(writer, index=False, sheet_name='Dif_Resumo')
            df_dif_detalhada_excel.to_excel(writer, index=False, sheet_name='Dif_Detalhada')
            
        st.markdown("---")
        st.download_button(
            label="📥 Baixar Análise Completa em Excel",
            data=buffer.getvalue(),
            file_name="Analise_DR_vs_PR_v4-1.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.info("👈 Por favor, carregue os arquivos nas Abas 1 (DR) e 2 (PR) para visualizar o cruzamento.")
