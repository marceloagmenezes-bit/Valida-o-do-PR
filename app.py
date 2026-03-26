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
    st.subheader("Etapa 2: Nova Consolidação Inteligente dos Arquivos PR")
    arquivos_pr = st.file_uploader("Upload dos arquivos Excel (PR)", type=["xlsx", "xls", "xlsm"], accept_multiple_files=True, key="upload_pr")
    
    if arquivos_pr:
        lista_pr_bruto = []
        lista_pr_resumo = []
        erros_pr = []
        
        for arq in arquivos_pr:
            try:
                xls_pr = pd.ExcelFile(arq)
                aba_alvo = next((aba for aba in xls_pr.sheet_names if "production request" in aba.lower()), xls_pr.sheet_names[0])
                
                # Lê o quadrado B:W bruto, sem pular linhas fixas
                df_raw = pd.read_excel(xls_pr, sheet_name=aba_alvo, header=None, usecols="B:W")
                
                # O "Caçador": Procura em qual linha está escrito "Produto" e "Planta" para usar como cabeçalho
                header_idx = -1
                for i in range(min(15, len(df_raw))):
                    linha_texto = " ".join([str(x).upper() for x in df_raw.iloc[i].values])
                    if 'PRODUTO' in linha_texto and 'PLANTA' in linha_texto:
                        header_idx = i
                        break
                
                if header_idx == -1:
                    erros_pr.append(f"Cabeçalho não encontrado no arquivo {arq.name}.")
                    continue
                
                # Define o cabeçalho exato que estava no arquivo e corta os dados dali para baixo
                nomes_originais = df_raw.iloc[header_idx].values
                nomes_colunas = []
                # Cria nomes únicos para evitar conflitos no pandas
                for c in nomes_originais:
                    nome = str(c).strip()
                    if nome == 'nan' or nome == 'None' or nome == '': nome = 'Info_Extra'
                    # Garante que não repita nome
                    if nome in nomes_colunas: nome = nome + "_2"
                    nomes_colunas.append(nome)
                    
                df_dados = df_raw.iloc[header_idx + 1:].copy()
                df_dados.columns = nomes_colunas
                
                # Limpa linhas 100% vazias
                if 'Produto' in df_dados.columns:
                    df_dados = df_dados.dropna(subset=['Produto'], how='all')
                
                # Filtra apenas o que importa
                df_dados['Produto'] = df_dados['Produto'].astype(str).str.strip().str.upper()
                df_dados['Planta'] = df_dados['Planta'].astype(str).str.strip().str.upper()
                
                mask_produtos = df_dados['Produto'].isin(produtos_alvo)
                mask_planta = df_dados['Planta'] != 'GENERAL RODRIGUEZ'
                
                df_bruto = df_dados[mask_produtos & mask_planta].copy()
                
                if df_bruto.empty:
                    continue
                
                df_bruto['Arquivo_Origem'] = arq.name
                
                # --- IDENTIFICA E LIMPA OS MESES AUTOMATICAMENTE ---
                cols_para_soma = []
                col_total = None
                
                for col in df_bruto.columns:
                    col_str = str(col).lower()
                    
                    # Identifica se a coluna é o TOTAL
                    if 'total' in col_str:
                        col_total = col
                    
                    # Identifica se é mês do SEGUNDO semestre (Jul-Dez)
                    elif '07-01' in col_str or '2026-07' in col_str: cols_para_soma.append((col, 'Jul'))
                    elif '08-01' in col_str or '2026-08' in col_str: cols_para_soma.append((col, 'Ago'))
                    elif '09-01' in col_str or '2026-09' in col_str: cols_para_soma.append((col, 'Set'))
                    elif '10-01' in col_str or '2026-10' in col_str: cols_para_soma.append((col, 'Out'))
                    elif '11-01' in col_str or '2026-11' in col_str: cols_para_soma.append((col, 'Nov'))
                    elif '12-01' in col_str or '2026-12' in col_str: cols_para_soma.append((col, 'Dez'))
                    
                    # Zera meses antigos (Jan-Jun)
                    elif any(m in col_str for m in ['2026-01', '2026-02', '2026-03', '2026-04', '2026-05', '2026-06']):
                        df_bruto[col] = 0
                
                # Converte os meses do 2º semestre para números
                colunas_reais_dos_meses = [c[0] for c in cols_para_soma]
                for mes_col in colunas_reais_dos_meses:
                    df_bruto[mes_col] = pd.to_numeric(df_bruto[mes_col], errors='coerce').fillna(0)
                
                # Recalcula o total apenas com a soma exata de Jul a Dez
                if col_total:
                    df_bruto[col_total] = df_bruto[colunas_reais_dos_meses].sum(axis=1)
                
                # Renomeia os cabeçalhos para ficarem bonitos na aba Bruta
                rename_dict = {c[0]: c[1] for c in cols_para_soma}
                df_bruto.rename(columns=rename_dict, inplace=True)
                
                lista_pr_bruto.append(df_bruto)
                
                # --- PREPARA O RESUMO (Etapa 3) ---
                df_resumo_temp = pd.DataFrame()
                
                df_resumo_temp['Marca'] = df_bruto['Marca'] if 'Marca' in df_bruto.columns else 'N/A'
                df_resumo_temp['Mercado'] = df_bruto['Mercado'] if 'Mercado' in df_bruto.columns else 'N/A'
                df_resumo_temp['Produto'] = df_bruto['Produto']
                df_resumo_temp['Série'] = df_bruto['Série'] if 'Série' in df_bruto.columns else 'N/A'
                
                for col_original, mes_nome in cols_para_soma:
                    df_resumo_temp[mes_nome] = df_bruto[mes_nome]
                    
                lista_pr_resumo.append(df_resumo_temp)
                
            except Exception as e:
                erros_pr.append(f"Erro no arquivo {arq.name}: {e}")
        
        if erros_pr:
            for erro in erros_pr:
                st.error(erro)
                
        if lista_pr_resumo:
            df_pr_full = pd.concat(lista_pr_resumo, ignore_index=True)
            df_pr_full['Mercado'] = df_pr_full['Mercado'].astype(str).str.strip().str.upper().replace(de_para_mercados)
            df_pr_full['Marca'] = df_pr_full['Marca'].astype(str).str.strip().str.upper().replace(de_para_marcas)
            
            chaves = ['Marca', 'Mercado', 'Produto', 'Série']
            df_pr_full[chaves] = df_pr_full[chaves].fillna('N/A')
            
            for mes in meses_comparacao:
                if mes not in df_pr_full.columns:
                    df_pr_full[mes] = 0
                else:
                    df_pr_full[mes] = pd.to_numeric(df_pr_full[mes], errors='coerce').fillna(0)
            
            df_pr_resumo_final = df_pr_full.groupby(chaves)[meses_comparacao].sum().reset_index()
            df_pr_resumo_final['Total PR'] = df_pr_resumo_final[meses_comparacao].sum(axis=1)
            
            if lista_pr_bruto:
                df_pr_bruto_final = pd.concat(lista_pr_bruto, ignore_index=True)
                st.session_state['df_pr_bruto'] = df_pr_bruto_final 
            
            st.session_state['df_pr'] = df_pr_resumo_final
            
            st.success(f"{len(lista_pr_bruto)} arquivos consolidados com a Nova Engenharia V3.0! Cálculo perfeito estabelecido.")
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
        
        df_dif_resumo_excel = df_merge.groupby(['Marca', 'Mercado', 'Produto'])[colunas_diferenca].sum().reset_index()
        df_dif_detalhada_excel = df_merge[['Marca', 'Mercado', 'Produto', 'Série'] + colunas_diferenca]
        
        df_dif_tela = df_merge.groupby(['Marca', 'Mercado', 'Produto'])['Dif_Total'].sum().reset_index()
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
            file_name="Analise_DR_vs_PR_v3-0.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.info("👈 Por favor, carregue os arquivos nas Abas 1 (DR) e 2 (PR) para visualizar o cruzamento.")
