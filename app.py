import streamlit as st
import pandas as pd
import io
import urllib.parse

# Configuração da página
st.set_page_config(page_title="Validador DR vs PR", layout="wide")

st.title("Validador de Demanda e Produção v2.1 🚀")

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
                
                st.session_state['df_dr'] = df_dr.groupby(['Marca', 'Mercado', 'Produto', 'Série'])[meses].sum().reset_index()
                
                st.success(f"Aba '{aba_selecionada}' processada com sucesso!")
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
        
        # --- CABEÇALHO CLARO E EXPLICATIVO (O QUE O USUÁRIO VÊ NA TELA E NO EXCEL) ---
        colunas_absolutas = [
            'B_Ciclo', 'C_Planta', 'D_Código', 'E_FC', 
            'F_Marca', 'G_Mercado', 'H_Produto', 'I_Série',
            'J_Jan', 'K_Fev', 'L_Mar', 'M_Abr', 'N_Mai', 'O_Jun', 'P_Sem1',
            'Q_Julho', 'R_Agosto', 'S_Setembro', 'T_Outubro', 'U_Novembro', 'V_Dezembro', 'W_Total_Ano'
        ]
        
        for arq in arquivos_pr:
            try:
                xls_pr = pd.ExcelFile(arq)
                abas_pr = xls_pr.sheet_names
                aba_alvo = next((aba for aba in abas_pr if "production request" in aba.lower()), abas_pr[0])
                
                # Leitura Cirúrgica (Carrega como um bloco bruto de B até W)
                df_raw = pd.read_excel(xls_pr, sheet_name=aba_alvo, header=None, usecols="B:W")
                
                # Garante que a matriz tem exatas 22 colunas, preenchendo o que faltar
                while len(df_raw.columns) < 22:
                    df_raw[len(df_raw.columns)] = None

                # NOMEIA AS COLUNAS PRIMEIRO (Evita que o Python erre a mira)
                df_raw.columns = colunas_absolutas
                df_raw['Arquivo_Origem'] = arq.name

                # O RADAR (Buscando onde os dados começam com base na palavra PRODUTO na coluna H)
                linha_inicio_dados = 4
                for i in range(min(20, len(df_raw))):
                    val_h = str(df_raw['H_Produto'].iloc[i]).strip().upper()
                    if 'PRODUTO' in val_h or 'PROD' in val_h:
                        linha_inicio_dados = i + 1
                        break
                
                # Recorta apenas os dados
                df_dados = df_raw.iloc[linha_inicio_dados:].copy()
                
                # --- O TRUQUE PARA SALVAR AS CÉLULAS MESCLADAS ---
                # Arrasta o nome da Planta, Marca, Mercado e Produto para as linhas em branco abaixo delas
                colunas_para_preencher = ['C_Planta', 'F_Marca', 'G_Mercado', 'H_Produto']
                df_dados[colunas_para_preencher] = df_dados[colunas_para_preencher].ffill()
                
                # Remove apenas as linhas que estão 100% em branco nos códigos
                df_dados.dropna(subset=['H_Produto', 'I_Série'], how='all', inplace=True)
                
                # --- FILTROS DE NEGÓCIO ---
                df_dados['H_Produto'] = df_dados['H_Produto'].astype(str).str.strip().str.upper()
                df_dados['C_Planta'] = df_dados['C_Planta'].astype(str).str.strip().str.upper()
                
                mask_produtos = df_dados['H_Produto'].isin(produtos_alvo)
                mask_planta = df_dados['C_Planta'] != 'GENERAL RODRIGUEZ'
                
                df_bruto = df_dados[mask_produtos & mask_planta].copy()
                
                if df_bruto.empty:
                    continue
                
                # --- BORRACHA E MATEMÁTICA ---
                # Limpa meses antigos
                col_jan_jun = ['J_Jan', 'K_Fev', 'L_Mar', 'M_Abr', 'N_Mai', 'O_Jun', 'P_Sem1']
                df_bruto.loc[:, col_jan_jun] = None
                
                # Força números puros no semestre 2
                meses_semestre2 = ['Q_Julho', 'R_Agosto', 'S_Setembro', 'T_Outubro', 'U_Novembro', 'V_Dezembro']
                for mes in meses_semestre2:
                    df_bruto[mes] = pd.to_numeric(df_bruto[mes], errors='coerce').fillna(0)
                
                # Recalcula o Total da coluna W do zero
                df_bruto['W_Total_Ano'] = df_bruto[meses_semestre2].sum(axis=1)
                
                lista_pr_bruto.append(df_bruto)
                
                # --- PREPARA O RESUMO PARA A COMPARAÇÃO (Etapa 3) ---
                df_resumo_temp = df_bruto[['F_Marca', 'G_Mercado', 'H_Produto', 'I_Série'] + meses_semestre2].copy()
                
                de_para_nomes = {
                    'F_Marca': 'Marca',
                    'G_Mercado': 'Mercado',
                    'H_Produto': 'Produto',
                    'I_Série': 'Série',
                    'Q_Julho': 'Jul',
                    'R_Agosto': 'Ago',
                    'S_Setembro': 'Set',
                    'T_Outubro': 'Out',
                    'U_Novembro': 'Nov',
                    'V_Dezembro': 'Dez'
                }
                df_resumo_temp.rename(columns=de_para_nomes, inplace=True)
                lista_pr_resumo.append(df_resumo_temp)
                
            except Exception as e:
                erros_pr.append(f"Erro no arquivo {arq.name}: {e}")
        
        if erros_pr:
            for erro in erros_pr:
                st.error(erro)
                
        if lista_pr_resumo:
            # Consolida o Resumo Final
            df_pr_full = pd.concat(lista_pr_resumo, ignore_index=True)
            df_pr_full['Mercado'] = df_pr_full['Mercado'].astype(str).str.strip().str.upper().replace(de_para_mercados)
            df_pr_full['Marca'] = df_pr_full['Marca'].astype(str).str.strip().str.upper().replace(de_para_marcas)
            
            df_pr_resumo_final = df_pr_full.groupby(['Marca', 'Mercado', 'Produto', 'Série'])[meses_comparacao].sum().reset_index()
            df_pr_resumo_final['Total PR'] = df_pr_resumo_final[meses_comparacao].sum(axis=1)
            
            # Consolida o Bruto Final
            if lista_pr_bruto:
                df_pr_bruto_final = pd.concat(lista_pr_bruto, ignore_index=True)
                st.session_state['df_pr_bruto'] = df_pr_bruto_final 
            
            st.session_state['df_pr'] = df_pr_resumo_final
            
            st.success(f"{len(lista_pr_bruto)} arquivos consolidados! Colunas formatadas e mesclas desfeitas com sucesso.")
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
            file_name="Analise_DR_vs_PR_v2-1.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.info("👈 Por favor, carregue os arquivos nas Abas 1 (DR) e 2 (PR) para visualizar o cruzamento.")
