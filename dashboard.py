import streamlit as st
import pandas as pd
import plotly.express as px
import os

# Configuração da página para ficar larga estilo Power BI
st.set_page_config(page_title="Auditoria de Coletas", layout="wide")

# --- SISTEMA DE LOGIN ---
def check_password():
    def password_entered():
        # Tenta pegar a senha do painel de Secrets da Nuvem, se não achar usa a senha fixa
        senha_correta = st.secrets.get("senha_painel", "e-recoli2026")
        if st.session_state["password"] == senha_correta:
            st.session_state["password_correct"] = True
            del st.session_state["password"]
        else:
            st.session_state["password_correct"] = False

    if "password_correct" not in st.session_state:
        st.markdown("### 🔒 Acesso Restrito Homedock")
        st.text_input("🔑 Digite a senha para acessar o painel:", type="password", on_change=password_entered, key="password")
        return False
    elif not st.session_state["password_correct"]:
        st.markdown("### 🔒 Acesso Restrito Homedock")
        st.text_input("🔑 Digite a senha para acessar o painel:", type="password", on_change=password_entered, key="password")
        st.error("😕 Senha incorreta. Tente novamente.")
        return False
    else:
        return True

if not check_password():
    st.stop()
# --- FIM DO SISTEMA DE LOGIN ---

st.title("📊 Auditoria de Dados")
st.markdown("Estudo de coleta de mercadoria Homedock")

# Função para carregar os dados
@st.cache_data(show_spinner="Carregando dados consolidados...")
def load_data(mtime):
    try:
        script_dir = os.path.dirname(os.path.abspath(__file__))
        data_path = os.path.join(script_dir, 'Consolidado_Faturas_Coletas.xlsx')
        if not os.path.exists(data_path):
            return pd.DataFrame()
        
        df = pd.read_excel(data_path)
        
        # Padronizar o nome das transportadoras
        df['NOME DA TRANSPORTADORA'] = df['NOME DA TRANSPORTADORA'].fillna('NÃO ENCONTRADA').astype(str).str.strip().str.upper()
        df.loc[df['NOME DA TRANSPORTADORA'] == 'NAN', 'NOME DA TRANSPORTADORA'] = 'NÃO ENCONTRADA'
        
        # Unificar todas as variações de E-RECOLI
        df.loc[df['NOME DA TRANSPORTADORA'].str.contains('RECOLI', na=False), 'NOME DA TRANSPORTADORA'] = 'E-RECOLI'
        
        # Criar a coluna 'Ano' a partir da Data da Coleta
        data_temp = pd.to_datetime(df['DATA DA COLETA'], errors='coerce')
        df['ANO'] = data_temp.dt.year.fillna(-1).astype(int).astype(str)
        df['ANO'] = df['ANO'].replace('-1', 'Sem Data')
        
        # Mapeamento para meses em português
        meses_pt = {'01': 'Jan', '02': 'Fev', '03': 'Mar', '04': 'Abr', '05': 'Mai', '06': 'Jun', 
                    '07': 'Jul', '08': 'Ago', '09': 'Set', '10': 'Out', '11': 'Nov', '12': 'Dez'}
                    
        # Criar a coluna 'MÊS_COMPLETO' para ordenação (YYYY-MM)
        if 'DATA DE VENCIMENTO' in df.columns:
            data_venc = pd.to_datetime(df['DATA DE VENCIMENTO'], format='%d/%m/%Y', errors='coerce')
            df['MÊS_COMPLETO'] = data_venc.dt.strftime('%Y-%m').fillna('Sem Data')
            meses_str = data_venc.dt.strftime('%m')
            anos_str = data_venc.dt.year.fillna(0).astype(int).astype(str).str[-2:]
            df['MÊS'] = meses_str.map(meses_pt).fillna('Sem Data') + '/' + anos_str
            df.loc[df['MÊS'].str.startswith('Sem Data'), 'MÊS'] = 'Sem Data'
        else:
            df['MÊS_COMPLETO'] = data_temp.dt.strftime('%Y-%m').fillna('Sem Data')
            meses_str = data_temp.dt.strftime('%m')
            anos_str = data_temp.dt.year.fillna(0).astype(int).astype(str).str[-2:]
            df['MÊS'] = meses_str.map(meses_pt).fillna('Sem Data') + '/' + anos_str
            df.loc[df['MÊS'].str.startswith('Sem Data'), 'MÊS'] = 'Sem Data'
        
        if 'STATUS' in df.columns:
            df['STATUS'] = df['STATUS'].fillna('Sem Status').astype(str)
            df.loc[df['STATUS'].str.strip() == '', 'STATUS'] = 'Sem Status'
            df.loc[df['STATUS'].str.lower() == 'nan', 'STATUS'] = 'Sem Status'
            
            # Nova regra: se "Sem Status" e tem Data da Coleta -> "Falta Devolver"
            # Se "Sem Status" e NÃO tem Data da Coleta -> "Falta Coletar"
            mask_sem_status = df['STATUS'] == 'Sem Status'
            mask_com_data = data_temp.notna()
            
            df.loc[mask_sem_status & mask_com_data, 'STATUS'] = 'Falta Devolver'
            df.loc[mask_sem_status & ~mask_com_data, 'STATUS'] = 'Falta Coletar'
        
        return df
    except Exception as e:
        st.error(f"Erro ao carregar o arquivo: {e}")
        return pd.DataFrame()

# Checar o horário de modificação do arquivo para invalidar o cache automaticamente se o arquivo mudar
script_dir = os.path.dirname(os.path.abspath(__file__))
consolidado_path = os.path.join(script_dir, 'Consolidado_Faturas_Coletas.xlsx')

try:
    arquivo_mtime = os.path.getmtime(consolidado_path)
except OSError:
    arquivo_mtime = 0

df = load_data(arquivo_mtime)

if not df.empty:
    # --- BARRA LATERAL (SIDEBAR) ESTILO POWER BI ---
    st.sidebar.header("🔍 Filtros")
    
    # --- VISÕES RÁPIDAS SALVAS ---
    st.sidebar.markdown("### 📌 Visões Rápidas")
    
    if "visao_erecoli_ativa" not in st.session_state:
        st.session_state["visao_erecoli_ativa"] = False

    def toggle_visao_erecoli():
        if not st.session_state["visao_erecoli_ativa"]:
            st.session_state["dup_filter"] = "Mostrar APENAS Cobranças Válidas (Para Pagamento)"
            st.session_state["transp_filter"] = [t for t in ['E-RECOLI', 'E-REVERSA'] if t in df['NOME DA TRANSPORTADORA'].unique().tolist()]
            st.session_state["ano_filter"] = [a for a in ['2024', '2025', '2026'] if a in df['ANO'].unique().tolist()]
            st.session_state["mes_filter"] = []
            st.session_state["uf_filter"] = []
            if 'STATUS' in df.columns:
                 st.session_state["status_filter"] = [s for s in ['Falta Coletar', 'Falta Devolver', 'RETORNOU CD'] if s in df['STATUS'].dropna().astype(str).unique().tolist()]
            st.session_state["visao_erecoli_ativa"] = True
        else:
            st.session_state["dup_filter"] = "Mostrar Todas as Notas"
            st.session_state["transp_filter"] = []
            st.session_state["ano_filter"] = []
            st.session_state["mes_filter"] = []
            st.session_state["uf_filter"] = []
            st.session_state["status_filter"] = []
            st.session_state["visao_erecoli_ativa"] = False

    lbl_btn = "❌ Limpar Visão E-RECOLI" if st.session_state.get("visao_erecoli_ativa") else "💰 Auditoria E-RECOLI"
    hlp_btn = "Remove os filtros aplicados." if st.session_state.get("visao_erecoli_ativa") else "Aplica os filtros para ver as cobranças indevidas e o valor total para pagamento da E-Recoli."
    st.sidebar.button(lbl_btn, on_click=toggle_visao_erecoli, use_container_width=True, help=hlp_btn)
    st.sidebar.markdown("---")
    
    # Botão para atualizar a base de dados
    with st.sidebar.expander("📂 Atualizar Base de Dados"):
        st.markdown("<small>Faça o upload dos novos arquivos para atualizar os dados.</small>", unsafe_allow_html=True)
        upload_coletas_sb = st.file_uploader("Tracking Coletas", type=["xlsx", "xls"], key="up_col_sb")
        upload_faturas_sb = st.file_uploader("Resumo Faturas", type=["xlsx", "xls"], key="up_fat_sb")
        
        if st.button("🔄 Processar Novas Planilhas", use_container_width=True, type="primary"):
            if upload_coletas_sb and upload_faturas_sb:
                with st.spinner("Lendo planilhas originais e reconstruindo a base... Aguarde."):
                    try:
                        from merge_script import consolidate_data
                        df_novo = consolidate_data(faturas_file=upload_faturas_sb, coletas_file=upload_coletas_sb)
                        df_novo.to_excel(consolidado_path, index=False)
                        st.cache_data.clear()
                        st.success("Atualizado!")
                        st.rerun()
                    except Exception as e:
                        import traceback
                        st.error(f"Erro ao atualizar: {e}")
                        st.text(traceback.format_exc())
            else:
                st.warning("Por favor, carregue os DOIS arquivos antes de clicar.")

    # Campo de busca livre
    busca_nf = st.sidebar.text_input("🔎 Buscar Fatura / NF / Ticket:", "", key="busca_filter")
    
    st.sidebar.markdown("---")
    
    # Filtro Exclusivo de Duplicadas (Radio)
    filtro_duplicadas = st.sidebar.radio(
        "🚨 Filtro de Duplicidade:",
        [
            "Mostrar Todas as Notas", 
            "Mostrar APENAS Cobranças Válidas (Para Pagamento)",
            "Mostrar APENAS as Repetidas (Valor Excedente)"
        ],
        key="dup_filter"
    )
    
    # Agrupando os outros filtros em um expander para limpar a sidebar
    with st.sidebar.expander("Filtros Avançados", expanded=False):
        transportadoras_lista = sorted(df['NOME DA TRANSPORTADORA'].unique().tolist())
        transp_selecionada = st.multiselect("Transportadora:", transportadoras_lista, key="transp_filter", help="Deixe vazio para ver Todas")
        
        anos_lista = sorted(df['ANO'].unique().tolist())
        ano_selecionado = st.multiselect("Ano da Coleta:", anos_lista, key="ano_filter", help="Deixe vazio para ver Todos")
        
        if 'MÊS' in df.columns:
            # Ordenar os meses cronologicamente para o filtro
            meses_ordenados_df = df[['MÊS_COMPLETO', 'MÊS']].drop_duplicates().sort_values('MÊS_COMPLETO')
            mes_lista = meses_ordenados_df['MÊS'].tolist()
            mes_selecionado = st.multiselect("Mês:", mes_lista, key="mes_filter", help="Deixe vazio para ver Todos")
        else:
            mes_selecionado = []
            
        if 'UF' in df.columns:
            uf_lista = sorted(df['UF'].dropna().astype(str).unique().tolist())
            uf_selecionado = st.multiselect("UF:", uf_lista, key="uf_filter", help="Deixe vazio para ver Todas")
        else:
            uf_selecionado = []
        
        if 'STATUS' in df.columns:
            status_lista = sorted(df['STATUS'].dropna().astype(str).unique().tolist())
            status_selecionado = st.multiselect("Status:", status_lista, key="status_filter", help="Deixe vazio para ver Todos")
        else:
            status_selecionado = []
    
    # --- APLICANDO OS FILTROS ESTÁTICOS DA SIDEBAR ---
    df_filtrado = df.copy()
    
    if transp_selecionada:
        df_filtrado = df_filtrado[df_filtrado['NOME DA TRANSPORTADORA'].isin(transp_selecionada)]
    if ano_selecionado:
        df_filtrado = df_filtrado[df_filtrado['ANO'].isin(ano_selecionado)]
    if mes_selecionado:
        df_filtrado = df_filtrado[df_filtrado['MÊS'].isin(mes_selecionado)]
    if uf_selecionado:
        df_filtrado = df_filtrado[df_filtrado['UF'].isin(uf_selecionado)]
    if status_selecionado:
        df_filtrado = df_filtrado[df_filtrado['STATUS'].isin(status_selecionado)]
        
    if busca_nf:
        busca_nf = busca_nf.strip().lower()
        ticket_mask = False
        if 'TICKET' in df_filtrado.columns:
            ticket_mask = df_filtrado['TICKET'].astype(str).str.lower().str.contains(busca_nf)
            
        mask = (
            df_filtrado['NUMERO DA FATURA'].astype(str).str.lower().str.contains(busca_nf) |
            df_filtrado['NOTA FISCAL'].astype(str).str.lower().str.contains(busca_nf) |
            ticket_mask
        )
        df_filtrado = df_filtrado[mask]
        
    # Ordenar por vencimento mais antigo
    if 'DATA DE VENCIMENTO' in df_filtrado.columns:
        temp_date = pd.to_datetime(df_filtrado['DATA DE VENCIMENTO'], format='%d/%m/%Y', errors='coerce')
        df_filtrado['TEMP_DATE'] = temp_date
        df_filtrado = df_filtrado.sort_values(by='TEMP_DATE', ascending=True, na_position='last').drop(columns=['TEMP_DATE'])

    # Calcular o valor de duplicidade real antes dos filtros de exibição afetarem as linhas originais
    valor_duplicado_real = 0.0
    if 'VALOR FRETE' in df_filtrado.columns:
        mask_excesso = df_filtrado.duplicated(subset=['NOTA FISCAL'], keep='first')
        valor_duplicado_real = df_filtrado.loc[mask_excesso, 'VALOR FRETE'].sum()
        
    # Aplicar o filtro de Duplicidade selecionado pelo usuário
    if filtro_duplicadas == "Mostrar APENAS as Repetidas (Valor Excedente)":
        duplicadas_mask = df_filtrado.duplicated(subset=['NOTA FISCAL'], keep='first')
        df_filtrado = df_filtrado[duplicadas_mask]
    elif filtro_duplicadas == "Mostrar APENAS Cobranças Válidas (Para Pagamento)":
        validas_mask = ~df_filtrado.duplicated(subset=['NOTA FISCAL'], keep='first')
        df_filtrado = df_filtrado[validas_mask]
        
    # Lendo seleções cruzadas dos gráficos
    sel_mes = st.session_state.get("chart_mes", {}).get("selection", {}).get("points", [])
    sel_uf = st.session_state.get("chart_uf", {}).get("selection", {}).get("points", [])
    sel_transp = st.session_state.get("chart_transp", {}).get("selection", {}).get("points", [])
    
    # Criar um dataframe final que obedece tanto aos filtros da sidebar quanto aos cliques nos gráficos
    df_plot = df_filtrado.copy()
    
    mes_clicados = [p['x'] for p in sel_mes] if sel_mes else []
    uf_clicadas = [p['y'] for p in sel_uf] if sel_uf else []
    transp_clicadas = [p['y'] for p in sel_transp] if sel_transp else []
    
    if mes_clicados:
        df_plot = df_plot[df_plot['MÊS'].isin(mes_clicados)]
    if uf_clicadas:
        df_plot = df_plot[df_plot['UF'].isin(uf_clicadas)]
    if transp_clicadas:
        df_plot = df_plot[df_plot['NOME DA TRANSPORTADORA'].isin(transp_clicadas)]

    # --- ABAS PRINCIPAIS ---
    tab_analise, tab_auditoria = st.tabs(["📈 Visão Analítica", "📑 Auditoria Detalhada"])

    with tab_analise:
        # -- CARDS DE KPI --
        valor_total = df_plot['VALOR FRETE'].sum() if 'VALOR FRETE' in df_plot.columns else 0.0
        qtd_notas = len(df_plot)
        qtd_faturas = df_plot['NUMERO DA FATURA'].nunique() if 'NUMERO DA FATURA' in df_plot.columns else 0
        ticket_medio = valor_total / qtd_faturas if qtd_faturas > 0 else 0.0
        
        col_kpi1, col_kpi2, col_kpi3, col_kpi4 = st.columns(4)
        col_kpi1.metric("Valor Total do Frete", f"R$ {valor_total:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
        col_kpi2.metric("Qtd de Conhecimentos", f"{qtd_faturas}")
        col_kpi3.metric("Qtd Notas Carregadas", f"{qtd_notas}")
        col_kpi4.metric("Ticket Médio das faturas", f"R$ {ticket_medio:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
        
        st.markdown("<br>", unsafe_allow_html=True)
        
        if 'VALOR FRETE' in df_plot.columns and 'MÊS' in df_plot.columns and 'UF' in df_plot.columns:
            # Gráfico 1: Gastos por Mês
            df_grafico_mes = df_filtrado.copy()
            if uf_clicadas:
                df_grafico_mes = df_grafico_mes[df_grafico_mes['UF'].isin(uf_clicadas)]
            if transp_clicadas:
                df_grafico_mes = df_grafico_mes[df_grafico_mes['NOME DA TRANSPORTADORA'].isin(transp_clicadas)]
                
            # Agrupar e calcular métricas para tooltip
            gastos_mes = df_grafico_mes.groupby(['MÊS_COMPLETO', 'MÊS'], as_index=False).agg(
                VALOR_FRETE=('VALOR FRETE', 'sum'),
                QTD_NOTAS=('NOTA FISCAL', 'count') if 'NOTA FISCAL' in df_grafico_mes.columns else ('VALOR FRETE', 'count'),
                QTD_FATURAS=('NUMERO DA FATURA', 'nunique') if 'NUMERO DA FATURA' in df_grafico_mes.columns else ('VALOR FRETE', 'count')
            )
            gastos_mes = gastos_mes.sort_values(by='MÊS_COMPLETO')
            
            fig_mes = px.bar(
                gastos_mes, 
                x='MÊS', 
                y='VALOR_FRETE',
                title="Análise de Gastos por Mês",
                color_discrete_sequence=['#3b82f6'], # Cor primária do tema
                template="plotly_dark",
                text_auto='.2s',
                hover_data={
                    'MÊS_COMPLETO': False, 
                    'VALOR_FRETE': ':,.2f', 
                    'QTD_FATURAS': True, 
                    'QTD_NOTAS': True
                },
                labels={'VALOR_FRETE': 'Valor Frete (R$)', 'QTD_FATURAS': 'Qtd Faturas (Conhecimentos)', 'QTD_NOTAS': 'Qtd Notas Carregadas'}
            )
            fig_mes.update_layout(plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', margin=dict(l=0, r=0, t=40, b=0), xaxis_title="", yaxis_title="R$")
            st.plotly_chart(fig_mes, use_container_width=True, on_select="rerun", key="chart_mes")
            
            st.markdown("<br>", unsafe_allow_html=True)
            
            # Dividindo a tela em duas colunas para UF e Transportadora
            col_uf, col_transp = st.columns(2)
            
            with col_uf:
                # Gráfico 2: Gastos por UF
                df_grafico_uf = df_filtrado.copy()
                if mes_clicados:
                    df_grafico_uf = df_grafico_uf[df_grafico_uf['MÊS'].isin(mes_clicados)]
                if transp_clicadas:
                    df_grafico_uf = df_grafico_uf[df_grafico_uf['NOME DA TRANSPORTADORA'].isin(transp_clicadas)]
                    
                gastos_uf = df_grafico_uf.groupby('UF', as_index=False).agg(
                    VALOR_FRETE=('VALOR FRETE', 'sum')
                )
                gastos_uf = gastos_uf.sort_values(by='VALOR_FRETE', ascending=True) # Horizontal bar sort
                
                fig_uf = px.bar(
                    gastos_uf, 
                    x='VALOR_FRETE', 
                    y='UF',
                    orientation='h',
                    title="Gastos por UF",
                    color_discrete_sequence=['#ef4444'], # Destaque vermelho
                    template="plotly_dark",
                    text_auto='.2s',
                    labels={'VALOR_FRETE': 'Valor Frete (R$)'}
                )
                
                # Configurar altura: se tiver mais de 7 UFs, criar altura gigante para o gráfico e prender num container pequeno (gera scrollbar)
                if len(gastos_uf) > 7:
                    altura_grafico_uf = len(gastos_uf) * 50 + 100
                    fig_uf.update_layout(height=altura_grafico_uf, plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', margin=dict(l=0, r=0, t=40, b=0), xaxis_title="R$", yaxis_title="")
                    with st.container(height=450):
                        st.plotly_chart(fig_uf, use_container_width=True, on_select="rerun", key="chart_uf")
                else:
                    fig_uf.update_layout(height=450, plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', margin=dict(l=0, r=0, t=40, b=0), xaxis_title="R$", yaxis_title="")
                    st.plotly_chart(fig_uf, use_container_width=True, on_select="rerun", key="chart_uf")
            
            with col_transp:
                # Gráfico 3: Gastos por Transportadora
                df_grafico_transp = df_filtrado.copy()
                if mes_clicados:
                    df_grafico_transp = df_grafico_transp[df_grafico_transp['MÊS'].isin(mes_clicados)]
                if uf_clicadas:
                    df_grafico_transp = df_grafico_transp[df_grafico_transp['UF'].isin(uf_clicadas)]
                    
                gastos_transp = df_grafico_transp.groupby('NOME DA TRANSPORTADORA', as_index=False).agg(
                    VALOR_FRETE=('VALOR FRETE', 'sum')
                )
                gastos_transp = gastos_transp.sort_values(by='VALOR_FRETE', ascending=True) # Horizontal bar sort
                
                fig_transp = px.bar(
                    gastos_transp, 
                    x='VALOR_FRETE', 
                    y='NOME DA TRANSPORTADORA',
                    orientation='h',
                    title="Gastos por Transportadora",
                    color_discrete_sequence=['#10b981'], # Destaque verde esmeralda
                    template="plotly_dark",
                    text_auto='.2s',
                    labels={'VALOR_FRETE': 'Valor Frete (R$)'}
                )
                
                if len(gastos_transp) > 7:
                    altura_grafico_transp = len(gastos_transp) * 50 + 100
                    fig_transp.update_layout(height=altura_grafico_transp, plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', margin=dict(l=0, r=0, t=40, b=0), xaxis_title="R$", yaxis_title="")
                    with st.container(height=450):
                        st.plotly_chart(fig_transp, use_container_width=True, on_select="rerun", key="chart_transp")
                else:
                    fig_transp.update_layout(height=450, plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', margin=dict(l=0, r=0, t=40, b=0), xaxis_title="R$", yaxis_title="")
                    st.plotly_chart(fig_transp, use_container_width=True, on_select="rerun", key="chart_transp")
            
            # Avisos se cross-filtering ativado
            if mes_clicados or uf_clicadas or transp_clicadas:
                st.info("💡 Você clicou em um gráfico! Os dados acima estão filtrados com base na sua seleção. Clique na mesma barra novamente ou no fundo do gráfico para limpar a seleção.")

    with tab_auditoria:
        import io
        st.subheader(f"Tabela de Detalhamento ({len(df_filtrado)} registros)")
        
        m1, m2 = st.columns(2)
        total_sidebar = df_filtrado['VALOR FRETE'].sum() if 'VALOR FRETE' in df_filtrado.columns else 0.0
        str_total_sidebar = f"R$ {total_sidebar:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        
        if filtro_duplicadas == "Mostrar APENAS as Repetidas (Valor Excedente)":
            m1.metric("🚨 Valor Excedente (Soma das Repetições na Tabela)", str_total_sidebar)
        elif filtro_duplicadas == "Mostrar APENAS Cobranças Válidas (Para Pagamento)":
            m1.metric("💰 Valor Total Seguro (Livre de Duplicidade)", str_total_sidebar)
        else:
            m1.metric("💰 Valor Total na Tabela (Misturado)", str_total_sidebar)
            if valor_duplicado_real > 0:
                str_duplicado = f"R$ {valor_duplicado_real:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                m2.metric("🚨 Risco de Duplicidade Oculto", str_duplicado)

        tem_motivo = 'MOTIVO' in df_filtrado.columns
        tem_vencimento = 'DATA DE VENCIMENTO' in df_filtrado.columns

        colunas_visiveis = ['NOTA FISCAL', 'NUMERO DA FATURA']
        if tem_vencimento: colunas_visiveis.append('DATA DE VENCIMENTO')
        colunas_visiveis.extend(['TICKET'])
        if 'VALOR FRETE' in df_filtrado.columns: colunas_visiveis.append('VALOR FRETE')
        colunas_visiveis.extend(['NOME DA TRANSPORTADORA', 'STATUS'])
        if tem_motivo: colunas_visiveis.append('MOTIVO')
        colunas_visiveis.extend(['DATA DA COLETA', 'ANO'])
        if 'MÊS' in df_filtrado.columns: colunas_visiveis.append('MÊS')
        if 'CEP' in df_filtrado.columns: colunas_visiveis.append('CEP')
        if 'UF' in df_filtrado.columns: colunas_visiveis.append('UF')
            
        duplicadas = df_filtrado.duplicated(subset=['NOTA FISCAL'], keep=False)
        
        def highlight_duplicadas(row):
            if duplicadas.loc[row.name]:
                return ['background-color: rgba(255, 75, 75, 0.3)'] * len(row)
            return [''] * len(row)

        styled_df_excel = df_filtrado[colunas_visiveis].style.apply(highlight_duplicadas, axis=1)
            
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            styled_df_excel.to_excel(writer, index=False, sheet_name='Auditoria')
            
        st.download_button(
            label="📥 Exportar Tabela para Excel",
            data=buffer.getvalue(),
            file_name="Auditoria_Filtrada.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        
        if 'VALOR FRETE' in df_filtrado.columns:
            styled_df_web = styled_df_excel.format({
                'VALOR FRETE': lambda x: f"R$ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".") if pd.notnull(x) else "R$ 0,00"
            })
        else:
            styled_df_web = styled_df_excel

        st.dataframe(styled_df_web, use_container_width=True, hide_index=True, height=600)
    
else:
    st.info("👋 **Nenhuma base de dados encontrada!** Por favor, faça o upload das planilhas originais para começar a usar o Dashboard.")
    st.markdown("### ⚙️ Configuração Inicial")
    
    col_up1, col_up2 = st.columns(2)
    with col_up1:
        st.markdown("**1. Arquivo de Coletas:**")
        upload_coletas_main = st.file_uploader("Planilha: Tracking Coletas", type=["xlsx", "xls"], key="up_col_main")
    with col_up2:
        st.markdown("**2. Arquivo de Faturas:**")
        upload_faturas_main = st.file_uploader("Planilha: Resumo Faturas", type=["xlsx", "xls"], key="up_fat_main")
        
    st.markdown("<br>", unsafe_allow_html=True)
    
    if st.button("🚀 Processar e Iniciar Dashboard", type="primary", use_container_width=True):
        if upload_coletas_main and upload_faturas_main:
            with st.spinner("Processando e cruzando as planilhas pela primeira vez. Isso pode demorar alguns segundos..."):
                try:
                    from merge_script import consolidate_data
                    df_novo = consolidate_data(faturas_file=upload_faturas_main, coletas_file=upload_coletas_main)
                    df_novo.to_excel(consolidado_path, index=False)
                    st.cache_data.clear()
                    st.success("Tudo pronto! Carregando dashboard...")
                    st.rerun()
                except Exception as e:
                    import traceback
                    st.error(f"Erro ao processar as planilhas: {e}")
                    st.text(traceback.format_exc())
        else:
            st.warning("⚠️ Faça o upload dos dois arquivos acima para continuar.")
