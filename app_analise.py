from datetime import datetime, timedelta
import os
import pandas as pd
import streamlit as st

# Configuração da Página
st.set_page_config(
    page_title="Análise de Adesão Logística | Grupo RMC Mariano",
    page_icon="📊",
    layout="wide",
)

st.markdown(
    """
    <div style="background: linear-gradient(135deg, #ffffff 0%, #f1f5f9 100%); padding: 20px; border-radius: 12px; border: 1px solid #e2e8f0; border-left: 8px solid #5a8c71; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05); margin-bottom: 20px; text-align: center;">
        <h2 style="color: #1e293b; margin: 0; font-size: 24px; font-weight: 700;">GRUPO RMC MARIANO - PAINEL DE ANÁLISE DE ADESÃO</h2>
        <p style="color: #64748b; margin: 5px 0 0 0; font-size: 14px;">Monitoramento de Pedidos, Datas, Semanas e Categorias de Utilização</p>
    </div>
""",
    unsafe_allow_html=True,
)

# Caminho do ficheiro Excel padrão na pasta local
EXCEL_PADRAO = "20261002_GestaoPedidos_Visão_detalhada_da_utilização_por_pedido_bd72136d03b4.xlsx"

# Painel Lateral para Upload, Botão de Atualizar e Filtros
st.sidebar.markdown("### 📁 Gestão de Ficheiro")

if st.sidebar.button("🔄 Atualizar Página / Dados", use_container_width=True):
    st.cache_data.clear()
    st.rerun()

arquivo_carregado = st.sidebar.file_uploader(
    "Arraste ou envie um novo Excel aqui:", type=["xlsx", "xls"]
)

# Função para carregar os dados com a conversão correta de data BR
@st.cache_data
def carregar_dados(uploaded_file):
    if uploaded_file is not None:
        df = pd.read_excel(uploaded_file, sheet_name="ADESÃO A PLATAFORMA LOGÍSTICA")
    else:
        if os.path.exists(EXCEL_PADRAO):
            df = pd.read_excel(EXCEL_PADRAO, sheet_name="ADESÃO A PLATAFORMA LOGÍSTICA")
        else:
            return None

    # Converte colunas de data forçando o formato brasileiro (Dia/Mês/Ano)
    df["DATA DE APROVAÇÃO DO PEDIDO"] = pd.to_datetime(
        df["DATA DE APROVAÇÃO DO PEDIDO"], format="%d/%m/%Y", errors="coerce"
    )
    df["DATA DE FINALIZAÇÃO DO PEDIDO"] = pd.to_datetime(
        df["DATA DE FINALIZAÇÃO DO PEDIDO"], format="%d/%m/%Y", errors="coerce"
    )
    return df

df_original = carregar_dados(arquivo_carregado)

if df_original is None:
    st.warning(
        "⚠️ Nenhum ficheiro carregado. Por favor, envie um ficheiro Excel na barra lateral esquerda ou coloque o ficheiro padrão na pasta."
    )
else:
    if arquivo_carregado is not None:
        st.sidebar.success(f"✅ Ficheiro ativo: {arquivo_carregado.name}")
    else:
        st.sidebar.info("ℹ A usar o ficheiro padrão da pasta.")

    st.sidebar.markdown("---")
    st.sidebar.markdown("### ⚙️ Filtros do Painel")

    # Tratamento correto das datas mínimas e máximas do painel
    min_data_val = df_original["DATA DE APROVAÇÃO DO PEDIDO"].min()
    max_data_val = df_original["DATA DE APROVAÇÃO DO PEDIDO"].max()
    
    min_data = min_data_val.date() if pd.notnull(min_data_val) else datetime.today().date()
    max_data = max_data_val.date() if pd.notnull(max_data_val) else datetime.today().date()

    st.sidebar.markdown("#### 📅 Período de Aprovação")
    data_inicio = st.sidebar.date_input("Data Inicial:", value=min_data, min_value=min_data, max_value=max_data)
    data_fim = st.sidebar.date_input("Data Final:", value=max_data, min_value=min_data, max_value=max_data)

    # Filtro de Categoria de Adesão
    categorias_disponiveis = sorted(
        df_original["CATEGORIA DE ADESÃO"].dropna().unique().tolist()
    )
    cat_selecionadas = st.sidebar.multiselect(
        "Filtrar Categorias de Adesão:",
        options=categorias_disponiveis,
        default=categorias_disponiveis,
    )

    # Aplicação dos Filtros
    mask = (
        (df_original["DATA DE APROVAÇÃO DO PEDIDO"].dt.date >= data_inicio)
        & (df_original["DATA DE APROVAÇÃO DO PEDIDO"].dt.date <= data_fim)
        & (df_original["CATEGORIA DE ADESÃO"].isin(cat_selecionadas))
    )

    df_filtrado = df_original.loc[mask]

    # Métricas Principais (KPIs) consolidadas do mês/período selecionado
    total_pedidos = len(df_filtrado)
    
    usa_bem_qtd = len(df_filtrado[df_filtrado["CATEGORIA DE ADESÃO"] == "Usa Bem"])
    uso_mod_qtd = len(df_filtrado[df_filtrado["CATEGORIA DE ADESÃO"] == "Uso Moderado"])
    nao_usa_qtd = len(df_filtrado[df_filtrado["CATEGORIA DE ADESÃO"] == "Não Usa"])
    em_aberto_qtd = len(df_filtrado[df_filtrado["CATEGORIA DE ADESÃO"] == "Em Aberto"])

    perc_usa_bem = (usa_bem_qtd / total_pedidos * 100) if total_pedidos > 0 else 0
    perc_uso_mod = (uso_mod_qtd / total_pedidos * 100) if total_pedidos > 0 else 0
    perc_nao_usa = (nao_usa_qtd / total_pedidos * 100) if total_pedidos > 0 else 0
    perc_em_aberto = (em_aberto_qtd / total_pedidos * 100) if total_pedidos > 0 else 0

    st.markdown("### 📈 Indicadores Principais (Consolidado do Mês/Período)")
    col1, col2, col3, col4, col5 = st.columns(5)

    col1.metric("📦 Total de Pedidos", f"{total_pedidos:,}")
    col2.metric("🟢 Usa Bem", f"{usa_bem_qtd:,}", f"{perc_usa_bem:.2f}%")
    col3.metric("🟡 Uso Moderado", f"{uso_mod_qtd:,}", f"{perc_uso_mod:.2f}%")
    col4.metric("🔴 Não Usa", f"{nao_usa_qtd:,}", f"{perc_nao_usa:.2f}%")
    col5.metric("⏳ Em Aberto", f"{em_aberto_qtd:,}", f"{perc_em_aberto:.2f}%")

    st.markdown("---")

    # ==========================================
    # SEÇÃO: DIVISÃO EM 4 SEMANAS PROGRAMADAS
    # ==========================================
    st.markdown("### 📅 Análise Comparativa por Semanas do Mês")
    st.markdown("Divisão automática do período selecionado em 4 blocos semanais para acompanhamento da evolução percentual.")

    if not df_filtrado.empty:
        dt_inicio_base = pd.to_datetime(data_inicio)
        dt_fim_base = pd.to_datetime(data_fim)
        
        total_dias = (dt_fim_base - dt_inicio_base).days + 1
        tamanho_semana = total_dias / 4.0

        dados_semanas = []
        for i in range(4):
            inicio_sem = dt_inicio_base + timedelta(days=int(i * tamanho_semana))
            if i == 3:
                fim_sem = dt_fim_base 
            else:
                fim_sem = dt_inicio_base + timedelta(days=int((i + 1) * tamanho_semana) - 1)

            df_sem = df_filtrado[
                (df_filtrado["DATA DE APROVAÇÃO DO PEDIDO"] >= inicio_sem) & 
                (df_filtrado["DATA DE APROVAÇÃO DO PEDIDO"] <= fim_sem)
            ]

            tot_sem = len(df_sem)
            usa_sem = len(df_sem[df_sem["CATEGORIA DE ADESÃO"] == "Usa Bem"])
            p_usa_sem = (usa_sem / tot_sem * 100) if tot_sem > 0 else 0

            mod_sem = len(df_sem[df_sem["CATEGORIA DE ADESÃO"] == "Uso Moderado"])
            p_mod_sem = (mod_sem / tot_sem * 100) if tot_sem > 0 else 0

            nao_sem = len(df_sem[df_sem["CATEGORIA DE ADESÃO"] == "Não Usa"])
            p_nao_sem = (nao_sem / tot_sem * 100) if tot_sem > 0 else 0

            aberto_sem = len(df_sem[df_sem["CATEGORIA DE ADESÃO"] == "Em Aberto"])
            p_aberto_sem = (aberto_sem / tot_sem * 100) if tot_sem > 0 else 0

            dados_semanas.append({
                "Semana": f"Semana {i+1} ({inicio_sem.strftime('%d/%m')} a {fim_sem.strftime('%d/%m')})",
                "Total Pedidos": tot_sem,
                "Usa Bem (Qtd)": usa_sem,
                "% Usa Bem": f"{p_usa_sem:.2f}%",
                "Uso Moderado (Qtd)": mod_sem,
                "% Uso Moderado": f"{p_mod_sem:.2f}%",
                "Não Usa (Qtd)": nao_sem,
                "% Não Usa": f"{p_nao_sem:.2f}%",
                "Em Aberto (Qtd)": aberto_sem,
                "% Em Aberto": f"{p_aberto_sem:.2f}%"
            })

        df_tabela_semanas = pd.DataFrame(dados_semanas)
        st.dataframe(df_tabela_semanas, use_container_width=True)
    else:
        st.info("Nenhum dado disponível para gerar as semanas com os filtros atuais.")

    st.markdown("---")

    # Visualização Gráfica e Tabela Resumo Consolidada
    col_graf, col_tab = st.columns([1, 1])

    with col_graf:
        st.markdown("#### 📊 Distribuição Consolidada por Categoria")
        if not df_filtrado.empty:
            resumo_cat = (
                df_filtrado["CATEGORIA DE ADESÃO"]
                .value_counts()
                .reset_index(name="Quantidade")
            )
            resumo_cat.columns = ["Categoria", "Quantidade"]
            st.bar_chart(resumo_cat.set_index("Categoria"))
        else:
            st.info("Nenhum dado encontrado para o período selecionado.")

    with col_tab:
        st.markdown("#### 📋 Resumo Percentual Consolidado")
        if not df_filtrado.empty:
            tabela_resumo = (
                df_filtrado["CATEGORIA DE ADESÃO"]
                .value_counts()
                .reset_index(name="Total de Pedidos")
            )
            tabela_resumo.columns = ["Categoria de Adesão", "Total de Pedidos"]
            tabela_resumo["Percentual (%)"] = (
                (tabela_resumo["Total de Pedidos"] / total_pedidos) * 100
            ).round(2)
            st.dataframe(tabela_resumo, use_container_width=True)

    st.markdown("---")
    st.markdown("### 🔍 Detalhamento Completo dos Pedidos Filtrados")
    st.dataframe(df_filtrado, use_container_width=True)