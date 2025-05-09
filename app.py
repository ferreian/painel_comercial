import streamlit as st
import pandas as pd
from st_aggrid import AgGrid, GridOptionsBuilder, JsCode
import io
import plotly.express as px
import plotly.graph_objects as go
import numpy as np
import statsmodels.api as sm


st.set_page_config(
    page_title="Painel Comercial",
    page_icon="📊",
    layout="wide"
)

st.title("📈 Painel Comercial")
st.markdown(
    "Explore os dados do Painel Comercial. Utilize os filtros para refinar a visualização dos dados.")

# Caminho do arquivo
caminho_arquivo = "dataset/resultados.xlsx"

try:
    with st.spinner("Carregando dados..."):
        df = pd.read_excel(caminho_arquivo, sheet_name="resultados")

        # >>> Colunas calculadas <<<
        if "numeroLinhas" in df.columns and "comprimentoLinha" in df.columns:
            df["areaParcela"] = df["numeroLinhas"] * \
                df["comprimentoLinha"] * 0.5

        colunas_plantas = ["numeroPlantas10Metros1a", "numeroPlantas10Metros2a",
                           "numeroPlantas10Metros3a", "numeroPlantas10Metros4a"]
        if all(col in df.columns for col in colunas_plantas):
            df["numeroPlantasMedio10m"] = df[colunas_plantas].replace(
                0, pd.NA).mean(axis=1, skipna=True)
            df["Pop_Final"] = (20000 * df["numeroPlantasMedio10m"]) / 10

        if "numeroPlantasMedio10m" in df.columns:
            df["popMediaFinal"] = (10000 / 0.5) * \
                (df["numeroPlantasMedio10m"] / 10)

        if all(col in df.columns for col in ["pesoParcela", "umidadeParcela", "areaParcela"]):
            df["producaoCorrigida"] = ((df["pesoParcela"] * (100 - df["umidadeParcela"]) / 87) * (
                10000 / df["areaParcela"])).astype(float).round(1)

        if "producaoCorrigida" in df.columns:
            df["producaoCorrigidaSc"] = (
                df["producaoCorrigida"] / 60).astype(float).round(1)

        if all(col in df.columns for col in ["pesoMilGraos", "umidadeAmostraPesoMilGraos"]):
            df["PMG_corrigido"] = (
                df["pesoMilGraos"] * ((100 - df["umidadeAmostraPesoMilGraos"]) / 87)).astype(float).round(1)

        if all(col in df.columns for col in ["fazendaRef", "indexTratamento"]):
            df["ChaveFaixa"] = df["fazendaRef"].astype(
                str) + "_" + df["indexTratamento"].astype(str)

        for col in ["dataPlantio", "dataColheita"]:
            if col in df.columns:
                df[col] = pd.to_datetime(
                    df[col], origin="unix", unit="s").dt.strftime("%d/%m/%Y")

        # >>> Renomear colunas <<<
        mapeamento_colunas = {
            "nome": "Cultivar",
            "gm": "GM",
            "umidadeParcela": "U %",
            "fazendaRef": "FazendaRef",
            "nomeFazenda": "Fazenda",
            "nomeProdutor": "Produtor",
            "latitude": "Latitude",
            "longitude": "Longitude",
            "altitude": "Altitude",
            "microrregiao": "Microrregião",
            "dataPlantio": "Plantio",
            "dataColheita": "Colheita",
            "nomeCidade": "Cidade",
            "nomeEstado": "Estado",
            "codigoEstado": "Código Estado",
            "macro": "Macro",
            "rec": "REC",
            "Pop_Final": "População Final",
            "producaoCorrigida": "Prod_kg_@13%",
            "producaoCorrigidaSc": "Prod_sc_@13%"
        }
        df.rename(columns=mapeamento_colunas, inplace=True)

    # >>> Layout de filtros e resultados <<<
    col_filtros, col_resultados = st.columns([0.15, 0.85])

    with col_filtros:
        st.header("🎧 Filtros")
        filtros = {"Macro": "macro_", "REC": "rec_",
                   "Microrregião": "micro_", "Estado": "estado_", "Cidade": "cidade_"}
        for coluna, chave in filtros.items():
            if coluna in df.columns:
                with st.expander(coluna):
                    opcoes = sorted(df[coluna].dropna().unique())
                    filtro = []
                    for opcao in opcoes:
                        if st.checkbox(str(opcao), key=f"{chave}{opcao}"):
                            filtro.append(opcao)
                    if filtro:
                        df = df[df[coluna].isin(filtro)]

    with col_resultados:
        st.header(
            "📈 Conjunto de dados dos ensaios de desenvolvimento de produto - 2024/2025")
        st.success("Arquivo carregado com sucesso!")

        # Exibe tabela principal
        colunas_visiveis = ["FazendaRef", "Fazenda", "Cultivar", "GM", "U %", "Plantio", "Colheita",
                            "População Final", "Prod_kg_@13%", "Prod_sc_@13%", "Altitude",
                            "Microrregião", "Cidade", "Estado", "Código Estado", "Macro", "REC"]
        df = df[[col for col in colunas_visiveis if col in df.columns]]

        gb = GridOptionsBuilder.from_dataframe(df)
        for col in df.select_dtypes(include=["float"]).columns:
            gb.configure_column(
                field=col, type=["numericColumn"], valueFormatter="x.toFixed(1)")
        gb.configure_default_column(
            cellStyle={'color': 'black', 'fontSize': '14px'})
        gb.configure_grid_options(headerHeight=30)
        custom_css = {".ag-header-cell-label": {"font-weight": "bold",
                                                "font-size": "15px", "color": "black"}}

        AgGrid(df, gridOptions=gb.build(), height=600, custom_css=custom_css,
               theme='streamlit', fit_columns_on_grid_load=True)

        # Exportação
        csv = df.to_csv(index=False).encode('utf-8')
        st.download_button("📥 Baixar tabela (CSV)", data=csv,
                           file_name="resultado_painel_comercial.csv", mime='text/csv')

        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='Resultados')
        st.download_button("📥 Baixar tabela (Excel)", data=buffer.getvalue(),
                           file_name="resultado_painel_comercial.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        # 🔥 Análise Head-to-Head (AGORA DENTRO DA MESMA COLUNA!)
        st.markdown("---")
        st.markdown("### ⚔️ Análise Head-to-Head entre Cultivares")
        st.markdown("""
        <small>A Análise Head-to-Head compara dois cultivares em cada FazendaRef, mostrando vitórias, derrotas e empates (diferenças entre -1 e 1 sc/ha).</small>
        """, unsafe_allow_html=True)

        if st.button("🔁 Rodar Análise Head-to-Head"):
            df["Local"] = df["FazendaRef"]
            df_h2h = df[["Local", "Fazenda", "Cidade", "Cultivar",
                         "Prod_sc_@13%", "População Final", "U %"]].dropna()
            df_h2h = df_h2h[df_h2h["Prod_sc_@13%"] > 0]

            resultados_h2h = []
            for local, grupo in df_h2h.groupby("Local"):
                fazenda_val = grupo["Fazenda"].iloc[0]
                cidade_val = grupo["Cidade"].iloc[0]
                cultivares = grupo["Cultivar"].unique()

                for head in cultivares:
                    head_row = grupo[grupo["Cultivar"] == head]
                    if head_row.empty:
                        continue
                    prod_head = head_row["Prod_sc_@13%"].values[0]
                    pop_head = head_row["População Final"].values[0]
                    umid_head = head_row["U %"].values[0]

                    for check in cultivares:
                        if head == check:
                            continue
                        check_row = grupo[grupo["Cultivar"] == check]
                        if check_row.empty:
                            continue
                        prod_check = check_row["Prod_sc_@13%"].values[0]
                        pop_check = check_row["População Final"].values[0]
                        umid_check = check_row["U %"].values[0]
                        diff = prod_head - prod_check
                        win = int(diff > 1)
                        draw = int(-1 <= diff <= 1)

                        resultados_h2h.append({
                            "Local": local, "Fazenda": fazenda_val, "Cidade": cidade_val,
                            "Head": head, "Check": check,
                            "Head_Mean": round(prod_head, 1), "Check_Mean": round(prod_check, 1),
                            "População Final Head": round(pop_head, 0), "U % Head": round(umid_head, 1),
                            "População Final Check": round(pop_check, 0), "U % Check": round(umid_check, 1),
                            "Difference": round(diff, 1), "Number_of_Win": win, "Is_Draw": draw,
                            "Percentage_of_Win": 100.0 if win else 0.0, "Number_of_Comparison": 1
                        })

            df_resultado_h2h = pd.DataFrame(resultados_h2h)
            st.session_state["df_resultado_h2h"] = df_resultado_h2h
            st.success("✅ Análise concluída!")

        if "df_resultado_h2h" in st.session_state:
            df_resultado_h2h = st.session_state["df_resultado_h2h"]
            colunas_visiveis = df_resultado_h2h.columns.tolist()

            st.markdown("### 🎯 Selecione os cultivares para exibir na Tabela")
            col1, col2 = st.columns(2)
            head_filtrado = col1.selectbox(
                "Cultivar Head", sorted(df_resultado_h2h["Head"].unique()))
            check_filtrado = col2.selectbox(
                "Cultivar Check", sorted(df_resultado_h2h["Check"].unique()))

            df_filtrado = df_resultado_h2h[
                (df_resultado_h2h["Head"] == head_filtrado) & (
                    df_resultado_h2h["Check"] == check_filtrado)
            ]

            st.markdown(
                f"### 📋 Tabela Head to Head: <b>{head_filtrado} x {check_filtrado}</b>",
                unsafe_allow_html=True)

            if not df_filtrado.empty:
                gb = GridOptionsBuilder.from_dataframe(df_filtrado)
                cell_style_js = JsCode("""
                    function(params) {
                        let value = params.value;
                        let min = 0;
                        let max = 100;
                        let ratio = (value - min) / (max - min);
                        let r, g, b;
                        if (ratio < 0.5) { r = 253; g = 98 + ratio*2*(200-98); b = 94 + ratio*2*(15-94); }
                        else { r = 242-(ratio-0.5)*2*(242-1); g = 200-(ratio-0.5)*2*(200-184); b = 15+(ratio-0.5)*2*(170-15); }
                        return {'backgroundColor':'rgb('+r+','+g+','+b+')','color':'black','fontWeight':'bold','fontSize':'16px'};
                    }
                """)
                for col in df_filtrado.select_dtypes(include=["float"]).columns:
                    gb.configure_column(
                        col, type=["numericColumn"], valueFormatter="x.toFixed(1)")
                gb.configure_column("Head_Mean", cellStyle=cell_style_js)
                gb.configure_column("Check_Mean", cellStyle=cell_style_js)
                gb.configure_default_column(
                    cellStyle={'color': 'black', 'fontSize': '14px'})
                gb.configure_grid_options(headerHeight=30)
                custom_css = {".ag-header-cell-label": {"font-weight": "bold",
                                                        "font-size": "15px", "color": "black"}}

                AgGrid(df_filtrado, gridOptions=gb.build(), height=800, custom_css=custom_css,
                       theme='streamlit', fit_columns_on_grid_load=True, allow_unsafe_jscode=True)

                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
                    df_resultado_h2h.to_excel(
                        writer, sheet_name="head_to_head", index=False)
                st.download_button("📥 Baixar Resultado Head-to-Head (Completo)", buffer.getvalue(),
                                   file_name="resultado_head_to_head.xlsx",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
                    df_filtrado.to_excel(
                        writer, sheet_name="head_to_head_filtrado", index=False)
                st.download_button(f"📥 Baixar Resultado Filtrado ({head_filtrado} vs {check_filtrado})",
                                   buffer.getvalue(),
                                   file_name=f"head_to_head_{head_filtrado}_vs_{check_filtrado}.xlsx",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            else:
                st.warning("⚠️ Nenhum dado disponível para essa combinação.")

        # ⬇️ NOVA ANÁLISE: CARDS DE RESUMO HEAD-TO-HEAD
        # ✅ VERIFICA SE JÁ FOI GERADA A ANÁLISE
        if "df_resultado_h2h" in st.session_state:
            df_resultado_h2h = st.session_state["df_resultado_h2h"]

            st.markdown(
                "### 🔹 Selecione os cultivares para comparação Head to Head")
            cultivares_unicos = sorted(df_resultado_h2h["Head"].unique())
            colA, colB, colC = st.columns([0.3, 0.4, 0.3])

            with colA:
                head_select = st.selectbox(
                    "Selecionar Cultivar Head", options=cultivares_unicos, key="head_select")
            with colB:
                st.markdown("<h1 style='text-align: center;'>X</h1>",
                            unsafe_allow_html=True)
            with colC:
                check_select = st.selectbox(
                    "Selecionar Cultivar Check", options=cultivares_unicos, key="check_select")

            if head_select and check_select and head_select != check_select:
                df_selecionado = df_resultado_h2h[
                    (df_resultado_h2h["Head"] == head_select) & (
                        df_resultado_h2h["Check"] == check_select)
                ]

                num_locais = df_selecionado["Local"].nunique()
                vitorias = df_selecionado[df_selecionado["Difference"] > 1].shape[0]
                derrotas = df_selecionado[df_selecionado["Difference"]
                                          < -1].shape[0]
                empates = df_selecionado[df_selecionado["Difference"].between(
                    -1, 1)].shape[0]
                max_diff = df_selecionado["Difference"].max(
                ) if not df_selecionado.empty else 0
                min_diff = df_selecionado["Difference"].min(
                ) if not df_selecionado.empty else 0
                media_diff_vitorias = df_selecionado[df_selecionado["Difference"] > 1]["Difference"].mean(
                ) or 0
                media_diff_derrotas = df_selecionado[df_selecionado["Difference"]
                                                     < -1]["Difference"].mean() or 0

                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    st.markdown(f"""
                        <div style="background-color:#f2f2f2; padding:15px; border-radius:10px; text-align:center;">
                            <h5 style="font-weight:bold; color:#333;">📍 Número de Locais</h5>
                            <div style="font-size:14px; color:#f2f2f2;">&nbsp;</div>
                            <h2 style="margin: 10px 0; color:#333; font-weight:bold; font-size: 4em;">{num_locais}</h2>
                            <div style="font-size:14px; color:#f2f2f2;">&nbsp;</div>
                        </div>
                    """, unsafe_allow_html=True)
                with col2:
                    st.markdown(f"""
                        <div style="background-color:#01B8AA80; padding:15px; border-radius:10px; text-align:center;">
                            <h5 style="font-weight:bold; color:#004d47;">✅ Vitórias</h5>
                            <div style="font-size:14px; color:#004d47;">Max: {max_diff:.1f} sc/ha</div>
                            <h2 style="margin: 10px 0; color:#004d47; font-weight:bold; font-size: 4em;">{vitorias}</h2>
                            <div style="font-size:14px; color:#004d47;">Média: {media_diff_vitorias:.1f} sc/ha</div>
                        </div>
                    """, unsafe_allow_html=True)
                with col3:
                    st.markdown(f"""
                        <div style="background-color:#F2C80F80; padding:15px; border-radius:10px; text-align:center;">
                            <h5 style="font-weight:bold; color:#8a7600;">➖ Empates</h5>
                            <div style="font-size:14px; color:#8a7600;">Entre -1 e 1 sc/ha</div>
                            <h2 style="margin: 10px 0; color:#8a7600; font-weight:bold; font-size: 4em;">{empates}</h2>
                            <div style="font-size:14px; color:#F2C80F80;">&nbsp;</div>
                        </div>
                    """, unsafe_allow_html=True)
                with col4:
                    st.markdown(f"""
                        <div style="background-color:#FD625E80; padding:15px; border-radius:10px; text-align:center;">
                            <h5 style="font-weight:bold; color:#7c1f1c;">❌ Derrotas</h5>
                            <div style="font-size:14px; color:#7c1f1c;">Min: {min_diff:.1f} sc/ha</div>
                            <h2 style="margin: 10px 0; color:#7c1f1c; font-weight:bold; font-size: 4em;">{derrotas}</h2>
                            <div style="font-size:14px; color:#7c1f1c;">Média: {media_diff_derrotas:.1f} sc/ha</div>
                        </div>
                    """, unsafe_allow_html=True)

            # >>>>> Comparação Multicheck <<<<<
            st.markdown("---")
            st.markdown("### 🔹 Comparação Head x Múltiplos Checks")
            st.markdown("""
            <small>
            Essa análise permite comparar um cultivar (Head) com vários outros (Checks) ao mesmo tempo. 
            Ela apresenta o percentual de vitórias, produtividade média e a diferença média de performance 
            em relação aos demais cultivares selecionados.
            </small>
            """, unsafe_allow_html=True)

            head_unico = st.selectbox(
                "Cultivar Head", options=cultivares_unicos, key="multi_head")
            opcoes_checks = [c for c in cultivares_unicos if c != head_unico]
            checks_selecionados = st.multiselect(
                "Cultivares Check", options=opcoes_checks, key="multi_checks")

            if head_unico and checks_selecionados:
                df_multi = df_resultado_h2h[
                    (df_resultado_h2h["Head"] == head_unico) & (
                        df_resultado_h2h["Check"].isin(checks_selecionados))
                ]
                if not df_multi.empty:
                    prod_head_media = df_multi["Head_Mean"].mean().round(1)
                    st.markdown(
                        f"#### 🎯 Cultivar Head: **{head_unico}** | Produtividade Média: **{prod_head_media} sc/ha**")
                    resumo = df_multi.groupby("Check").agg({
                        "Number_of_Win": "sum",
                        "Number_of_Comparison": "sum",
                        "Check_Mean": "mean"
                    }).reset_index()
                    resumo.rename(columns={
                        "Check": "Cultivar Check",
                        "Number_of_Win": "Vitórias",
                        "Number_of_Comparison": "Num_Locais",
                        "Check_Mean": "Prod_sc_ha_media"
                    }, inplace=True)
                    resumo["% Vitórias"] = (
                        resumo["Vitórias"] / resumo["Num_Locais"] * 100).round(1)
                    resumo["Prod_sc_ha_media"] = resumo["Prod_sc_ha_media"].round(
                        1)
                    resumo["Diferença Média"] = (
                        prod_head_media - resumo["Prod_sc_ha_media"]).round(1)
                    resumo = resumo[["Cultivar Check", "% Vitórias",
                                     "Num_Locais", "Prod_sc_ha_media", "Diferença Média"]]

                    col_tabela, col_grafico = st.columns([1.4, 1.6])
                    with col_tabela:
                        st.markdown("### 📊 Tabela Comparativa")
                        gb = GridOptionsBuilder.from_dataframe(resumo)
                        gb.configure_default_column(
                            cellStyle={'fontSize': '14px'})
                        gb.configure_grid_options(headerHeight=30)
                        custom_css = {
                            ".ag-header-cell-label": {"font-weight": "bold", "font-size": "15px", "color": "black"}}
                        AgGrid(resumo, gridOptions=gb.build(),
                               height=400, custom_css=custom_css)

                        buffer = io.BytesIO()
                        with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
                            resumo.to_excel(
                                writer, sheet_name="comparacao_multi_check", index=False)
                        st.download_button(label="📅 Baixar Comparação (Excel)", data=buffer.getvalue(),
                                           file_name=f"comparacao_{head_unico}_vs_checks.xlsx",
                                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                    with col_grafico:
                        fig_diff = go.Figure()
                        cores_personalizadas = resumo["Diferença Média"].apply(
                            lambda x: "#01B8AA" if x > 1 else "#FD625E" if x < -1 else "#F2C80F"
                        )
                        fig_diff.add_trace(go.Bar(
                            y=resumo["Cultivar Check"],
                            x=resumo["Diferença Média"],
                            orientation='h',
                            text=resumo["Diferença Média"].round(1),
                            textposition="outside",
                            textfont=dict(
                                size=16, family="Arial Black", color="black"),
                            marker_color=cores_personalizadas
                        ))
                        fig_diff.update_layout(
                            title=dict(text="📊 Diferença Média de Produtividade", font=dict(
                                size=20, family="Arial Black")),
                            xaxis=dict(title=dict(
                                text="Diferença Média (sc/ha)", font=dict(size=16)), tickfont=dict(size=14)),
                            yaxis=dict(title=dict(text="Check"),
                                       tickfont=dict(size=14)),
                            margin=dict(t=30, b=40, l=60, r=30),
                            height=400,
                            showlegend=False
                        )
                        st.plotly_chart(fig_diff, use_container_width=True)
                else:
                    st.info(
                        "❓ Nenhuma comparação disponível com os Checks selecionados.")
        else:
            st.warning(
                "⚠️ Você precisa rodar a análise Head-to-Head primeiro clicando no botão 🔁 acima.")

        # =======================================================
        # >>>>> Análise de Índice Ambiental <<<<<
        # =======================================================

        st.markdown("---")
        with st.expander("📉 Índice Ambiental: Média do Local x Produção do Material", expanded=False):
            df_dispersao = df.copy()  # usa a tabela principal já filtrada

            # ➡️ define o nome da coluna de produtividade a ser usada
            coluna_prod = "Prod_sc_@13%"

            # verifica se a coluna existe
            if coluna_prod in df_dispersao.columns:
                # calcula média do local
                df_media_local = df_dispersao.groupby("FazendaRef")[coluna_prod].mean(
                ).reset_index().rename(columns={coluna_prod: "Media_Local"})
                df_dispersao = df_dispersao.merge(
                    df_media_local, on="FazendaRef", how="left")

                # remove registros inválidos
                df_dispersao = df_dispersao.replace([np.inf, -np.inf], np.nan)
                df_dispersao = df_dispersao.dropna(
                    subset=["Media_Local", coluna_prod])

                cultivares_disp = sorted(
                    df_dispersao["Cultivar"].dropna().unique())
                cultivar_default = "78KA42"

                if cultivar_default in cultivares_disp:
                    valor_default = [cultivar_default]
                elif cultivares_disp:
                    valor_default = [cultivares_disp[0]]
                else:
                    valor_default = []

                cultivares_selecionadas = st.multiselect(
                    "🧬 Selecione as Cultivares:", cultivares_disp, default=valor_default)
                mostrar_outras = st.checkbox(
                    "👀 Mostrar outras cultivares", value=True)

                if mostrar_outras:
                    df_dispersao["Cor"] = df_dispersao["Cultivar"].apply(
                        lambda x: x if x in cultivares_selecionadas else "Outras")
                else:
                    df_dispersao = df_dispersao[df_dispersao["Cultivar"].isin(
                        cultivares_selecionadas)]
                    df_dispersao["Cor"] = df_dispersao["Cultivar"]

                color_map = {cult: px.colors.qualitative.Plotly[i % 10] for i, cult in enumerate(
                    cultivares_selecionadas)}
                if mostrar_outras:
                    color_map["Outras"] = "#d3d3d3"

                fig_disp = px.scatter(
                    df_dispersao,
                    x="Media_Local",
                    y=coluna_prod,
                    color="Cor",
                    color_discrete_map=color_map,
                    labels={
                        "Media_Local": "Média do Local",
                        coluna_prod: "Produção do Material",
                        "Cor": "Cultivar"
                    }
                )

                # ➕ adiciona linha de tendência por cultivar
                for cultivar in cultivares_selecionadas:
                    df_cult = df_dispersao[df_dispersao["Cultivar"] == cultivar].copy(
                    )

                    # 🧹 limpa NaN e Inf
                    df_cult = df_cult.replace([np.inf, -np.inf], np.nan)
                    df_cult = df_cult.dropna(
                        subset=["Media_Local", coluna_prod])

                    if not df_cult.empty and df_cult.shape[0] > 1:
                        try:
                            X_train = df_cult[["Media_Local"]].astype(float)
                            X_train = sm.add_constant(X_train)
                            y_train = df_cult[coluna_prod].astype(float)

                            model = sm.OLS(y_train, X_train).fit()

                            x_vals = np.linspace(df_dispersao["Media_Local"].min(
                            ), df_dispersao["Media_Local"].max(), 100)
                            X_pred = pd.DataFrame({"Media_Local": x_vals})
                            X_pred = sm.add_constant(X_pred)
                            y_pred = model.predict(X_pred)

                            fig_disp.add_trace(go.Scatter(
                                x=x_vals,
                                y=y_pred,
                                mode="lines",
                                name=f"Tendência - {cultivar}",
                                line=dict(color=color_map.get(
                                    cultivar, "black"), dash="solid")
                            ))

                        except Exception as e:
                            st.warning(
                                f"⚠️ Erro ao calcular tendência para {cultivar}: {e}")
                    else:
                        st.info(
                            f"⚠️ Cultivar '{cultivar}' não tem dados suficientes para regressão (mínimo 2 linhas válidas).")

                # estilos
                font_bold = dict(size=20, family="Arial Bold", color="black")

                fig_disp.update_layout(
                    plot_bgcolor="white",
                    title=dict(
                        text="Índice Ambiental: Cultivares Selecionadas", font=font_bold),
                    xaxis=dict(title=dict(text="Média do Local", font=font_bold),
                               tickfont=font_bold, showgrid=True, gridcolor="lightgray"),
                    yaxis=dict(title=dict(text="Produção do Material", font=font_bold),
                               tickfont=font_bold, showgrid=True, gridcolor="lightgray"),
                    legend=dict(orientation="h", yanchor="bottom",
                                y=1.02, xanchor="right", x=1, font=font_bold)
                )

                st.plotly_chart(fig_disp, use_container_width=True)

            else:
                st.info(
                    f"⚠️ Coluna '{coluna_prod}' não encontrada na tabela de dados.")


except FileNotFoundError:
    st.error(f"Arquivo não encontrado: {caminho_arquivo}")
except ValueError as e:
    st.error(f"Aba 'resultados' não encontrada no arquivo Excel: {e}")
except Exception as e:
    st.error(f"Erro ao carregar o arquivo: {e}")
