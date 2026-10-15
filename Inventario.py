import streamlit as st
import plotly.graph_objects as go
import pandas as pd
import io
import json
import os
from datetime import datetime
from openpyxl import Workbook
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.styles import Border, Side, Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter

# Diretórios físicos de persistência local de dados
ARQUIVO_BANCO = "banco_dados.json"
ARQUIVO_ESTOQUE = "estoque_bobinas.json"
ARQUIVO_EXCEL_AUTO = "Inventario_VillageMall.xlsx"


def carregar_dados():
    """Carrega todo o histórico gravado em disco"""
    if os.path.exists(ARQUIVO_BANCO):
        try:
            with open(ARQUIVO_BANCO, "r", encoding="utf-8") as f:
                return json.load(f)
        except:
            return []
    return []


def salvar_dados(dados):
    """Grava de forma segura a base em formato JSON"""
    with open(ARQUIVO_BANCO, "w", encoding="utf-8") as f:
        json.dump(dados, f, ensure_ascii=False, indent=4)


def carregar_estoque_por_tamanho():
    """Carrega o saldo de bobinas por diâmetro técnico real"""
    if os.path.exists(ARQUIVO_ESTOQUE):
        try:
            with open(ARQUIVO_ESTOQUE, "r", encoding="utf-8") as f:
                dados = json.load(f)
                if isinstance(dados, dict) and "estoque" not in dados:
                    return dados
        except:
            return {}
    return {}


def salvar_estoque_por_tamanho(dicionario_estoque):
    """Salva fisicamente os estoques individuais de bobinas"""
    with open(ARQUIVO_ESTOQUE, "w", encoding="utf-8") as f:
        json.dump(dicionario_estoque, f, ensure_ascii=False, indent=4)


# Configuração inicial do servidor web da interface do usuário
st.set_page_config(
    page_title="Dashboard Operacional",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilização profissional em CSS para Dark Mode e Orelhas Coloridas laterais
st.markdown("""
    <style>
        .stApp {
            background-color: #1a1a1a;
        }
        .bottom-sidebar-container {
            position: fixed;
            bottom: 20px;
            left: 15px;
            width: 16rem;
            z-index: 999;
        }
        section[data-testid="stSidebar"] button {
            border-radius: 4px 12px 12px 4px !important;
            transition: all 0.3s ease !important;
        }
        section[data-testid="stSidebar"] button:hover {
            background-color: #2d2d2d !important;
            color: #ffffff !important;
        }
        div.element-container:nth-of-type(3) button { border-left: 6px solid #00ffcc !important; }
        div.element-container:nth-of-type(4) button { border-left: 6px solid #3388ff !important; }
        div.element-container:nth-of-type(5) button { border-left: 6px solid #9933ff !important; }
        div.element-container:nth-of-type(6) button { border-left: 6px solid #ff3399 !important; }
        div.element-container:nth-of-type(7) button { border-left: 6px solid #ff9933 !important; }
        div.element-container:nth-of-type(8) button { border-left: 6px solid #ffff33 !important; }
        div.element-container:nth-of-type(9) button { border-left: 6px solid #ff3333 !important; }
        section[data-testid="stSidebar"] [data-testid="stDownloadButton"] button {
            border-left: 6px solid #23a55a !important; 
            background-color: #1e3d2f !important;
            color: #4cd964 !important;
            font-weight: bold !important;
            margin-top: 15px;
        }
        section[data-testid="stSidebar"] [data-testid="stDownloadButton"] button:hover {
            background-color: #23a55a !important;
            color: white !important;
            border-left: 6px solid #ffffff !important;
        }
    </style>
""", unsafe_allow_html=True)

# Definição e inicialização de variáveis globais de estado
if "dados_inventario" not in st.session_state:
    st.session_state["dados_inventario"] = carregar_dados()

if "tela_atual" not in st.session_state:
    st.session_state["tela_atual"] = "DASHBOARD"

if "estoque_bobinas_multi" not in st.session_state:
    st.session_state["estoque_bobinas_multi"] = carregar_estoque_por_tamanho()

thin_side = Side(border_style="thin", color="000000")
border_tabela = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

# Dicionário estático de cabeçalhos de tabelas
colunas_por_modulo = {
    "SERVIDORES": ["Categoria", "Nome do Servidor", "N° do Rack", "Endereço", "Data/Hora"],
    "REDE / IP": ["Categoria", "Nome do Equipamento","Endereço IP", "Status (Usado / Vago)",
                  "Data/Hora"],
    "PINPAD": ["Categoria", "Local", "Modelo", "Serial", "Data/Hora"],
    "FERRAMENTAS": ["Categoria", "Nome", "Descrição", "Quantidade", "Serial", "Data/Hora"],
    "BOBINAS": ["Categoria", "Tipo de Bobina", "Local/Entrada", "Quantidade Retirada", "Quem Trocou", "Observação",
                "Data/Hora"],
    "CHAMADOS": ["Categoria", "Empresa", "N° do Chamado", "Descrição do Chamado", "Analista/Técnico", "Status",
                 "Data/Hora"]
}


# --- MOTOR CENTRALIZADO DO EXCEL CORRIGIDO (ZERA O ERRO UNRESOLVED REFERENCE) ---
def gerar_arquivo_excel(df_dados):
    todas_categorias = ["DASHBOARD_Graficos", "SERVIDORES", "REDE - IP", "PINPAD", "FERRAMENTAS", "BOBINAS", "CHAMADOS"]
    output_buffer = io.BytesIO()

    # Definição de fontes corporativas limpas
    fonte_titulo_dashboard = Font(name="Segoe UI", size=11, bold=True, color="1F497D")
    fonte_dados_dashboard = Font(name="Segoe UI", size=11, bold=False)

    # Estilização profissional para os cabeçalhos das tabelas de dados
    fill_cabecalho = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    fonte_cabecalho = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")

    # Alinhamento centralizado para os dados ficarem organizados nas células
    alinhamento_centro = Alignment(horizontal="center", vertical="center", wrap_text=True)

    with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
        if df_dados.empty:
            for cat in todas_categorias:
                nome_aba = cat[:30]
                modulo_nome = "REDE / IP" if cat == "REDE - IP" else cat
                cols = colunas_por_modulo.get(modulo_nome, ["Categoria"])
                pd.DataFrame(columns=cols).to_excel(writer, index=False, sheet_name=nome_aba)
        else:
            pd.DataFrame().to_excel(writer, index=False, sheet_name="DASHBOARD_Graficos")

            for cat in todas_categorias:
                if cat != "DASHBOARD_Graficos":
                    modulo_nome = "REDE / IP" if cat == "REDE - IP" else cat
                    sub_df = df_dados[df_dados["Categoria"] == modulo_nome]
                    nome_aba = cat[:30]
                    cols_certas = colunas_por_modulo.get(modulo_nome, ["Categoria"])

                    if not sub_df.empty:
                        cols_existentes = [c for c in cols_certas if c in sub_df.columns]
                        if cols_existentes:
                            sub_df[cols_existentes].to_excel(writer, index=False, sheet_name=nome_aba)
                        else:
                            sub_df.dropna(how='all', axis=1).to_excel(writer, index=False, sheet_name=nome_aba)
                    else:
                        pd.DataFrame(columns=cols_certas).to_excel(writer, index=False, sheet_name=nome_aba)

            workbook = writer.book
            graficos_sheet = workbook["DASHBOARD_Graficos"]

            resumo_df = df_dados.groupby("Categoria").size().reset_index(name="Quantidade")

            # Escreve os cabeçalhos ocultos de apoio com fonte profissional
            celula_cat = graficos_sheet.cell(row=1, column=25, value="Categoria")
            celula_cat.font = fonte_titulo_dashboard

            celula_qtd = graficos_sheet.cell(row=1, column=26, value="Quantidade")
            celula_qtd.font = fonte_titulo_dashboard

            for i, (cat_val, qtd_val) in enumerate(zip(resumo_df["Categoria"], resumo_df["Quantidade"]), start=2):
                c1 = graficos_sheet.cell(row=i, column=25, value=str(cat_val))
                c2 = graficos_sheet.cell(row=i, column=26, value=int(qtd_val))
                c1.font = fonte_dados_dashboard
                c2.font = fonte_dados_dashboard

            total_linhas = len(resumo_df) + 1
            if total_linhas > 1:
                labels = Reference(graficos_sheet, min_col=25, min_row=2, max_row=total_linhas)
                data = Reference(graficos_sheet, min_col=26, min_row=1, max_row=total_linhas)

                # Gráfico de Pizza limpo
                pie = PieChart()
                pie.add_data(data, titles_from_data=True)
                pie.set_categories(labels)
                pie.title = "Distribuição por Categoria"
                graficos_sheet.add_chart(pie, "B2")

                # Gráfico de Barras limpo
                bar = BarChart()
                bar.type = "col"
                bar.style = 10
                bar.add_data(data, titles_from_data=True)
                bar.set_categories(labels)
                bar.title = "Volume Total por Módulo"
                graficos_sheet.add_chart(bar, "B18")

            # --- PROCESSAMENTO VISUAL DE TODAS AS ABAS DE DADOS ---
            cores_lista = ["4F81BD", "3388FF", "9933FF", "FF3399", "FF9933", "FFFF33", "FF3333"]
            for idx, cat_name in enumerate(todas_categorias):
                nome_aba = cat_name[:30]
                if nome_aba in workbook.sheetnames:
                    sheet = workbook[nome_aba]
                    sheet.sheet_properties.tabColor = cores_lista[idx % len(cores_lista)]

                    # Aplica a formatação em todas as abas de dados (menos na aba do Dashboard)
                    if nome_aba != "DASHBOARD_Graficos":
                        # Aumenta ligeiramente a altura da linha do cabeçalho para dar espaço ao texto
                        sheet.row_dimensions[1].height = 24

                        for row in sheet.iter_rows(min_row=1, max_row=sheet.max_row, min_col=1,
                                                   max_col=sheet.max_column):
                            for cell in row:
                                cell.border = border_tabela
                                cell.alignment = alinhamento_centro

                                if cell.row == 1:
                                    # Formatação do Cabeçalho (Azul escuro com texto Branco)
                                    cell.fill = fill_cabecalho
                                    cell.font = fonte_cabecalho
                                else:
                                    # Formatação das linhas de dados normais
                                    cell.font = Font(name="Segoe UI", size=10, bold=False)

                        # 🌟 LOGICA ATUALIZADA E SEGURA PARA O AUTO-AJUSTE UNIVERSAL DE COLUNAS
                        for col_idx, col in enumerate(sheet.columns, start=1):
                            max_len = 0
                            col_letter = get_column_letter(col_idx)
                            for cell in col:
                                if cell.value:
                                    max_len = max(max_len, len(str(cell.value)))
                            # Define a largura ideal adicionando uma folga confortável nas laterais
                            sheet.column_dimensions[col_letter].width = max(max_len + 5, 14)

    return output_buffer.getvalue()


# NAVEGAÇÃO COMPLETA DA SIDEBAR LATERAL
with st.sidebar:
    st.markdown("<h2 style='color: #00ffcc; font-family: \"Arial Black\"; margin-bottom: 0px;'>🏢 VillageMall</h2>",
                unsafe_allow_html=True)
    st.markdown("<p style='color: #ff007f; font-weight: bold; letter-spacing: 1px;'>🚀 GESTÃO OPERACIONAL</p>",
                unsafe_allow_html=True)
    st.markdown("<br>", unsafe_allow_html=True)

    if st.button("📊 DASHBOARD", use_container_width=True): st.session_state["tela_atual"] = "DASHBOARD"
    if st.button("🖥️ SERVIDORES", use_container_width=True): st.session_state["tela_atual"] = "SERVIDORES"
    if st.button("🌐 REDE / IP", use_container_width=True): st.session_state["tela_atual"] = "REDE / IP"
    if st.button("💳 PINPAD", use_container_width=True): st.session_state["tela_atual"] = "PINPAD"
    if st.button("🛠️ FERRAMENTAS", use_container_width=True): st.session_state["tela_atual"] = "FERRAMENTAS"
    if st.button("🧻 BOBINAS", use_container_width=True): st.session_state["tela_atual"] = "BOBINAS"
    if st.button("📞 CHAMADOS", use_container_width=True): st.session_state["tela_atual"] = "CHAMADOS"

    df_atual = pd.DataFrame(st.session_state["dados_inventario"])
    excel_data = gerar_arquivo_excel(df_atual)

    st.download_button(
        label="📥 EXPORTAR PARA EXCEL",
        data=excel_data,
        file_name="Inventario_VillageMall.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

tela = st.session_state["tela_atual"]
nome_filtro = "REDE - IP" if tela == "REDE / IP" else tela
cols_modulo = colunas_por_modulo.get(tela if tela != "REDE / IP" else "REDE - IP", ["Categoria"])

if tela == "DASHBOARD":
    st.title("📊 Painel Indicador Geral")
    if not df_atual.empty:
        resumo = df_atual.groupby("Categoria").size().reset_index(name="Total")
        col_grafico1, col_grafico2 = st.columns(2)
        with col_grafico1:
            fig_bar = go.Figure(data=[go.Bar(x=resumo["Categoria"], y=resumo["Total"], marker_color="#00ffcc")])
            fig_bar.update_layout(title="Volume por Subgrupo Operacional (Barras)", template="plotly_dark")
            st.plotly_chart(fig_bar, use_container_width=True)
        with col_grafico2:
            fig_pie = go.Figure(data=[go.Pie(labels=resumo["Categoria"], values=resumo["Total"], hole=0.3,
                                             marker=dict(colors=["#00ffcc", "#3388ff", "#9933ff", "#ff3399", "#ff9933",
                                                                 "#ffff33", "#ff3333"]))])
            fig_pie.update_layout(title="Distribuição por Módulo (Pizza)", template="plotly_dark")
            st.plotly_chart(fig_pie, use_container_width=True)
    else:
        st.info("Nenhum dado cadastrado para exibir gráficos operacionais.")

else:
    st.title(f"📂 Gerenciamento: {tela}")

    ops_b = ["58mmx300m ENTRADA", "76mmx300m PAGAMENTO", "79mmx80m CAIXA ASSISTIDA", "57mmx08m PAPA FILA"]
    ops_t = ["L1", "VIPR1", "ViPR2", "SS1R1", "SS1R2", "SS1R3", "SS2R11", "SS2R2", "SS2R3", "PAGAMENTO PORTARIA B",
             "PORTARIA C", "VIP PANORAMICO", "SS1 AMÉRICAS", "SS1 PANORAMICO", "SS1 LAGOA1", "SS1 LAGOA2",
             "SS1 LAGOA3", "SS2 PANORAMICO", "TPC AMÉRICAS", "TPC PANORAMICO", "TPC LAGOA"]
    ops_empresas = ["WPS", "PERTO", "HG", "CLUB CAR", "PRIME"]

    # FORMULÁRIO COM DELEGAÇÃO ESTÁTICA FIXA E EXPLÍCITA
    with st.form(f"form_{tela}", clear_on_submit=True):
        st.subheader("Adicionar Novo Registro")
        novo_registro = {"Categoria": nome_filtro}

        if tela == "SERVIDORES":
            novo_registro["Nome do Servidor"] = st.text_input("Nome do Servidor")
            novo_registro["N° do Rack"] = st.text_input("N° do Rack")
            novo_registro["Endereço"] = st.text_input("Endereço")
        elif tela == "REDE / IP":
            novo_registro["Nome do Equipamento"] = st.text_input("Nome do Equipamento")
            novo_registro["Status do IP"] = st.text_input("Status do IP")
            novo_registro["Endereço IP"] = st.text_input("Endereço IP")
            novo_registro["Status (Usado / Vago)"] = st.selectbox("Status (Usado / Vago)", ["Usado", "Vago"])
        elif tela == "PINPAD":
            novo_registro["Local"] = st.text_input("Local")
            novo_registro["Modelo"] = st.text_input("Modelo")
            novo_registro["Serial"] = st.text_input("Serial")
        elif tela == "FERRAMENTAS":
            novo_registro["Nome"] = st.text_input("Nome")
            novo_registro["Descrição"] = st.text_input("Descrição")
            novo_registro["Quantidade"] = st.text_input("Quantidade")
            novo_registro["Serial"] = st.text_input("Serial")
        elif tela == "BOBINAS":
            novo_registro["Tipo de Bobina"] = st.selectbox("Tipo de Bobina", ops_b)
            novo_registro["Local/Entrada"] = st.selectbox("Local/Entrada", ops_t)
            novo_registro["Quantidade Retirada"] = st.number_input("Quantidade Retirada", min_value=1, value=1, step=1)
            novo_registro["Quem Trocou"] = st.text_input("Quem Trocou")
            novo_registro["Observação"] = st.text_input("Observação")

            f_chave = novo_registro["Tipo de Bobina"]
            f_saldo = st.session_state["estoque_bobinas_multi"].get(f_chave, 0)
            st.warning(f"Estoque Disponível para o tamanho ({f_chave}): {f_saldo} unidades")
        elif tela == "CHAMADOS":
            novo_registro["Empresa"] = st.selectbox("Empresa", ops_empresas)
            novo_registro["N° do Chamado"] = st.text_input("N° do Chamado")
            novo_registro["Descrição do Chamado"] = st.text_input("Descrição do Chamado")
            novo_registro["Analista/Técnico"] = st.text_input("Analista/Técnico")
            novo_registro["Status"] = st.selectbox("Status", ["Aberto", "Em Atendimento", "Concluído"])

        if st.form_submit_button("💾 Salvar Registro"):
            if tela == "BOBINAS":
                try:
                    f_chave = novo_registro["Tipo de Bobina"]
                    baixa_qtd = int(novo_registro.get("Quantidade Retirada", 1))
                    saldo_existente = st.session_state["estoque_bobinas_multi"].get(f_chave, 0)
                    st.session_state["estoque_bobinas_multi"][f_chave] = max(0, int(saldo_existente) - baixa_qtd)
                    salvar_estoque_por_tamanho(st.session_state["estoque_bobinas_multi"])
                except:
                    pass

            novo_registro["Data/Hora"] = datetime.now().strftime("%d/%m/%Y %H:%M")
            st.session_state["dados_inventario"].append(novo_registro)
            salvar_dados(st.session_state["dados_inventario"])

            df_auto = pd.DataFrame(st.session_state["dados_inventario"])
            with open(ARQUIVO_EXCEL_AUTO, "wb") as f_l:
                f_l.write(gerar_arquivo_excel(df_auto))

            st.success("Dados salvos e planilha Excel atualizada automaticamente!")
            st.rerun()

    # --- TABELAS E COMPONENTES INFERIORES DE VISUALIZAÇÃO ---
st.markdown("<br>", unsafe_allow_html=True)

if tela != "DASHBOARD":
    if tela == "BOBINAS":
        aba_tabela, aba_estoque_manual = st.tabs(["📋 HISTÓRICO DE TROCAS", "📦 ABASTECER ESTOQUE"])

        with aba_tabela:
            st.subheader("Registros Atuais")

            if not df_atual.empty and "Categoria" in df_atual.columns:
                df_filtrado = df_atual[df_atual["Categoria"] == nome_filtro]
            else:
                df_filtrado = pd.DataFrame(columns=cols_modulo)

            if not df_filtrado.empty:
                cols_exibir = [c for c in cols_modulo if c in df_filtrado.columns]
                df_com_botoes = df_filtrado[cols_exibir].copy()
                df_com_botoes["Editar"] = "✏️"
                df_com_botoes["Excluir"] = "🗑️"

                chave_tabela = "data_editor_bobinas_exclusivo"
                st.data_editor(df_com_botoes, use_container_width=True, key=chave_tabela)
            else:
                st.info("Nenhum item adicionado no subgrupo BOBINAS até o momento.")

        with aba_estoque_manual:
            st.subheader("⚙️ Entrada Manual de Novas Bobinas")
            col_tamanho_add, col_qtd_add, col_btn_add = st.columns(3)

            with col_tamanho_add:
                tamanho_abastecer = st.selectbox("Selecione o Tamanho Técnico da Bobina:", ops_b,
                                                 key="tamanho_abastecer_manual")

            with col_qtd_add:
                saldo_manual_atual = st.session_state["estoque_bobinas_multi"].get(tamanho_abastecer, 0)
                quantidade_entrada = st.number_input(f"Saldo Atual: {saldo_manual_atual} un. Quantidade recebida:",
                                                     min_value=1, value=10, step=1, key="qtd_abastecer_manual")

            with col_btn_add:
                st.markdown("<br>", unsafe_allow_html=True)
                if st.button("📥 Adicionar ao Estoque", use_container_width=True, key="btn_abastecer_manual"):
                    st.session_state["estoque_bobinas_multi"][tamanho_abastecer] = int(saldo_manual_atual) + int(
                        quantidade_entrada)
                    salvar_estoque_por_tamanho(st.session_state["estoque_bobinas_multi"])

                    df_auto = pd.DataFrame(st.session_state["dados_inventario"])
                    with open(ARQUIVO_EXCEL_AUTO, "wb") as f_l:
                        f_l.write(gerar_arquivo_excel(df_auto))

                    st.success(f"Estoque do tamanho {tamanho_abastecer} abastecido com sucesso!")
                    st.rerun()

    else:
        st.subheader("Registros Atuais")

        if not df_atual.empty and "Categoria" in df_atual.columns:
            df_filtrado = df_atual[df_atual["Categoria"] == nome_filtro]
        else:
            df_filtrado = pd.DataFrame(columns=cols_modulo)

        if not df_filtrado.empty:
            cols_exibir = [c for c in cols_modulo if c in df_filtrado.columns]
            df_com_botoes = df_filtrado[cols_exibir].copy()
            df_com_botoes["Editar"] = "✏️"
            df_com_botoes["Excluir"] = "🗑️"

            chave_tabela = f"data_editor_estavel_{tela.lower().replace(' ', '_')}_final"
            st.data_editor(df_com_botoes, use_container_width=True, key=chave_tabela)
        else:
            st.info("Nenhum item adicionado neste módulo.")
