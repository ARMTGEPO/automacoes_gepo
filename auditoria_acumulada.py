import streamlit as st
import pandas as pd
import tempfile

# === Estilo visual dos botões ===
st.markdown("""
    <style>
    [data-testid="stBaseButton-secondary"] {
        background-color: #28a745 !important;
        color: white !important;
        font-weight: bold;
        border: none;
    }
    [data-testid="stBaseButton-secondary"]:hover {
        background-color: #218838 !important;
        color: white !important;
    }
    </style>
""", unsafe_allow_html=True)

st.logo(image="img/senac_logo.png")

# === Função para gerar Excel com inconsistências ===
def gerar_excel_inconsistencias(inconsistencias, output_path="relatorio_inconsistencias_acumulada.xlsx"):
    with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
        for nome, df in inconsistencias.items():
            sheet_name = nome[:31]
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            worksheet = writer.sheets[sheet_name]
            for i, col in enumerate(df.columns):
                max_len = max(df[col].astype(str).apply(len).max(), len(col)) + 2
                worksheet.set_column(i, i, max_len)


# === Função principal de validação ===
def validar_e_gerar_relatorio(df_acc):
    inconsistencias = {}

    # Selecionar colunas relevantes (garante que todas existam)
    colunas = [
        "Unidade Operativa da Turma",
        "CPF do Aluno",
        "Data da Matrícula",
        "Matrícula",
        "Nome do Aluno",
        "Estado da Turma",
        "Sigla da Turma",
        "Turno Da Turma",
        "Turma",
        "Título do Curso",
        "Inicio da Execução da Turma",
        "Termino da Execução da Turma",
        "Estado da Matrícula do Aluno",
        "Data de Lançamento do Estado da Matrícula",
        "Data de Ocorrência do Estado da Matrícula",
        "Modalidade Recurso",
    ]

    # Somente colunas existentes
    colunas_existentes = [c for c in colunas if c in df_acc.columns]
    df = df_acc[colunas_existentes].copy()

    # Converter datas
    for c in [
        "Inicio da Execução da Turma",
        "Termino da Execução da Turma",
        "Data de Lançamento do Estado da Matrícula",
        "Data de Ocorrência do Estado da Matrícula",
    ]:
        if c in df.columns:
            df[c] = pd.to_datetime(df[c], errors='coerce')

    # =====================================================
    # === 1️⃣ ALUNOS FAZENDO O MESMO TÍTULO CONCOMITANTE ===
    # =====================================================
    if {"CPF do Aluno", "Título do Curso", "Turma"}.issubset(df.columns):
        df_merged = df.merge(df, on=["CPF do Aluno", "Título do Curso"], suffixes=("_1", "_2"))
        df_merged = df_merged[df_merged["Turma_1"] != df_merged["Turma_2"]]

        cond_concomitancia = (
            (df_merged["Inicio da Execução da Turma_1"] <= df_merged["Termino da Execução da Turma_2"]) &
            (df_merged["Inicio da Execução da Turma_2"] <= df_merged["Termino da Execução da Turma_1"])
        )

        df_concomitantes = df_merged[cond_concomitancia].copy()

        # Remover duplicidades (A-B / B-A)
        df_concomitantes["pair_id"] = df_concomitantes.apply(
            lambda x: "_".join(sorted([x["Turma_1"], x["Turma_2"]])),
            axis=1
        )
        df_concomitantes = df_concomitantes.drop_duplicates(subset=["CPF do Aluno", "Título do Curso", "pair_id"])
        df_concomitantes.drop(columns="pair_id", inplace=True)

        # Filtrar estados inválidos
        if "Estado da Matrícula do Aluno_1" in df_concomitantes.columns:
            excluidos = ["17 - Matricula Cancelada", "15 - Transferência Externa"]
            df_concomitantes = df_concomitantes[
                (~df_concomitantes["Estado da Matrícula do Aluno_1"].isin(excluidos)) &
                (~df_concomitantes["Estado da Matrícula do Aluno_2"].isin(excluidos))
            ]

        inconsistencias['Alunos fazendo o mesmo título de forma concomitante'] = df_concomitantes.copy()

    # =======================================
    # === 2️⃣ TRATAMENTO: TRIPLO PSG ========
    # =======================================
    if "Modalidade Recurso" in df.columns:
        # ✅ Compatível com todas as versões do pandas
        df_psg = df[df["Modalidade Recurso"].fillna("").str.upper() == "PSG"].copy()

        # Eliminar duplicados de mesma matrícula
        if {"CPF do Aluno", "Matrícula"}.issubset(df_psg.columns):
            df_psg = df_psg.drop_duplicates(subset=["CPF do Aluno", "Matrícula"])

        # Função para detectar sobreposição tripla
        def turmas_concomitantes(grupo):
            eventos = []
            for idx, row in grupo.iterrows():
                eventos.append((row["Inicio da Execução da Turma"], +1, idx))
                eventos.append((row["Termino da Execução da Turma"], -1, idx))

            eventos.sort(key=lambda x: x[0])
            count = 0
            ativos = set()
            conflitos = []

            for data, delta, idx in eventos:
                if pd.isna(data):
                    continue
                if delta == +1:
                    ativos.add(idx)
                    count += 1
                else:
                    count -= 1
                    ativos.discard(idx)

                if count >= 3:
                    conflitos.append(list(ativos))

            if conflitos:
                indices = set().union(*conflitos)
                return grupo.loc[list(indices)]
            else:
                return pd.DataFrame()

        conflitos = df_psg.groupby("CPF do Aluno", group_keys=False).apply(turmas_concomitantes)

        if not conflitos.empty:
            resumo_conflitos = conflitos[
                [
                    "Unidade Operativa da Turma",
                    "Inicio da Execução da Turma",
                    "Termino da Execução da Turma",
                    "Turma",
                    "Título do Curso",
                    "Matrícula",
                    "CPF do Aluno",
                    "Nome do Aluno",
                    "Modalidade Recurso",
                    "Estado da Matrícula do Aluno",
                    "Estado da Turma",
                    "Data de Lançamento do Estado da Matrícula",
                    "Data de Ocorrência do Estado da Matrícula",
                ]
            ].sort_values(["CPF do Aluno", "Inicio da Execução da Turma"])

            inconsistencias["Triplo PSG"] = resumo_conflitos.copy()

    return inconsistencias


# === INTERFACE STREAMLIT ===
st.title("📝 Auditoria de Produção Acumulada")

username = st.session_state.get('username', 'usuário')
st.markdown(f"Seja bem-vindo, **{username}**. Faça o upload do relatório de produção acumulada para iniciar a auditoria.")

with st.expander("📖 Análises Realizadas", expanded=True):
    st.write("""
    1. **Triplo PSG** → Verifica se alunos com recurso PSG estão em mais de duas turmas com ocorrência concomitante.  
    2. **Mesmo título concomitante** → Identifica alunos matriculados em turmas diferentes do mesmo título no mesmo período.
    """)

st.sidebar.write("Para iniciar, faça o upload da base de produção acumulada.")
uploaded_file = st.sidebar.file_uploader("📂 Selecione o arquivo CSV", type=["csv"])

if uploaded_file:
    try:
        df = pd.read_csv(uploaded_file, sep=';', encoding='latin1')
        st.dataframe(df.head(5))

        # Converter datas
        for c in [
            "Inicio da Execução da Turma",
            "Termino da Execução da Turma",
            "Data de Lançamento do Estado da Matrícula",
            "Data de Ocorrência do Estado da Matrícula",
        ]:
            if c in df.columns:
                df[c] = pd.to_datetime(df[c], dayfirst=True, errors='coerce')

        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_excel:
            inconsistencias = validar_e_gerar_relatorio(df)
            gerar_excel_inconsistencias(inconsistencias, tmp_excel.name)

            if inconsistencias:
                with open(tmp_excel.name, "rb") as f_excel:
                    st.sidebar.download_button(
                        label="📊 Baixar Relatório Excel",
                        data=f_excel,
                        file_name="relatorio_inconsistencias.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                        type='secondary',
                        icon=":material/download:"
                    )
                st.sidebar.success("✅ Relatórios gerados com sucesso!")
            else:
                st.sidebar.info("Nenhuma inconsistência encontrada.")

    except Exception as e:
        st.error(f"❌ Erro ao processar o arquivo: {e}")
