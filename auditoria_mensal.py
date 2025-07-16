import streamlit as st
import pandas as pd
from fpdf import FPDF, XPos, YPos
from datetime import datetime
import tempfile
from email_validator import validate_email, EmailNotValidError

st.set_page_config(page_title="Auditoria de Produção", layout="wide", page_icon=":bar_chart")
st.logo(image="img/senac_logo.png")

st.markdown("""
    <style>
    /* Aplica a cor verde ao botão com base no data-testid */
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



# === Função para gerar Excel com inconsistências ===
def gerar_excel_inconsistencias(inconsistencias, output_path="relatorio_inconsistencias.xlsx"):
    with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
        for nome, df in inconsistencias.items():
            # Limitar o nome da aba para 31 caracteres (limitação do Excel)
            sheet_name = nome[:31]
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            
            # Ajustar largura das colunas
            worksheet = writer.sheets[sheet_name]
            for i, col in enumerate(df.columns):
                max_len = max(df[col].astype(str).apply(len).max(), len(col)) + 2
                worksheet.set_column(i, i, max_len)
        writer.close()

# === Função de validação e geração do relatório ===
def validar_e_gerar_relatorio(df, primeiro_dia_mes_anterior, output_path="relatorio_inconsistencias.pdf"):
    inconsistencias = {}
    estados_efetivados = ["1 - Aprovado", "2 - Reprovado", "3 - Evadido"]

    # Parâmetro - Considerar apenas turmas de ação educacional
    df = df[df['Quadro de Ação'] == "Ação Educacional"]
    df['Unidade Operativa da Turma'] = df['Unidade Operativa da Turma'].str.split("SENAC CEP ").str[-1]

    # Regra 1 - Identificar matrículas desistentes com carga horária contabilizada. - Desistentes com CH Apurada
    mask1 = (df["Estado conf. CODEPE"] == "4 - Desistente") & (df["CH_Apurada_Mes"] > 0)
    inconsistencias['Desistentes com CH Apurada'] = df.loc[mask1, [
        "Matrícula", "Turma", "CPF do Aluno", "Estado conf. CODEPE", "CH_Apurada_Mes"
    ]]

    # Regra 2 -Turmas Encerradas em Exercício Anterior contabilizando CH
    df_filtrado_2 = df[df["Estado conf. CODEPE"] != "4 - Desistente"]
    
    mask2 = (df["Termino da Execução da Turma"] < primeiro_dia_mes_anterior) & (df["CH_Apurada_Mes"] != 0.00)
    inconsistencias['Turmas Encerradas em Exercício Anterior contabilizando CH'] = df.loc[mask2, [
        "Unidade Operativa da Turma", "Matrícula", "Turma", "Nome do Aluno", 
        "Estado da Turma", "Estado conf. CODEPE", "CH_Apurada_Mes", "Termino da Execução da Turma"
    ]]

    # Regra 3  - Matrículas que sofreram ajustes de CH na competência
    mask3 = (df_filtrado_2["Ajuste_Senac_Mes"] != 0.00)
    inconsistencias['Matrículas com ajuste de CH'] = df_filtrado_2.loc[mask3, [
        "Unidade Operativa da Turma", "Matrícula", "Turma", "Nome do Aluno", 
        "Estado da Turma", "Estado conf. CODEPE", "Estado da Matrícula do Aluno", "CH_Apurada_Mes"
    ]]

    # Regra 4 - Identificar múltiplas matrículas ativas por CPF na mesma turma.
    df_filtrado_2 = df[df["Estado da Matrícula do Aluno"] != "Matrícula Cancelada"]
    mask4 = df_filtrado_2.groupby(["CPF do Aluno", "Turma"])["Matrícula"].transform('nunique') > 1
    inconsistencias['CPF com Múltiplas Matrículas na mesma turma'] = df_filtrado_2.loc[mask4, [
        "Unidade Operativa da Turma", "Matrícula", "Turma", "CPF do Aluno", "Estado conf. CODEPE", 
        "Estado da Matrícula do Aluno", "Recurso Financeiro", "CH_Apurada_Mes", "Condição"
    ]].sort_values(["CPF do Aluno", "Turma"])

    # Regra 5 - Identificar matrículas de um mesmo CPF em mais de 02 turmas PSG.
    mask_psg = (df["Recurso Financeiro"] == "PSG")
    cpf_counts = df.loc[mask_psg, "CPF do Aluno"].value_counts().reset_index()
    cpf_counts.columns = ["CPF do Aluno", "count"]
    cpf_invalidos = cpf_counts[cpf_counts["count"] > 2]["CPF do Aluno"]
    mask5 = df["CPF do Aluno"].isin(cpf_invalidos) & mask_psg
    inconsistencias['CPF com > 2 Turmas PSG'] = df.loc[mask5, [
        "Matrícula", "Turma", "CPF do Aluno", "Recurso Financeiro", 
        "Estado conf. CODEPE", "CH_Apurada_Mes"
    ]].sort_values("CPF do Aluno")

    # Regra 6 - Identificar matrículas efetivadas em turmas em processo. Entende-se como matrículas efetivadas aquelas com Estado CODEPE "Aprova, Reprovada ou Evadida".
    mask6 = df["Estado conf. CODEPE"].isin(estados_efetivados) & df["Estado da Turma"].str.contains("Em Processo")
    inconsistencias['Efetivados em Turmas Ativas'] = df.loc[mask6, [
        "Unidade Operativa da Turma", "Matrícula", "Turma", "CPF do Aluno", "Estado conf. CODEPE", 
        "Estado da Turma", "CH_Apurada_Mes"
    ]]

    # Regra 7 - Validação Rápida de E-mails
    df["Data de Nascimento do Aluno"] = pd.to_datetime(df["Data de Nascimento do Aluno"], errors='coerce', dayfirst=True)
    hoje = pd.to_datetime(datetime.today())
    df["idade"] = ((hoje - df["Data de Nascimento do Aluno"]).dt.days / 365.25).fillna(0)

    # Aplicando somente para alunos maiores de 18
    df_maiores = df[df["idade"] >= 18].copy()

    # Lista de domínios comuns incorretos
    dominios_proibidos = [
        "alunosenac.com", "ghotmail.com", "hotail.com", "gamil.com",
        "outloook.com", "gmial.com", "002gmail.com", "77gmail.com","g.mail.com", "verificar.com",
        "verificar.com.br", "edu.mt.gov.bt", "gmal.com", "gmai.com", "hotimail.com", "homail.comm",
        "verfificar.com", "gamail.com", "gmail.co", "33gmail.com",
        "gmil.com", "gmail.com.com", "yaool.com",
        "gmail.coom", "00gmail.com", "outlokk.com", "edu.mt.gv.br",
        "edu.mtgov.br", "gmail.cocm", "84gmail.com"
    ]

    def email_invalido(email):
        if not isinstance(email, str) or email.strip().lower() in ["", "na", "null", "nan", "-", "--"]:
            return True
        try:
            # Validação apenas de sintaxe
            email_info = validate_email(email, check_deliverability=False)
            dominio = email_info.domain.lower()
            if dominio in dominios_proibidos:
                return True
        except EmailNotValidError:
            return True
        return False

    # Aplicando a verificação nos maiores de idade
    df_maiores["Email Invalido"] = df_maiores["E-mail"].apply(email_invalido)

    # Armazenando a inconsistência
    inconsistencias['E-mails Inválidos'] = df_maiores[df_maiores["Email Invalido"] == True][[
        "Unidade Operativa da Turma", "Matrícula", "Turma", "Nome do Aluno", "CPF do Aluno", "E-mail"
    ]]
    
    # Geração do PDF
    pdf = FPDF(orientation='L')
    pdf.set_auto_page_break(auto=True, margin=15)

    table_header_color = (100, 180, 255)
    row_color = (220, 230, 240)
    alternate_row_color = (240, 248, 255)

    competencia = primeiro_dia_mes_anterior.strftime('%m/%Y')
    pdf.add_page()
    pdf.set_font('Helvetica', 'B', 24)
    pdf.cell(0, 40, "Relatório de Inconsistências", new_x=XPos.LMARGIN, new_y=YPos.NEXT, align='C')
    pdf.set_font('Helvetica', '', 16)
    pdf.cell(0, 20, f"Data: {datetime.today().strftime('%d/%m/%Y')}", new_x=XPos.LMARGIN, new_y=YPos.NEXT, align='C')
    pdf.cell(0, 20, f"Competência: {competencia}", new_x=XPos.LMARGIN, new_y=YPos.NEXT, align='C')
    pdf.cell(0, 20, f"Total de Validações: {len(inconsistencias)}", new_x=XPos.LMARGIN, new_y=YPos.NEXT, align='C')

    pdf.add_page()
    pdf.set_font('Helvetica', 'B', 18)
    pdf.cell(0, 10, "Validações Realizadas", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.ln(10)

    max_tipo_width = max([pdf.get_string_width(tipo) for tipo in inconsistencias.keys()]) + 10
    summary_width = min(max(max_tipo_width, 100), 220)
    pdf.set_font('Helvetica', 'B', 10)
    pdf.set_fill_color(*table_header_color)
    pdf.cell(summary_width, 8, "Tipo de Inconsistência", border=1, align='C', fill=True)
    pdf.cell(30, 8, "Registros", border=1, new_x=XPos.LMARGIN, new_y=YPos.NEXT, align='C', fill=True)

    pdf.set_font('Helvetica', '', 9)
    for i, (tipo, df_inc) in enumerate(inconsistencias.items()):
        fill = alternate_row_color if i % 2 else row_color
        pdf.set_fill_color(*fill)
        display_text = tipo
        if pdf.get_string_width(tipo) > summary_width - 5:
            while pdf.get_string_width(display_text + "...") > summary_width - 5 and len(display_text) > 10:
                display_text = display_text[:-1]
            display_text = display_text.rstrip() + "..." if len(display_text) < len(tipo) else display_text
        pdf.cell(summary_width, 7, display_text, border=1, align='L', fill=True)
        pdf.cell(30, 7, str(len(df_inc)), border=1, new_x=XPos.LMARGIN, new_y=YPos.NEXT, align='C', fill=True)

    for tipo, df_inc in inconsistencias.items():
        if not df_inc.empty:
            pdf.add_page()
            pdf.set_font('Helvetica', 'B', 14)
            pdf.cell(0, 8, f"Inconsistência: {tipo}", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
            pdf.ln(3)
            pdf.set_font('Helvetica', '', 8)
            col_widths = []
            for col in df_inc.columns:
                header_width = pdf.get_string_width(str(col)) + 4
                sample_values = df_inc[col].dropna().sample(min(20, len(df_inc)), random_state=1)
                max_value_width = max([pdf.get_string_width(str(val)) + 4 for val in sample_values]) if not sample_values.empty else 0
                col_width = max(header_width, max_value_width, 20)
                col_widths.append(col_width)

            total_width = sum(col_widths)
            if total_width > 280:
                scale_factor = 280 / total_width
                col_widths = [w * scale_factor for w in col_widths]

            pdf.set_font('Helvetica', 'B', 8)
            pdf.set_fill_color(*table_header_color)
            for col, width in zip(df_inc.columns, col_widths):
                pdf.cell(width, 6, str(col), border=1, align='C', fill=True)
            pdf.ln()

            pdf.set_font('Helvetica', '', 8)
            for row_idx, row in df_inc.iterrows():
                fill = alternate_row_color if row_idx % 2 else row_color
                pdf.set_fill_color(*fill)
                for value, width in zip(row, col_widths):
                    text = str(value)
                    if pdf.get_string_width(text) > width - 2:
                        while pdf.get_string_width(text + "...") > width - 2 and len(text) > 3:
                            text = text[:-1]
                        text = text + "..." if len(text) > 3 else text
                    pdf.cell(width, 6, text, border=1, align='C', fill=True)
                pdf.ln()

    pdf.add_page()
    pdf.set_font('Helvetica', 'I', 8)
    pdf.cell(0, 8, "Relatório gerado automaticamente pelo sistema de validação", align='C')
    pdf.cell(0, 8, f"{username}", align='C')
    pdf.output(output_path)
    
    return inconsistencias  # Retorna o dicionário de inconsistências para gerar o Excel

# === Interface Principal ===
st.title("📝 Auditoria de Produção Mensal")
username = st.session_state.get("username", "usuário")
st.markdown(f"Olá, **{username}**! Faça upload do Relatório Analítico Mensal de Produção.")

with st.expander("📖 Validações aplicadas", expanded=True):
    st.write("""
    1. **Desistentes com CH Apurada**: Verifica se alunos desistentes têm carga horária apurada no período.
    2. **Turmas Encerradas em Exercício Anterior contabilizando CH**: Identifica turmas encerradas no mês anterior com carga horária apurada no período.
    2. **Ajustes de CH realizados na competência**: Identifica os ajustes de CH no período visando identificar possíveis erros.
    3. **CPF com Múltiplas Matrículas na mesma turma**: Detecta alunos com mais de uma matrícula na mesma turma com CH para ambas as matrículas.
    4. **CPF com > 2 Turmas PSG**: Verifica se alunos com recurso PSG estão em mais de duas turmas.
    5. **Efetivados em Turmas Ativas**: Identifica alunos com estado no CODEPE "1 - Aprovado", "2 - Reprovado", "3 - Evadido" em turmas "Em processo".
    6. **E-mails Inválidos**: Valida e-mails de alunos maiores de 18 anos, verificando se estão no formato correto e se não pertencem a domínios comuns incorretos.
    Essas validações ajudam a garantir a integridade dos dados e a conformidade com as regras do sistema.
             """)

st.sidebar.write("Para iniciar, informe a competência e faça o upload do arquivo CSV obtido do SIG.")
competencia_input = st.sidebar.text_input("🗓️ Informe a competência (MM/AAAA)", placeholder="Ex: 05/2025")
uploaded_file = st.sidebar.file_uploader("📂 Selecione o arquivo CSV", type=["csv"])

if competencia_input and uploaded_file:
    try:
        competencia_dt = datetime.strptime(competencia_input, "%m/%Y")
        primeiro_dia_mes_anterior = competencia_dt.replace(day=1)

        df = pd.read_csv(uploaded_file, encoding='latin-1', sep=';')
        df["CH_Apurada_Mes"] = df["CH_Apurada_Mes"].astype(str).str.replace(',', '.').astype(float)
        df["Ajuste_Senac_Mes"] = df["Ajuste_Senac_Mes"].astype(str).str.replace(',', '.').astype(float)
        df["Termino da Execução da Turma"] = pd.to_datetime(df["Termino da Execução da Turma"], dayfirst=True, errors='coerce')

        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_pdf:
            # Primeiro geramos o PDF e obtemos as inconsistências
            inconsistencias = validar_e_gerar_relatorio(df, primeiro_dia_mes_anterior, tmp_pdf.name)
            
            # Agora geramos o Excel em um arquivo temporário separado
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_excel:
                gerar_excel_inconsistencias(inconsistencias, tmp_excel.name)
                
                with open(tmp_pdf.name, "rb") as f_pdf:
                    st.sidebar.download_button(
                        label="Baixar PDF",
                        data=f_pdf,
                        file_name="relatorio_inconsistencias.pdf",
                        mime="application/pdf",
                        use_container_width=True,
                        type='primary',
                        icon=":material/download:"
                    )

                with open(tmp_excel.name, "rb") as f_excel:
                    st.sidebar.download_button(
                        label="Baixar Excel",
                        data=f_excel,
                        file_name="relatorio_inconsistencias.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                        type='secondary',
                        icon=":material/download:"
                    )

        st.sidebar.success("✅ Relatórios gerados com sucesso!")

    except ValueError:
        st.error("❌ Competência inválida. Use o formato MM/AAAA.")
    except Exception as e:
        st.error(f"❌ Erro ao processar o arquivo: {e}")
elif uploaded_file and not competencia_input:
    st.warning("⚠️ Informe a competência antes de processar o arquivo.")
        
# Estatísticas descritivas do dataframe
if 'df' in locals():
    st.divider()
    st.subheader("📊 Estatísticas Descritivas")

    # === Filtros ===
    with st.expander("🔎 Filtrar dados"):
        col1, col2 = st.columns(2)

        with col1:
            filtro_estado_turma = st.multiselect("Estado da Turma", options=sorted(df["Estado da Turma"].dropna().unique()), placeholder="Selecione estados")
            filtro_recurso = st.multiselect("Recurso Financeiro", options=sorted(df["Recurso Financeiro"].dropna().unique()), placeholder="Selecione recursos")
            filtro_estado_matricula = st.multiselect("Estado da Matrícula do Aluno", options=sorted(df["Estado da Matrícula do Aluno"].dropna().unique()), placeholder="Selecione estados")
            filtro_condicao = st.multiselect("Condição", options=sorted(df["Condição"].dropna().unique()), placeholder="Selecione condições")
        with col2:
            filtro_estado_codepe = st.multiselect("Estado conf. CODEPE", options=sorted(df["Estado conf. CODEPE"].dropna().unique()),  placeholder="Selecione estados")
            filtro_uo = st.multiselect("Unidade Operativa da Turma", options=sorted(df["Unidade Operativa da Turma"].dropna().unique()), placeholder="Selecione unidades")
            filtro_modalidade = st.multiselect("Modalidade", options=sorted(df["Modalidade"].dropna().unique()), placeholder="Selecione modalidades")
            filtro_quadro_acao = st.multiselect("Quadro de Ação", options=sorted(df["Quadro de Ação"].dropna().unique()), placeholder="Selecione quadros", default=["Ação Educacional"])

    # Aplica os filtros
    df_filtrado = df.copy()

    if filtro_estado_turma:
        df_filtrado = df_filtrado[df_filtrado["Estado da Turma"].isin(filtro_estado_turma)]
    if filtro_recurso:
        df_filtrado = df_filtrado[df_filtrado["Recurso Financeiro"].isin(filtro_recurso)]
    if filtro_estado_matricula:
        df_filtrado = df_filtrado[df_filtrado["Estado da Matrícula do Aluno"].isin(filtro_estado_matricula)]
    if filtro_estado_codepe:
        df_filtrado = df_filtrado[df_filtrado["Estado conf. CODEPE"].isin(filtro_estado_codepe)]
    if filtro_uo:
        df_filtrado = df_filtrado[df_filtrado["Unidade Operativa da Turma"].isin(filtro_uo)]
    if filtro_modalidade:
        df_filtrado = df_filtrado[df_filtrado["Modalidade"].isin(filtro_modalidade)]
    if filtro_quadro_acao:
        df_filtrado = df_filtrado[df_filtrado["Quadro de Ação"].isin(filtro_quadro_acao)]
    if filtro_condicao:
        df_filtrado = df_filtrado[df_filtrado["Condição"].isin(filtro_condicao)]

    # === Agrupamento por Recurso Financeiro ===
    resumo = df_filtrado.groupby("Recurso Financeiro")[["CH_Apurada_Mes", "Matricula_Apurada_Mes"]].sum().reset_index()
    resumo = resumo.rename(columns={
        "CH_Apurada_Mes": "Carga Horária Total",
        "Matricula_Apurada_Mes": "Matrículas Apuradas"
    })

    # Formatação numérica estilo brasileiro
    for col in ["Carga Horária Total", "Matrículas Apuradas"]:
        resumo[col] = resumo[col].map(lambda x: f"{x:,.0f}".replace(",", "X").replace(".", ",").replace("X", "."))

    # Totais
    total_ch = resumo["Carga Horária Total"].str.replace('.', '', regex=False).str.replace(',', '.', regex=False).astype(float).sum()
    total_mat = resumo["Matrículas Apuradas"].str.replace('.', '', regex=False).str.replace(',', '.', regex=False).astype(float).sum()

    total_row = pd.DataFrame([["Total Geral", 
                            f"{total_ch:,.0f}".replace(",", "X").replace(".", ",").replace("X", "."),
                            f"{total_mat:,.0f}".replace(",", "X").replace(".", ",").replace("X", ".")]],
                            columns=resumo.columns)

    resumo = pd.concat([resumo, total_row], ignore_index=True)

    # CSS com alinhamento especial
    st.markdown("""
        <style>
        .custom-table {
            width: 100%;
            border-collapse: collapse;
            font-family: sans-serif;
        }
        .custom-table thead th {
            background-color: #f0f2f6;
            color: #333333;
            padding: 0.1rem;
            text-align: center;
            border-bottom: 1px solid #cccccc;
        }
        .custom-table tbody td {
            padding: 0.1rem;
            text-align: center;
            border-bottom: 1px solid #eaeaea;
            font-size: 15px;
        }
        .custom-table tbody td:first-child {
            text-align: left;
        }
        </style>
    """, unsafe_allow_html=True)

    # === KPI
    col1, col2, col3 = st.columns(3)
    
    # Total CH Apurada no mês
    valor_ch = round(df_filtrado["CH_Apurada_Mes"].sum(), 0)
    total_ch = f'{valor_ch:,.0f}'.replace(",", ".")

    # Total de Matrículas novas no mês
    valor_mt_nova = round(df_filtrado['Mat_Nova'].sum(), 0)
    total_mt_nova = f'{valor_mt_nova:,.0f}'.replace(',', '.')

    # Matrículas ajustadas no mês
    valor_mt_ajuste = round(df_filtrado['Ajuste_Matriculas_Mes'].sum(), 0)
    total_mt_ajuste = f'{valor_mt_ajuste:,.0f}'.replace(',','.')


    with col1:
        st.metric("#### Carga Horária Apurada no Mês", total_ch)

    with col2:
        st.metric("#### Matrículas Novas no Mês", total_mt_nova)

    with col3:
        st.metric("#### Ajuste de Matrículas no Mês", total_mt_ajuste)        

    # === TABELA 1: Recurso Financeiro ===
    st.markdown("### 📘 Resumo por Recurso Financeiro")
    st.markdown(resumo.to_html(index=False, escape=False, classes='custom-table'), unsafe_allow_html=True)

    # === TABELA 2: Modalidade ===
    # Agrupamento e soma
    resumo_modalidade = df_filtrado.groupby("Modalidade")[["CH_Apurada_Mes", "Matricula_Apurada_Mes"]].sum().reset_index()

    # Renomeando colunas
    resumo_modalidade = resumo_modalidade.rename(columns={
        "CH_Apurada_Mes": "Carga Horária Total",
        "Matricula_Apurada_Mes": "Matrículas Apuradas"
    })

    # SALVA uma cópia dos totais ANTES de formatar como string
    total_ch_m = resumo_modalidade["Carga Horária Total"].sum()
    total_mat_m = resumo_modalidade["Matrículas Apuradas"].sum()

    # Formata os valores como strings estilo brasileiro
    for col in ["Carga Horária Total", "Matrículas Apuradas"]:
        resumo_modalidade[col] = resumo_modalidade[col].map(lambda x: f"{x:,.0f}".replace(",", "X").replace(".", ",").replace("X", "."))

    # Cria a linha de totais já formatada
    total_row_modalidade = pd.DataFrame([[
        "Total Geral",
        f"{total_ch_m:,.0f}".replace(",", "X").replace(".", ",").replace("X", "."),
        f"{total_mat_m:,.0f}".replace(",", "X").replace(".", ",").replace("X", ".")
    ]], columns=resumo_modalidade.columns)

    # Adiciona ao DataFrame
    resumo_modalidade = pd.concat([resumo_modalidade, total_row_modalidade], ignore_index=True)


    st.markdown("### 📗 Resumo por Modalidade")
    st.markdown(resumo_modalidade.to_html(index=False, escape=False, classes='custom-table'), unsafe_allow_html=True)

    # === TABELA 3: Unidade Operativa ===
    resumo_uo = df_filtrado.groupby("Unidade Operativa da Turma")[["CH_Apurada_Mes", "Matricula_Apurada_Mes"]].sum().reset_index()
    resumo_uo = resumo_uo.rename(columns={
        "CH_Apurada_Mes": "Carga Horária Total",
        "Matricula_Apurada_Mes": "Matrículas Apuradas"
    })

    for col in ["Carga Horária Total", "Matrículas Apuradas"]:
        resumo_uo[col] = resumo_uo[col].map(lambda x: f"{x:,.0f}".replace(",", "X").replace(".", ",").replace("X", "."))

    total_ch_uo = resumo_uo["Carga Horária Total"].str.replace('.', '', regex=False).str.replace(',', '.', regex=False).astype(float).sum()
    total_mat_uo = resumo_uo["Matrículas Apuradas"].str.replace('.', '', regex=False).str.replace(',', '.', regex=False).astype(float).sum()

    total_row_uo = pd.DataFrame([["Total Geral", 
                                f"{total_ch_uo:,.0f}".replace(",", "X").replace(".", ",").replace("X", "."),
                                f"{total_mat_uo:,.0f}".replace(",", "X").replace(".", ",").replace("X", ".")]],
                                columns=resumo_uo.columns)

    resumo_uo = pd.concat([resumo_uo, total_row_uo], ignore_index=True)

    st.markdown("### 📙 Resumo por Unidade Operativa da Turma")
    st.markdown(resumo_uo.to_html(index=False, escape=False, classes='custom-table'), unsafe_allow_html=True)

    # === TABELA 4: Estado conf. CODEPE ===
    resumo_estado_codepe = df_filtrado.groupby("Estado conf. CODEPE")[["CH_Apurada_Mes", "Matricula_Apurada_Mes"]].sum().reset_index()
    resumo_estado_codepe = resumo_estado_codepe.rename(columns={
        "CH_Apurada_Mes": "Carga Horária Total",
        "Matricula_Apurada_Mes": "Matrículas Apuradas"
    })

    for col in ["Carga Horária Total", "Matrículas Apuradas"]:
        resumo_estado_codepe[col] = resumo_estado_codepe[col].map(lambda x: f"{x:,.0f}".replace(",", "X").replace(".", ",").replace("X", "."))

    total_ch_ec = resumo_estado_codepe["Carga Horária Total"].str.replace('.', '', regex=False).str.replace(',', '.', regex=False).astype(float).sum()
    total_mat_ec = resumo_estado_codepe["Matrículas Apuradas"].str.replace('.', '', regex=False).str.replace(',', '.', regex=False).astype(float).sum()

    total_row_ec = pd.DataFrame([["Total Geral", 
                                f"{total_ch_ec:,.0f}".replace(",", "X").replace(".", ",").replace("X", "."),
                                f"{total_mat_ec:,.0f}".replace(",", "X").replace(".", ",").replace("X", ".")]],
                                columns=resumo_estado_codepe.columns)

    resumo_estado_codepe = pd.concat([resumo_estado_codepe, total_row_ec], ignore_index=True)

    st.markdown("### 📒 Resumo por Estado conf. CODEPE")
    st.markdown(resumo_estado_codepe.to_html(index=False, escape=False, classes='custom-table'), unsafe_allow_html=True)

    st.write(primeiro_dia_mes_anterior)
