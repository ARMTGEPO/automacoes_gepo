import streamlit as st
import pandas as pd
from fpdf import FPDF, XPos, YPos
from datetime import datetime
import tempfile

st.set_page_config(page_title="Validador Produção", layout="centered", page_icon=":bar_chart")

# === Função de validação e geração do relatório ===
def validar_e_gerar_relatorio(df, primeiro_dia_mes_anterior, output_path="relatorio_inconsistencias.pdf"):
    inconsistencias = {}
    estados_efetivados = ["1 - Aprovado", "2 - Reprovado", "3 - Evadido"]

    # Regra 1
    mask1 = (df["Estado conf. CODEPE"] == "4 - Desistente") & (df["CH_Apurada_Mes"] > 0)
    inconsistencias['Desistentes com CH Apurada'] = df.loc[mask1, [
        "Matrícula", "Turma", "CPF do Aluno", "Estado conf. CODEPE", "CH_Apurada_Mes"
    ]]

    # Regra 2
    mask2 = (df["Termino da Execução da Turma"] < primeiro_dia_mes_anterior) & (df["CH_Apurada_Mes"] != 0.00)
    inconsistencias['Turmas Encerradas em Exercício Anterior contabilizando CH'] = df.loc[mask2, [
        "Unidade Operativa da Turma", "Matrícula", "Turma", "Nome do Aluno", 
        "Estado da Turma", "Estado conf. CODEPE", "CH_Apurada_Mes", "Termino da Execução da Turma"
    ]]

    # Regra 3
    df_filtrado = df[df["Estado da Matrícula do Aluno"] != "Matrícula Cancelada"]
    mask3 = df_filtrado.groupby(["CPF do Aluno", "Turma"])["Matrícula"].transform('nunique') > 1
    inconsistencias['CPF com Múltiplas Matrículas na mesma turma'] = df_filtrado.loc[mask3, [
        "Unidade Operativa da Turma", "Matrícula", "Turma", "CPF do Aluno", "Estado conf. CODEPE", 
        "Estado da Matrícula do Aluno", "Recurso Financeiro", "CH_Apurada_Mes"
    ]].sort_values(["CPF do Aluno", "Turma"])

    # Regra 4
    mask_psg = (df["Recurso Financeiro"] == "PSG") & (df["Estado conf. CODEPE"] == "6 - Em Processo")
    cpf_counts = df.loc[mask_psg, "CPF do Aluno"].value_counts().reset_index()
    cpf_counts.columns = ["CPF do Aluno", "count"]
    cpf_invalidos = cpf_counts[cpf_counts["count"] > 2]["CPF do Aluno"]
    mask4 = df["CPF do Aluno"].isin(cpf_invalidos) & mask_psg
    inconsistencias['CPF com > 2 Turmas PSG'] = df.loc[mask4, [
        "Matrícula", "Turma", "CPF do Aluno", "Recurso Financeiro", 
        "Estado conf. CODEPE", "CH_Apurada_Mes"
    ]].sort_values("CPF do Aluno")

    # Regra 5
    mask5 = df["Estado conf. CODEPE"].isin(estados_efetivados) & df["Estado da Turma"].str.contains("Em Processo")
    inconsistencias['Efetivados em Turmas Ativas'] = df.loc[mask5, [
        "Unidade Operativa da Turma", "Matrícula", "Turma", "CPF do Aluno", "Estado conf. CODEPE", 
        "Estado da Turma", "CH_Apurada_Mes"
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
    pdf.output(output_path)

# === Interface Principal ===
st.title("📝 Validação de Produção Mensal")
username = st.session_state.get("username", "usuário")
st.markdown(f"Olá, **{username}**! Faça upload do Relatório Analítico Mensal de Produção.")

competencia_input = st.text_input("🗓️ Informe a competência (MM/AAAA)", placeholder="Ex: 05/2025")
uploaded_file = st.file_uploader("📂 Selecione o arquivo CSV", type=["csv"])

if competencia_input and uploaded_file:
    try:
        competencia_dt = datetime.strptime(competencia_input, "%m/%Y")
        primeiro_dia_mes_anterior = competencia_dt.replace(day=1)

        df = pd.read_csv(uploaded_file, encoding='latin-1', sep=';')
        df["CH_Apurada_Mes"] = df["CH_Apurada_Mes"].astype(str).str.replace(',', '.').astype(float)
        df["Termino da Execução da Turma"] = pd.to_datetime(df["Termino da Execução da Turma"], dayfirst=True, errors='coerce')

        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
            validar_e_gerar_relatorio(df, primeiro_dia_mes_anterior, tmp.name)
            with open(tmp.name, "rb") as f:
                st.success("✅ Relatório gerado com sucesso!")
                st.download_button("📄 Baixar PDF", data=f, file_name="relatorio_inconsistencias.pdf", mime="application/pdf")

    except ValueError:
        st.error("❌ Competência inválida. Use o formato MM/AAAA.")
    except Exception as e:
        st.error(f"❌ Erro ao processar o arquivo: {e}")
elif uploaded_file and not competencia_input:
    st.warning("⚠️ Informe a competência antes de processar o arquivo.")
