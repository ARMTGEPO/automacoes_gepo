import streamlit as st


# Página de Auditoria Acumulada
st.title("📝 Auditoria de Produção Acumulada")

username = st.session_state.get('username', 'usuário')

st.markdown(f"Seja bem vindo {username}, faça o upload do relatório de produção acumulada para iniciar a auditoria.")

with st.expander("📖 Análises Realizadas", expanded=True):

    st.write("""
                Em desenvolvimento.
             """)
