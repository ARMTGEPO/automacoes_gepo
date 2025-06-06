import streamlit as st

# Login
def login():
    st.markdown("<h2 style='text-align: center;'>🔐 Login</h2>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center;'>Acesse o sistema com suas credenciais</p>", unsafe_allow_html=True)

    with st.form("login_form", clear_on_submit=False):
        username = st.text_input("👤 Usuário")
        password = st.text_input("🔑 Senha", type="password")
        submit = st.form_submit_button("Entrar")

        if submit:
            if username in st.secrets.auth.users:
                idx = st.secrets.auth.users.index(username)
                if st.secrets.auth.passwords[idx] == password:
                    st.session_state["authenticated"] = True
                    st.session_state["username"] = username
                    st.success(f"Bem-vindo, {username}!")
                    st.rerun()
                else:
                    st.error("Senha incorreta.")
            else:
                st.error("Usuário não encontrado.")

# Verifica autenticação
if "authenticated" not in st.session_state or not st.session_state["authenticated"]:
    login()
    st.stop()

# Menu com páginas
validador_mensal = st.Page("validador_mensal.py", title="Validador Produção Mensal", icon=":material/add_circle:")
validador_acumulado = st.Page("validador_acumulado.py", title="Validador Produção Acumulado", icon=":material/add_circle:")
pg = st.navigation([validador_mensal, validador_acumulado])

pg.run()
