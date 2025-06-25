import streamlit as st
from streamlit import dialog

# Diálogo de validação (exibe alertas)
@st.dialog('Atenção')
def login_validation(username, password):
    if username == '' or password == '':
        st.error('Por favor, preencha usuário e senha.')
    else:
        st.error('Credenciais inválidas.')

# Login
def login():
    st.markdown("<h2 style='text-align: center;'>🔐 Login</h2>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center;'>Acesse o sistema com suas credenciais</p>", unsafe_allow_html=True)

    # Layout centralizado com largura um pouco maior
    col1, col2, col3 = st.columns([1, 2.5, 1])
    with col2:
        with st.form("login_form", clear_on_submit=False):
            username = st.text_input("👤 Usuário", placeholder="Informe o usuário")
            password = st.text_input("🔑 Senha", type="password", placeholder="Informe a senha")
            submit = st.form_submit_button(
                label="Entrar",
                type="primary",
                use_container_width=True
            )


            if submit:
                if username in st.secrets.auth.users:
                    idx = st.secrets.auth.users.index(username)
                    if st.secrets.auth.passwords[idx] == password:
                        st.session_state["authenticated"] = True
                        st.session_state["username"] = username
                        st.success(f"✅ Bem-vindo, {username}!")
                        st.rerun()
                    else:
                        login_validation(username, password)
                else:
                    login_validation(username, password)

# Verifica autenticação
if "authenticated" not in st.session_state or not st.session_state["authenticated"]:
    login()
    st.stop()

# Menu com páginas
validador_mensal = st.Page("validador_mensal.py", title="Validador Produção Mensal", icon="📅")
validador_acumulado = st.Page("validador_acumulado.py", title="Validador Produção Acumulado", icon="📅")
pg = st.navigation([validador_mensal, validador_acumulado])

pg.run()
