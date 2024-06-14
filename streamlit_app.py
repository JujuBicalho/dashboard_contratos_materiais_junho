import streamlit as st
import streamlit.components.v1 as components
from threading import Thread
import os
import time

# Função para rodar o servidor Dash
def run_dash():
    os.system("python dash_contratos_materiais_app.py")

# Executar o servidor Dash em uma thread separada
thread = Thread(target=run_dash)
thread.daemon = True

thread.start()

# Esperar um pouco para o servidor iniciar 
import time
time.sleep(5)

# Título da página do Streamlit
st.title("Dashboard Contrato de Materiais Resumido")

# Mostrar o Dash no Streamlit
components.iframe("http://127.0.0.1:8055", width=800, height=600, scrolling=False)



