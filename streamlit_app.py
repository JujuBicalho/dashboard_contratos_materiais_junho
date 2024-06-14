import streamlit as st
import streamlit.components.v1 as components
from threading import Thread
import os

# Função para rodar o servidor Dash
def run_dash():
    os.system("python dash_contratos_materiais_app.py")

# Executar o servidor Dash em uma thread separada
thread = Thread(target=run_dash)
thread.daemon = True

thread.start()

# Esperar um pouco para o servidor iniciar (ajuste conforme necessário)
import time
time.sleep(5)

# Mostrar o Dash no Streamlit
components.iframe("http://127.0.0.1:8052", width=1200, height=800, scrolling=True)


