import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import plotly.express as px

# Carregar o arquivo Excel
file_path = 'BASE BI CONTRATOS.xlsx'  # Altere para o caminho correto do seu arquivo
xls = pd.ExcelFile(file_path)

# Carregar as abas relevantes
analise_df = pd.read_excel(xls, sheet_name='ANÁLISE', header=1)  # Definindo a segunda linha como cabeçalho
contratos_df = pd.read_excel(xls, sheet_name='Contratos', header=0)  # Definindo a primeira linha como cabeçalho
demanda_spt_df = pd.read_excel(xls, sheet_name='Demanda SPT')

# Remover espaços em branco dos nomes das colunas
analise_df.columns = [str(col).strip() for col in analise_df.columns]
contratos_df.columns = [str(col).strip() for col in contratos_df.columns]

# Filtrar e calcular métricas para o mês de junho
total_contratos = len(analise_df['Doc.compra'].unique())
valor_total_contratos = analise_df['Val.fixado'].astype(float).sum()
valor_global_pendente = analise_df['ValGlPend.'].astype(float).sum()

# Calcular valor consumido do contrato
contratos_df['valor_consumido_contrato'] = contratos_df['Val.fixado'].astype(float) - contratos_df['ValGlPend.'].astype(float)

# Converter a coluna de data 'FimValid/' para datetime
contratos_df['FimValid/'] = pd.to_datetime(contratos_df['FimValid/'], format='%d/%m/%Y')

# Definir o intervalo de 6 meses a partir da data atual
data_atual = datetime.now()
data_limite = data_atual + timedelta(days=180)

# Filtrar contratos próximos ao vencimento de 6 meses
contratos_prox_venc = analise_df[
    analise_df['FimValid/'] <= data_limite
]

# Filtrar contratos com consumo mínimo definido
contratos_com_minimo = contratos_df[
    (contratos_df['Consumo Mínimo'].str.lower() == 'sim') | (contratos_df['Consumo Mínimo'] == 1)
]
total_contratos_com_minimo = contratos_com_minimo['Doc.compra'].nunique()

# Aplicar a lógica da fórmula DAX em Python diretamente nos contratos com consumo mínimo
contratos_df['Consumo Mínimo Atingido'] = contratos_df.apply(
    lambda row: (
        "Consumo mínimo atingido" if pd.notna(row['Valor Consumo Mínimo']) and float(row['valor_consumido_contrato']) >= float(row['Valor Consumo Mínimo'])
        else "Consumo mínimo não atingido" if pd.notna(row['Valor Consumo Mínimo'])
        else "Não tem valor mínimo"
    ), axis=1
)

# Filtrar contratos que atingiram o consumo mínimo
contratos_minimo_atingido = contratos_df[
    (contratos_df['Consumo Mínimo Atingido'] == 'Consumo mínimo atingido') &
    (~contratos_df['Doc.compra'].isin(['JA10063222', 'JA10114401']))
]
total_contratos_minimo_atingido = contratos_minimo_atingido['Doc.compra'].nunique()

# Materiais sem contrato
materiais_sem_contrato = demanda_spt_df[demanda_spt_df['Contrato Vigente'] == "Não"].shape[0]

# Filtrar os contratos por Farol SALDO
contratos_abaixo_60 = analise_df[analise_df['Farol SALDO'].astype(float) < 0.6]['Doc.compra'].nunique()
contratos_acima_60 = analise_df[(analise_df['Farol SALDO'].astype(float) >= 0.6) & (analise_df['Farol SALDO'].astype(float) < 0.8)]['Doc.compra'].nunique()
contratos_acima_80 = analise_df[analise_df['Farol SALDO'].astype(float) >= 0.8]['Doc.compra'].nunique()

# Análise descritiva de junho
analise_descritiva = {
    "Contratos Prox. Vencimento": contratos_prox_venc.shape[0],
    "Consumo Abaixo de 60%": contratos_abaixo_60,
    "Consumo Acima de 60%": contratos_acima_60,
    "Consumo Acima de 80%": contratos_acima_80,
    "Total de Contratos": total_contratos,
    "Valor Total dos Contratos (Bi)": valor_total_contratos / 1e9,
    "Valor Global Pendente (Bi)": valor_global_pendente / 1e9,
    "Contratos com Consumo Mínimo": total_contratos_com_minimo,
    "Consumo Mínimo Atingido": total_contratos_minimo_atingido,
    "Materiais Sem Contrato": materiais_sem_contrato
}

# Gráfico de barras para mostrar os consumos
fig_consumo = px.bar(
    x=["Abaixo de 60%", "Entre 60% e 80%", "Acima de 80%"],
    y=[contratos_abaixo_60, contratos_acima_60, contratos_acima_80],
    title='Consumo',
    labels={'x': 'Consumo', 'y': 'Quantidade'},
    color=["Consumo Abaixo de 60%", "Consumo Acima de 60%", "Consumo Acima de 80%"],
    color_discrete_map={
        "Consumo Abaixo de 60%": "green",
        "Consumo Acima de 60%": "orange",
        "Consumo Acima de 80%": "red"
    }
)
fig_consumo.update_traces(texttemplate='%{y}', textposition='inside', insidetextanchor='middle', textfont=dict(color='white', size=25, family='Arial', weight='bold'))
fig_consumo.update_layout(showlegend=False, xaxis_title=None, yaxis_title=None, plot_bgcolor='#DCDCDC', paper_bgcolor='white', title_font=dict(size=30, family='Arial', color='#005a8d', weight='bold'), title_x=0.5)
fig_consumo.update_layout(xaxis=dict(tickfont=dict(size=14, weight='bold')))

# Layout do aplicativo Streamlit
st.set_page_config(layout="wide")

st.markdown("""
    <style>
    .stApp {
        background-color: white;
    }
    .big-font {
        font-size:30px !important;
    }
    .card-title {
        font-size: 20px;
    }
    .card-text {
        font-size: 28px;
    }
    </style>
    """, unsafe_allow_html=True)

st.title('Análise Descritiva de Junho - Contratos Materiais')

st.markdown("""
    <style>
    .stMarkdown h1 {
        background-color: #f0f8ff;
        padding: 20px;
        border-radius: 10px;
        text-align: center;
        color: #005a8d;
    }
    </style>
    """, unsafe_allow_html=True)

# Cartões
col1, col2, col3 = st.columns(3)
col4, col5, col6 = st.columns(3)

with col1:
    st.markdown(f"<div class='card-title'>Total de Contratos</div><div class='card-text'>{analise_descritiva['Total de Contratos']}</div>", unsafe_allow_html=True)

with col2:
    st.markdown(f"<div class='card-title'>Prox. Vencimento<br>(6 meses)</div><div class='card-text' style='color:red;'>{analise_descritiva['Contratos Prox. Vencimento']}</div>", unsafe_allow_html=True)

with col3:
    st.markdown(f"<div class='card-title'>Valor Total dos Contratos (Bi)</div><div class='card-text'>{analise_descritiva['Valor Total dos Contratos (Bi)']:.3f}</div>", unsafe_allow_html=True)

with col4:
    st.markdown(f"<div class='card-title'>Valor Global Pendente (Bi)</div><div class='card-text'>{analise_descritiva['Valor Global Pendente (Bi)']:.3f}</div>", unsafe_allow_html=True)

with col5:
    st.markdown(f"<div class='card-title'>Com Consumo Mínimo</div><div class='card-text'>{analise_descritiva['Contratos com Consumo Mínimo']}</div>", unsafe_allow_html=True)

with col6:
    st.markdown(f"<div class='card-title'>Consumo Mínimo Atingido</div><div class='card-text'>{analise_descritiva['Consumo Mínimo Atingido']}</div>", unsafe_allow_html=True)

# Cartão extra para "Materiais Sem Contrato"
st.markdown(f"<div class='card-title'>Materiais Sem Contrato</div><div class='card-text'>{analise_descritiva['Materiais Sem Contrato']}</div>", unsafe_allow_html=True)

# Gráfico
st.plotly_chart(fig_consumo, use_container_width=True)
