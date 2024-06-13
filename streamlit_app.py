import streamlit as st
import dash
from dash import dcc, html
import dash_bootstrap_components as dbc
import pandas as pd
from datetime import datetime, timedelta
import plotly.express as px
from dash.dependencies import Input, Output
from threading import Thread

# Função para rodar o servidor Dash
def run_dash():
    app.run_server(debug=False, use_reloader=False, dev_tools_props_check=False)

# Carregar o arquivo Excel
file_path = 'BASE BI CONTRATOS.xlsx'
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

# Criar DataFrame para o gráfico
data_consumo = pd.DataFrame({
    'Consumo': ["Abaixo de 60%", "Entre 60% e 80%", "Acima de 80%"],
    'Quantidade': [contratos_abaixo_60, contratos_acima_60, contratos_acima_80]
})

# Gráfico de barras para mostrar os consumos
fig_consumo = px.bar(
    data_consumo,
    x='Consumo',
    y='Quantidade',
    title='Consumo',
    color='Consumo',
    color_discrete_map={
        "Abaixo de 60%": "green",
        "Entre 60% e 80%": "orange",
        "Acima de 80%": "red"
    }
)

fig_consumo.update_traces(
    texttemplate='%{y}', 
    textposition='inside', 
    insidetextanchor='middle', 
    textfont=dict(color='white', size=15, family='Arial', weight='bold'),
    hovertemplate='<b>Consumo</b>: %{x}<br><b>Quantidade</b>: %{y}<extra></extra>'  
) 
   
fig_consumo.update_layout(
    showlegend=False, 
    xaxis_title=None, 
    yaxis_title=None, 
    plot_bgcolor='white', 
    paper_bgcolor='white', 
    title_font=dict(size=30, family='Arial', color='#005a8d', weight='bold'),
    title_x=0.5,
    xaxis=dict(tickfont=dict(size=15, family='Arial', color='black', weight='bold'))
)

fig_consumo.update_xaxes(
    tickfont=dict(size=13, family='Arial', color='black', weight='bold')
)

# Layout do aplicativo Dash
app = dash.Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP])

app.layout = dbc.Container([
    html.Div([
        html.H1("Análise Descritiva Contratos de Materiais - Junho 2024", style={
            'textAlign': 'center', 
            'color': '#005a8d', 
            'backgroundColor': '#F0F8FF', 
            'padding': '20px', 
            'border-radius': '10px',
            'width': '100%',
            'fontSize': '24px'  # Ajuste o tamanho da fonte conforme necessário
        }),
    ], style={'marginBottom': '40px', 'width': '100%'}),
    
    dbc.Row([
        dbc.Col([
            dbc.Card([
                dbc.CardBody([
                    html.H5("Total de Contratos", className="card-title", style={'textAlign': 'center'}),
                    html.H2(f"{analise_descritiva['Total de Contratos']}", className="card-text", style={'textAlign': 'center'}),
                ], style={'textAlign': 'center'}),
            ], color="info", inverse=True, style={'border-radius': '50px', 'height': '100%'}),
        ], width=3),
        
        dbc.Col([
            dbc.Card([
                dbc.CardBody([
                    html.H5(["Prox. Vencimento (6 meses)"], className="card-title", style={'textAlign': 'center', 'color': 'red'}),
                    html.H2(f"{analise_descritiva['Contratos Prox. Vencimento']}", className="card-text", style={'textAlign': 'center', 'color': 'red'}),
                ], style={'textAlign': 'center'}),
            ], color="info", inverse=True, style={'border-radius': '15px'}),
        ], width=3),      
        
        
        dbc.Col([
            dbc.Card([
                dbc.CardBody([
                    html.H5("Valor dos Contratos (Bi)", className="card-title", style={'textAlign': 'center'}),
                    html.H2(f"{analise_descritiva['Valor Total dos Contratos (Bi)']:.3f}", className="card-text", style={'textAlign': 'center'}),
                ], style={'textAlign': 'center'}),
            ], color="info", inverse=True, style={'border-radius': '15px'}),
        ], width=3),
    ], justify='center', className="mb-4"),
    
    dbc.Row([
        dbc.Col([
            dbc.Card([
                dbc.CardBody([
                    html.H5("Valor Global Pendente (Bi)", className="card-title", style={'textAlign': 'center'}),
                    html.H2(f"{analise_descritiva['Valor Global Pendente (Bi)']:.3f}", className="card-text", style={'textAlign': 'center'}),
                ], style={'textAlign': 'center'}),
            ], color="info", inverse=True, style={'border-radius': '15px', 'height': '100%'}),
        ], width=3),
        
        dbc.Col([
            dbc.Card([
                dbc.CardBody([
                    html.H5("Com Consumo Mínimo", className="card-title", style={'textAlign': 'center'}),
                    html.H2(f"{analise_descritiva['Contratos com Consumo Mínimo']}", className="card-text", style={'textAlign': 'center'}),
                ], style={'textAlign': 'center'}),
            ], color="info", inverse=True, style={'border-radius': '15px'}),
        ], width=3),
        
        dbc.Col([
            dbc.Card([
                dbc.CardBody([
                    html.H5("Consumo Mínimo Atingido", className="card-title", style={'textAlign': 'center'}),
                    html.H2(f"{analise_descritiva['Consumo Mínimo Atingido']}", className="card-text", style={'textAlign': 'center'}),
                ], style={'textAlign': 'center'}),
            ], color="info", inverse=True, style={'border-radius': '15px'}),
        ], width=3),
        
        dbc.Col([
            dbc.Card([
                dbc.CardBody([
                    html.H5("Materiais Sem Contrato", className="card-title", style={'textAlign': 'center'}),
                    html.H2(f"{analise_descritiva['Materiais Sem Contrato']}", className="card-text", style={'textAlign': 'center'}),
                ], style={'textAlign': 'center'}),
            ], color="info", inverse=True, style={'border-radius': '15px'}),
        ], width=3),
    ], justify='center', className="mb-4"),
    
    dbc.Row([
        dbc.Col([
            dcc.Graph(figure=fig_consumo)
        ], width=12)
    ])
], fluid=True, style={'backgroundColor': 'white', 'width': '100%'})

if __name__ == '__main__':
    # Rodar o Dash em uma thread separada
    thread = Thread(target=run_dash)
    thread.daemon = True
    thread.start()

    # Mostrar o Dash no Streamlit
    st.markdown(
        """
        <style>
        .iframe-container {
            width: 100%;
            height: 100vh;
        }
        .iframe {
            width: 100%;
            height: 100%;
            border: none;
        }
        </style>
        <div class="iframe-container">
            <iframe src='http://127.0.0.1:8050' class="iframe"></iframe>
        </div>
        """,
        unsafe_allow_html=True
    )
