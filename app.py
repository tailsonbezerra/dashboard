import pandas as pd
import dash
from dash import dcc, html
import plotly.express as px
import plotly.graph_objs as go
import plotly.io as pio
from flask import Flask

# Definir o tema padrão para plotly
pio.templates.default = "plotly_dark"

# Carregar o arquivo Excel
file_path = 'C:\\Users\\User\\Desktop\\projetop\\estagioo.xlsx'
df = pd.read_excel(file_path)

# Calcular as médias
media_idades = df['IDADE'].mean()
media_tempo_meses = df['Tempo em meses'].mean()
media_vencimento_total = df['Total de vencimentos'].mean()

# Normalizar os valores da coluna 'GÊNERO'
df['GÊNERO'] = df['GÊNERO'].str.strip().str.lower().replace({
    'feminino': 'Feminino',
    'masculino': 'Masculino'
}).fillna('vazio')

# Contar a quantidade de cada gênero
gender_counts = df['GÊNERO'].value_counts().reset_index()
gender_counts.columns = ['Gênero', 'Quantidade']

# Criar uma figura de pizza
fig_pie = px.pie(gender_counts, values='Quantidade', names='Gênero', title='Distribuição por Gênero')
fig_pie.update_traces(marker=dict(colors=['#ff69b4', '#1e90ff']))  # rosa para feminino, azul para masculino
fig_pie.update_layout(title_x=0.5, font=dict(size=18, family='Arial, sans-serif', color='white'))

# Definir as faixas etárias
bins = [0, 21, 25, 29, 34, 100]
labels = ['-18 até 21', '22 até 25', '26 até 29', '30 até 34', '35 ou +']
df['Faixa Etária'] = pd.cut(df['IDADE'], bins=bins, labels=labels, right=False)

# Contar a quantidade de pessoas em cada faixa etária e gênero
age_gender_counts = df[df['GÊNERO'] != 'vazio'].groupby(['Faixa Etária', 'GÊNERO'], observed=False).size().unstack().fillna(0)

# Criar a pirâmide etária
trace1 = go.Bar(
    y=age_gender_counts.index,
    x=-age_gender_counts['Masculino'],
    name='Masculino',
    orientation='h',
    marker=dict(color='#1e90ff')
)
trace2 = go.Bar(
    y=age_gender_counts.index,
    x=age_gender_counts['Feminino'],
    name='Feminino',
    orientation='h',
    marker=dict(color='#ff69b4')
)

data_pyramid = [trace1, trace2]

layout_pyramid = go.Layout(
    title='Distribuição Etária por Gênero',
    title_x=0.5,
    barmode='overlay',
    bargap=0.1,
    xaxis=dict(title='Quantidade'),
    yaxis=dict(title='Faixa Etária', categoryorder='category ascending'),
    template='plotly_dark',
    font=dict(size=18, family='Arial, sans-serif', color='white')
)

fig_pyramid = go.Figure(data=data_pyramid, layout=layout_pyramid)

# Lista de cidades da região metropolitana de Recife
rmr_cities = [
    'recife', 'olinda', 'jaboatão dos guararapes', 'paulista', 'cabo de santo agostinho',
    'camaragibe', 'igarassu', 'abreu e lima', 'ipojuca', 'itapissuma', 'moreno', 'araçoiaba',
    'itamaracá', 'são lourenço da mata'
]

# Normalizar a coluna 'CIDADE'
df['CIDADE'] = df['CIDADE'].str.strip().str.lower()

# Categorizar as cidades corretamente
def categorize_city(city):
    if pd.isna(city) or city == '':
        return 'vazio'
    elif city == 'recife':
        return 'Recife'
    elif city in rmr_cities:
        return 'Dentro da RMR'
    else:
        return 'Fora da RMR'

df['Região Metropolitana'] = df['CIDADE'].apply(categorize_city)

# Contar a quantidade de pessoas em cada categoria
rmr_counts = df['Região Metropolitana'].value_counts().reset_index()
rmr_counts.columns = ['Categoria', 'Quantidade']

# Criar uma figura de barras para a distribuição das cidades
fig_rmr = px.bar(rmr_counts, x='Categoria', y='Quantidade', title='Distribuição por Região Metropolitana', color='Categoria')
fig_rmr.update_layout(title_x=0.5, font=dict(size=18, family='Arial, sans-serif', color='white'))

# Analisar a coluna de iniciativa
df['INICIATIVA'] = df['INICIATIVA'].str.strip().str.lower().fillna('vazio')
initiative_counts = df['INICIATIVA'].value_counts().reset_index()
initiative_counts.columns = ['Iniciativa', 'Quantidade']

# Criar uma figura de pizza para a distribuição de iniciativas
fig_initiative_pie = go.Figure(data=[go.Pie(
    labels=initiative_counts['Iniciativa'], 
    values=initiative_counts['Quantidade'],
    hole=.4,
    textinfo='percent',
    textfont=dict(size=16, color='white'),
    marker=dict(colors=px.colors.qualitative.Plotly)
)])
fig_initiative_pie.update_layout(
    title='Distribuição por Iniciativa',
    title_x=0.5,
    font=dict(size=18, family='Arial, sans-serif', color='white'),
    annotations=[dict(text='Iniciativas', x=0.5, y=0.5, font_size=20, showarrow=False)]
)

# Identificar o nome exato da coluna 'TIPO\nCONVÊNIO'
convenio_column_name = [col for col in df.columns if 'CONVÊNIO' in col][0]

# Analisar a coluna de convênios
df['TIPO_CONVÊNIO'] = df[convenio_column_name].str.strip().str.lower().replace({
    'agente de integração': 'agente de integração',
    'concedente': 'concedente',
    'unidade da ufpe': 'unidade da ufpe'
}).fillna('vazio')
convenio_counts = df['TIPO_CONVÊNIO'].value_counts().reset_index()
convenio_counts.columns = ['Tipo de Convênio', 'Quantidade']

# Criar uma figura de pizza para a distribuição de tipos de convênios
fig_convenio_pie = go.Figure(data=[go.Pie(
    labels=convenio_counts['Tipo de Convênio'], 
    values=convenio_counts['Quantidade'],
    pull=[0.1, 0, 0],
    textinfo='percent+label',
    textfont=dict(size=16, color='white'),
    marker=dict(colors=px.colors.qualitative.Set2)
)])
fig_convenio_pie.update_layout(
    title='Distribuição por Tipo de Convênio',
    title_x=0.5,
    font=dict(size=18, family='Arial, sans-serif', color='white')
)

# Identificar o nome exato da coluna 'AGENTE\nINTEGRAÇÃO'
agente_column_name = [col for col in df.columns if 'INTEGRAÇÃO' in col][0]

# Analisar a coluna de agentes de integração
df['AGENTE_INTEGRAÇÃO'] = df[agente_column_name].str.strip().str.lower().fillna('vazio')
top_5_agentes = df[df['AGENTE_INTEGRAÇÃO'] != 'vazio']['AGENTE_INTEGRAÇÃO'].value_counts().nlargest(5).reset_index()
top_5_agentes.columns = ['Agente de Integração', 'Quantidade']

# Criar uma figura de barras para os top 5 agentes de integração
fig_top_5_agentes = px.bar(top_5_agentes, x='Agente de Integração', y='Quantidade', title='Top 5 Agentes de Integração')
fig_top_5_agentes.update_layout(title_x=0.5, font=dict(size=18, family='Arial, sans-serif', color='white'))

# Identificar o nome exato da coluna 'CARGA\nHORÁRIA'
carga_horaria_column_name = [col for col in df.columns if 'CARGA' in col][0]

# Analisar a coluna de carga horária
df['CARGA_HORÁRIA'] = df[carga_horaria_column_name].str.lower().str.strip()

# Criar um ranking dos tipos de 'CARGA_HORÁRIA'
carga_horaria_ranking = df['CARGA_HORÁRIA'].value_counts().reset_index()
carga_horaria_ranking.columns = ['CARGA_HORARIA', 'COUNT']


df['CARGA_HORARIA_NUM'] = df['CARGA_HORÁRIA'].str.extract('(\\d+)').astype(float)


average_carga_horaria = df['CARGA_HORARIA_NUM'].mean()


fig_carga_horaria = px.bar(carga_horaria_ranking, x='CARGA_HORARIA', y='COUNT',
labels={'CARGA_HORARIA': 'Carga Horária', 'COUNT': 'Contagem'},
title='Distribuição de Carga Horária')
fig_carga_horaria.update_layout(title_x=0.5, font=dict(size=18, family='Arial, sans-serif', color='white'))


server = Flask(__name__)


external_stylesheets = ['https://codepen.io/chriddyp/pen/bWLwgP.css']
app = dash.Dash(__name__, server=server, external_stylesheets=external_stylesheets)


app.layout = html.Div(className='dashboard-container', children=[
html.H1(children='Dashboard de Visualização da Tabela de Estágios Não Obrigatórios', style={'text-align': 'center', 'color': '#FFFFFF', 'font-family': 'Arial, sans-serif', 'font-size': '32px'}),
html.Div(children=[
html.Div(children=[
dcc.Graph(id='gender-pie-chart', figure=fig_pie, className='graph-container')
], style={'display': 'inline-block', 'width': '48%', 'padding': '10px', 'background-color': 'rgba(0, 0, 0, 0.3)', 'border-radius': '10px'}),
html.Div(children=[
dcc.Graph(id='age-gender-pyramid', figure=fig_pyramid, className='graph-container')
], style={'display': 'inline-block', 'width': '48%', 'padding': '10px', 'background-color': 'rgba(0, 0, 0, 0.3)', 'border-radius': '10px'}),
]),
html.Div(children=[
dcc.Graph(id='rmr-bar-chart', figure=fig_rmr, className='graph-container')
], style={'padding': '10px', 'background-color': 'rgba(0, 0, 0, 0.3)', 'border-radius': '10px'}),
html.Div(children=[
html.Div(children=[
dcc.Graph(id='initiative-pie-chart', figure=fig_initiative_pie, className='graph-container')
], style={'display': 'inline-block', 'width': '48%', 'padding': '10px', 'background-color': 'rgba(0, 0, 0, 0.3)', 'border-radius': '10px'}),
html.Div(children=[
dcc.Graph(id='convenio-pie-chart', figure=fig_convenio_pie, className='graph-container')
], style={'display': 'inline-block', 'width': '48%', 'padding': '10px', 'background-color': 'rgba(0, 0, 0, 0.3)', 'border-radius': '10px'}),
]),
html.Div(children=[
html.Div(children=[
dcc.Graph(id='top-5-agentes-bar-chart', figure=fig_top_5_agentes, className='graph-container')
], style={'display': 'inline-block', 'width': '48%', 'padding': '10px', 'background-color': 'rgba(0, 0, 0, 0.3)', 'border-radius': '10px'}),
html.Div(children=[
dcc.Graph(id='carga-horaria-ranking', figure=fig_carga_horaria, className='graph-container')
], style={'display': 'inline-block', 'width': '48%', 'padding': '10px', 'background-color': 'rgba(0, 0, 0, 0.3)', 'border-radius': '10px'}),
]),
html.Div(children=[
html.Div(children=[
html.P(children='Média das Idades:', style={'font-weight': 'bold', 'text-align': 'center', 'color': '#FFFFFF', 'font-size': '20px'}),
html.P(children=f'{media_idades:.2f} anos', style={'font-size': '24px', 'color': '#4CAF50', 'text-align': 'center'})
], className='stat-box', style={'width': '30%', 'display': 'inline-block', 'padding': '20px', 'background-color': 'rgba(0, 0, 0, 0.3)', 'border-radius': '10px', 'margin': '10px'}),
html.Div(children=[
html.P(children='Média do Tempo em Meses:', style={'font-weight': 'bold', 'text-align': 'center', 'color': '#FFFFFF', 'font-size': '20px'}),
html.P(children=f'{media_tempo_meses:.2f} meses', style={'font-size': '24px', 'color': '#2196F3', 'text-align': 'center'})
], className='stat-box', style={'width': '30%', 'display': 'inline-block', 'padding': '20px', 'background-color': 'rgba(0, 0, 0, 0.3)', 'border-radius': '10px', 'margin': '10px'}),
html.Div(children=[
html.P(children='Média do Vencimento Total:', style={'font-weight': 'bold', 'text-align': 'center', 'color': '#FFFFFF', 'font-size': '20px'}),
html.P(children=f'R$ {media_vencimento_total:.2f}', style={'font-size': '24px', 'color': '#FF5722', 'text-align': 'center'})
], className='stat-box', style={'width': '30%', 'display': 'inline-block', 'padding': '20px', 'background-color': 'rgba(0, 0, 0, 0.3)', 'border-radius': '10px', 'margin': '10px'}),
], style={'display': 'flex', 'justify-content': 'space-around', 'width': '100%', 'padding': '20px', 'background-color': '#2C3E50', 'border': '1px solid #34495E', 'border-radius': '10px', 'box-shadow': '0 4px 8px 0 rgba(0, 0, 0, 0.2)'})
], style={'background-color': '#2C3E50', 'padding': '20px', 'border-radius': '10px', 'box-shadow': '0 8px 16px 0 rgba(0, 0, 0, 0.3)'}),


if __name__ == '__main__':
    app.run_server(debug=True)