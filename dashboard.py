import streamlit as st
import pandas as pd
import plotly.express as px
from dictionary import *

#st.set_page_config(layout="wide")

# Caminho do arquivo Excel
x = r'D:\Programming\tibia_search\db_sheets\11-03-2025\11-03-2025_combined.xlsx'

# Caminho pasta excel (usar para criar uma função para visualização dos dados em diferentes linhas de tempo)
db_sheets = r'D:\Programming\tibia_search\db_sheets'

# Carregar a planilha específica, os valores serão todos str
df = pd.read_excel(x, sheet_name='DB', dtype=str)

# Troca os valores "-" para 0, para evitar mensagem de erro no terminal
df.replace("-", "0", inplace=True)

# Converter as colunas para número (caso tenham sido lidas como string)
df['$ Sell Offer'] = pd.to_numeric(df['$ Sell Offer'], errors='coerce')
df['$ Buy Offer'] = pd.to_numeric(df['$ Buy Offer'], errors='coerce')

# Arredondar valores das colunas
df[['$ Sell Offer', '$ Buy Offer']] = df[['$ Sell Offer', '$ Buy Offer']].round(2)

# Filtro
items = df['Item'].unique()
servidores_index = df['Servidor'].unique()

item_select = st.selectbox("Selecione o item:", items)

servidor_select = st.multiselect("Selecione os servidores:", servidores_index, default=servidores_index[0:85])

# atualiza a planilha após os filtros
cond = (df['Item'] == item_select) & (df['Servidor'].isin(servidor_select))

df = df[cond]

# Exibir o DataFrame na interface do Streamlit
st.dataframe(df)
