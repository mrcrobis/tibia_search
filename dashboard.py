import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(layout="wide")

# Caminho do arquivo Excel
x = r'D:\Programming\tibia_search\db_sheets\11-03-2025\11-03-2025_combined.xlsx'

# Carregar a planilha específica
df = pd.read_excel(x, sheet_name='DB')

# Arredondar valores das colunas
df[['$ Sell Offer', '$ Buy Offer']] = df[['$ Sell Offer', '$ Buy Offer']].round(2)

# Exibir o DataFrame na interface do Streamlit
st.dataframe(df)
