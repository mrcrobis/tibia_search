import pytesseract
import openpyxl
import os
from datetime import datetime

pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

print(f"Informe o Servidor: ")
identificador_servidor = input()

# X = 'arquivo'.png
x = ''

# Obtém a data atual no formato dd-mm-yyyy
data_atual = datetime.now().strftime('%d-%m-%Y')

excel_path = r'D:\Programming\tibia_search\tibia_search_db_source.xlsx'
image_path = os.path.join(r'C:\Users\joaov\AppData\Local\Tibia\packages\Tibia\screenshots',x)
sheet_name = 'DB'
#df = pd.read_excel(excel_path, sheet_name='DB', engine='openpyxl')

# Carregar a planilha usando openpyxl
wb = openpyxl.load_workbook(excel_path)
ws = wb['DB']

# posicao do mouse para clicar na caixa depot dentro do tibia (depot acima do personagem), não funcionará caso tenha algum item em cima da caixa do depot
mouse1x = 1163
mouse1y = 313

# posicao do mouse para clicar na aba onde irá abrir o depot em si
mouse2x = 2522
mouse2y = 382

# clica no item pesquisado no depot para abrir a tela de preços
mouse3x = 859
mouse3y = 516


mouse4x = 865
mouse4y = 552

# Substitua as coordenadas pelos valores corretos para sua imagem
REGIAO_SELL_OFFER = (1340, 287, 1467, 302)  # Exemplo de coordenadas (left, upper, right, lower)
REGIAO_BUY_OFFER = (1340, 519, 1467, 532)  # Exemplo de coordenadas (left, upper, right, lower)

aux = 0
