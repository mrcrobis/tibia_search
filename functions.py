import pyautogui
import openpyxl
import time
import pytesseract
import os
import tkinter as tk
import time
import shutil
import re
from PIL import Image
from tkinter import messagebox
from global_var import identificador_servidor
from dictionary import items, items_grand_sanguine
from global_var import x, aux, aux_sanguine, excel_path, image_path, mouse2x, mouse2y, mouse3x, mouse3y, mouse4x, mouse4y, sheet_name, wb, ws, REGIAO_BUY_OFFER, REGIAO_SELL_OFFER, data_atual


def copiar_e_modificar_excel(arquivo_origem, identificador_servidor):

    nome_arquivo = "{}_{}.xlsx".format(data_atual, identificador_servidor)
    
    # Diretório onde o arquivo será salvo
    pasta_destino = "db_sheets/{}".format(data_atual)
    
    # Cria a pasta se ela não existir
    os.makedirs(pasta_destino, exist_ok=True)

    # Caminho completo do novo arquivo
    nome_arquivo = os.path.join(pasta_destino, "{}_{}.xlsx".format(data_atual, identificador_servidor))

    # Faz uma cópia do arquivo original para manter as fórmulas
    shutil.copy(arquivo_origem, nome_arquivo)

    # Abre o arquivo copiado para edição
    wb = openpyxl.load_workbook(nome_arquivo)

    # Mantém apenas a planilha "DB"
    if "DB" in wb.sheetnames:
        for sheet_name in wb.sheetnames[:]:  # Criamos uma cópia da lista para evitar erro ao remover
            if sheet_name != "DB":
                wb.remove(wb[sheet_name])
    else:
        print("A planilha 'DB' não foi encontrada. Nenhuma alteração feita.")
        return

    # Seleciona a planilha "DB"
    ws = wb["DB"]

    # Apaga os valores das colunas E e F da linha 2 até a linha 27319
    for row in range(2, 27319):  # Linha 2 até 287 (Excel usa indexação 1-based)
        for col in ["E", "F"]:  # Colunas E e F
            cell = ws[f"{col}{row}"]
            if not cell.data_type == "f":  # Mantém as fórmulas
                cell.value = None  # Apaga apenas valores estáticos

    # Salva o arquivo modificado
    wb.save(nome_arquivo)
    print(f"Arquivo '{nome_arquivo}' criado e modificado com sucesso!")

def combinar_planilhas(pasta_entrada, nome_base):
    # Lista todos os arquivos na pasta
    arquivos_xlsx = [f for f in os.listdir(pasta_entrada) if f.endswith(".xlsx")]

    # Filtra os arquivos que começam com o nome_base
    arquivos_para_unir = [os.path.join(pasta_entrada, f) for f in arquivos_xlsx if re.match(f"{nome_base}_.+\\.xlsx", f)]

    if not arquivos_para_unir:
        print(f"Nenhum arquivo encontrado para combinar com o nome base '{nome_base}'.")
        return

    # Carrega a primeira planilha como base
    wb_base = openpyxl.load_workbook(arquivos_para_unir[0])
    planilha_principal = wb_base[wb_base.sheetnames[0]]

    # Percorre os outros arquivos e preenche os valores vazios
    for arquivo in arquivos_para_unir[1:]:
        wb_atual = openpyxl.load_workbook(arquivo)
        planilha_atual = wb_atual[wb_atual.sheetnames[0]]

        # Percorre todas as células da planilha atual
        for row in planilha_atual.iter_rows():
            for cell in row:
                row_index = cell.row
                col_index = cell.column

                # Obtém a célula correspondente na planilha principal
                cell_principal = planilha_principal.cell(row=row_index, column=col_index)

                # Se a célula na planilha principal estiver vazia, preenche com o valor da planilha atual
                if cell_principal.value is None or cell_principal.value == "":
                    cell_principal.value = cell.value

    # Salva o arquivo modificado na mesma pasta
    nome_saida = os.path.join(pasta_entrada, f"{nome_base}_unificado.xlsx")
    wb_base.save(nome_saida)

    print(f"Arquivo '{nome_saida}' criado com sucesso! Todas as planilhas com o nome base '{nome_base}' foram unificadas.")
    
def exibir_caixa_mensagem():
    root = tk.Tk()
    root.withdraw()
    messagebox.showinfo("Programa Finalizado", "O programa foi concluído. Clique OK para confirmar.")

def clicar(posicao):
    pyautogui.click(posicao)

def digitar(texto):
    pyautogui.write(texto)

def BuscarItem():
    global items, identificador_item
    identificador_item = items[aux]

    #pyautogui.rightClick((mouse1x, mouse1y))
    
    #time.sleep(1)
    
    pyautogui.rightClick((mouse2x, mouse2y))
    
    time.sleep(1)
    
    digitar(identificador_item)
    
    time.sleep(1)
    
    pyautogui.click((mouse3x, mouse3y))
    
    time.sleep(1)
    
    pyautogui.press('=')  
    
    time.sleep(1)
    
    pyautogui.press('esc')
    
    #time.sleep(1)
    
    #pyautogui.rightClick((mouse1x, mouse1y))

def BuscarItemSanguine():
    global items_grand_sanguine, identificador_item_grand_sanguine
    identificador_item_grand_sanguine = items_grand_sanguine[aux]

    #pyautogui.rightClick((mouse1x, mouse1y))
    
    #time.sleep(1)
    
    pyautogui.rightClick((mouse2x, mouse2y))
    
    time.sleep(1)
    
    digitar(identificador_item_grand_sanguine)
    
    time.sleep(1)
    pyautogui.click((mouse4x, mouse4y))
    
    time.sleep(1)
    
    pyautogui.press('=')  
    
    time.sleep(1)
    
    pyautogui.press('esc')
    
    #time.sleep(1)
    
    #pyautogui.rightClick((mouse1x, mouse1y))

def achar_servidor_e_item():
    global cell_sell_offer, cell_buy_offer, identificador_item, identificador_servidor, excel_path, ws, wb
    
    # Iterar sobre as linhas da planilha para encontrar a posição
    servidor_col_idx = 1  # Coluna B (0-indexed seria 1)
    item_col_idx = 2      # Coluna C (0-indexed seria 2)
    
    servidor_row = None
    item_row = None

    for row in ws.iter_rows(min_row=2, max_col=ws.max_column, max_row=ws.max_row):
        if row[servidor_col_idx].value == identificador_servidor:
            servidor_row = row[servidor_col_idx].row
            if row[item_col_idx].value == identificador_item:
                item_row = row[item_col_idx].row
                break
    
    # Verificar se encontramos algum resultado
    if item_row:
        print(f"Servidor '{identificador_servidor}' com item '{identificador_item}' encontrado na linha {item_row}.")
        # Definir a coluna
        coluna_sell_offer_letra = 'E'
        cell_sell_offer = f"{coluna_sell_offer_letra}{item_row}"
        coluna_buy_offer_letra = 'F'
        cell_buy_offer = f"{coluna_buy_offer_letra}{item_row}"
        print(cell_sell_offer)
        print(cell_buy_offer)
    else:
        print(f"Servidor '{identificador_servidor}' com item '{identificador_item}' não encontrado.")

def achar_servidor_e_item_sanguine():
    global cell_buy_offer, cell_sell_offer, identificador_item_grand_sanguine, identificador_servidor, excel_path, ws, wb
    
    # Iterar sobre as linhas da planilha para encontrar a posição
    servidor_col_idx = 1  # Coluna B (0-indexed seria 1)
    item_col_idx = 2      # Coluna C (0-indexed seria 2)
    
    servidor_row = None
    item_row = None

    for row in ws.iter_rows(min_row=2, max_col=ws.max_column, max_row=ws.max_row):
        if row[servidor_col_idx].value == identificador_servidor:
            servidor_row = row[servidor_col_idx].row
            if row[item_col_idx].value == identificador_item_grand_sanguine:
                item_row = row[item_col_idx].row
                break
    
    # Verificar se encontramos algum resultado
    if item_row:
        print(f"Servidor '{identificador_servidor}' com item '{identificador_item_grand_sanguine}' encontrado na linha {item_row}.")
        # Definir a coluna
        coluna_sell_offer_letra = 'E'
        cell_sell_offer = f"{coluna_sell_offer_letra}{item_row}"
        coluna_buy_offer_letra = 'F'
        cell_buy_offer = f"{coluna_buy_offer_letra}{item_row}"
        print(cell_sell_offer)
        print(cell_buy_offer)
    else:
        print(f"Servidor '{identificador_servidor}' com item '{identificador_item_grand_sanguine}' não encontrado.")

def preencher_valor():
    # localizar a linha exatada do itemXservidor e preencher os valores no preço de compra e preço de venda
    global cell_buy_offer, cell_sell_offer, sell_offer, buy_offer, excel_path, wb, ws

    print(f"sell_offer inicial: {sell_offer}")
    print(f"buy_offer inicial: {buy_offer}")

    # Verificar se as variáveis globais estão definidas corretamente
    if not all([cell_buy_offer, cell_sell_offer, sell_offer, buy_offer, excel_path]):
        print("Uma ou mais variáveis globais não estão definidas corretamente.")
        return
    
    # Verificar e definir valores padrão caso não sejam encontrados
    if sell_offer == '' or sell_offer == '0' or sell_offer == None or sell_offer == 'None':
        print("sell_offer não encontrado, definindo como 0.")
        sell_offer = '0'
    if buy_offer == '' or buy_offer == '0' or buy_offer == None or buy_offer == 'None':
        print("buy_offer não encontrado, definindo como 0.")
        buy_offer = '0'
   
    print(f"sell_offer inicial: {sell_offer}")
    print(f"buy_offer inicial: {buy_offer}")

    # Preencher as celulas
    ws[cell_sell_offer] = int(sell_offer)
    ws[cell_buy_offer] = int(buy_offer)

    print(f"Valores preenchidos: Preço de Compra na célula {cell_sell_offer} e Preço de Venda na célula {cell_buy_offer}.")
    print(f"")

def remover_arquivos(file_path):
    global x
    os.remove(file_path)
    print(f"Arquivo {file_path} removido com sucesso.")
    print(f"")
    x = ''

def pegar_nome_arquivo_png():
    global x, image_path
    # Listar todos os arquivos na pasta especificada
    arquivos = os.listdir(image_path)
    
    # Procurar o primeiro arquivo com a extensão .png
    for arquivo in arquivos:
        if arquivo.endswith('.png'):
            x = arquivo  # Atribuir o nome do arquivo PNG à variável global x
            print(x)
            return 
               
def extrair_valor_img(regiao):
    global x
    # atualiza o caminho da imagem com o valor de X
    image_path = os.path.join(r'C:\Users\joaov\AppData\Local\Tibia\packages\Tibia\screenshots',x)

    # Carregar a imagem
    img = Image.open(image_path)
    
    # Cortar a região da imagem
    area = img.crop(regiao)
    
    # Extrair texto da região
    texto = pytesseract.image_to_string(area, lang='eng')
    
    # Procurar o primeiro valor numérico no texto extraído
    for word in texto.split():
        try:
            # Remover pontos e vírgulas
            numero = word.replace(',', '').replace('.', '').replace('$', '')
            return int(numero)
        except ValueError:
            continue
    
    return None 

def salvar_db():
    global wb

    wb.save(excel_path)

def mainItem():

    global item, sell_offer, buy_offer, x, aux
    num_items = len(items)

    while aux < num_items:
        item = items[aux]
        BuscarItem()
        pegar_nome_arquivo_png()
        achar_servidor_e_item()
        sell_offer = str(extrair_valor_img(REGIAO_SELL_OFFER))
        buy_offer = str(extrair_valor_img(REGIAO_BUY_OFFER))
        preencher_valor()
        image_path = os.path.join(r'C:\Users\joaov\AppData\Local\Tibia\packages\Tibia\screenshots',x)
        remover_arquivos(image_path)
        image_path = os.path.join(r'C:\Users\joaov\AppData\Local\Tibia\packages\Tibia\screenshots',x)
        x = ''
        aux += 1

def mainItemSanguine():
    global item_grand_sanguine, sell_offer, buy_offer, x, aux
    num_items_grand_sanguine = len(items_grand_sanguine)
    aux = 0 # usa o mesmo aux do mainItem e reseta ao iniciar a função

    while aux < num_items_grand_sanguine:
        item_grand_sanguine = items_grand_sanguine[aux]
        BuscarItemSanguine()
        pegar_nome_arquivo_png()
        achar_servidor_e_item_sanguine()
        sell_offer = str(extrair_valor_img(REGIAO_SELL_OFFER))
        buy_offer = str(extrair_valor_img(REGIAO_BUY_OFFER))
        preencher_valor()
        image_path = os.path.join(r'C:\Users\joaov\AppData\Local\Tibia\packages\Tibia\screenshots',x)
        remover_arquivos(image_path)
        image_path = os.path.join(r'C:\Users\joaov\AppData\Local\Tibia\packages\Tibia\screenshots',x)
        x = ''
        aux += 1