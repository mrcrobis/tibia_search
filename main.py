from dictionary import items, items_grand_sanguine
from datetime import time

from global_var import x, aux, excel_path, image_path, mouse2x, mouse2y, mouse3x, mouse3y, mouse4x, mouse4y, sheet_name, wb, ws, identificador_servidor, mouse1x, mouse1y

from functions import achar_servidor_e_item, achar_servidor_e_item_secundario, BuscarItem, BuscarItemSecundario, clicar, combinar_planilhas, copiar_e_formatar_excel, digitar, exibir_caixa_mensagem, extrair_valor_img, mainItem, mainItemSecundario, pegar_nome_arquivo_png, preencher_valor, remover_arquivos, salvar_db

#test -----------------------------------------------------------------------

#arquivo_origem = excel_path
#copiar_e_formatar_excel(arquivo_origem, identificador_servidor)

#print('Informa o nome_base')
#nome_base = input()
#nome_base = "07-03-2025"  
#pasta_entrada = "db_sheets/{}".format(nome_base)
#combinar_planilhas(pasta_entrada, nome_base)

# testar a posicao do mouse
#posicao_mouse = pyautogui.position()

#print(f"A posição atual do mouse é: {posicao_mouse}")

# --------------------------------------------------------------------------

#execution -----------------------------------------------------------------

"""
start_time = time.time()
time.sleep(5)
mainItem()
mainItemSanguine()
salvar_db()
end_time = time.time()
execution_time = end_time - start_time
exibir_caixa_mensagem()
print(f"Tempo de execução: {execution_time:.2f} segundos")
"""

# ---------------------------------------------------------------------------


"""
TODO
* Implementar criação de arquivos excel
    - Será possível usar várias DB dentro do powerBI para fazer a assimilação de valores de preço de itens sem ter algum transtorno?
    - Essa é a forma mais otimizada de resolver esse problema?
* Implementar quantidade de itens/ofertas no market 
    - da mesma forma que o programa pega o preço dos itens usando a imagem ele pode pegar a quantidade
    de itens a venda em coluna e somar toda a quantidade, o lado negativo é que a quantidade de itens
    não irá bater com a quantidade de itens com o preço presente no db
* Implementar função para calculo de profit entre compra e venda de itens entre servidores 
    - listar quais e quantos itens estão disponíves para compra e venda entre os 2 servidores escolhidos,
    calcular o retorno monetário "ROI"
    - feito no powerBI (?)
* Implementar função de pesquisa de items com Tier (T1 - T10)
    - muito trabalho, 10 linhas no banco de dados para cada item classificação 4, tirando os outros itens
    com classificação inferior
"""
