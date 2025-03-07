from dictionary import items, items_grand_sanguine
from datetime import time
from global_var import x, aux, aux_sanguine, excel_path, image_path, mouse2x, mouse2y, mouse3x, mouse3y, mouse4x, mouse4y, sheet_name, wb, ws
from functions import achar_servidor_e_item, achar_servidor_e_item_sanguine, BuscarItem, BuscarItemSanguine, clicar, combinar_planilhas, copiar_e_modificar_excel, digitar, exibir_caixa_mensagem, extrair_valor_img, mainItem, mainItemSanguine, pegar_nome_arquivo_png, preencher_valor, remover_arquivos, salvar_db

#test -----------------------------------------------------------------------

print(f"Informe o Servidor: ")
identificador_servidor = input() 

arquivo_origem = excel_path
copiar_e_modificar_excel(arquivo_origem, identificador_servidor)



#pasta_entrada = "db_sheets"
#nome_base = "06-03-2025"  
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
* Corrigir o erro de import entre os arquivos
* Implementar criação de arquivos excel
    - Fazer com que o programa apague as colunas dos outros servidores para que fique salvo apenas o do servidor em especifico (otimizando a pesquisa por itens de um servidor em especifico) caso necessário, concatenar novamente os valores do DB de cada servidor para fazer tudo em apenas uma grande base
    - Será possível usar várias DB dentro do powerBI para fazer a assimilação de valores de preço de itens sem ter algum transtorno?
    - Essa é a forma mais otimizada de resolver esse problema?
    - Como fazer para que o programa pegue todas as planilhas e junte tudo em apenas uma planilha apenas jogando os valores das outras planilhas no mesmo indice que a planilha principal terá, porém sem os valores
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
