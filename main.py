from datetime import time
from global_var import *
from functions import *
from dictionary import *

#test -----------------------------------------------------------------------

copiar_e_formatar_excel(excel_path, identificador_servidor)

#print('Informa o nome_base')
#nome_base = input()
#nome_base = "10-03-2025"  
#pasta_entrada = "db_sheets/{}".format(nome_base)
#combinar_planilhas(pasta_entrada, nome_base)

# testar a posicao do mouse
#posicao_mouse = pyautogui.position()

#print(f"A posição atual do mouse é: {posicao_mouse}")


# Coleta de dados ------------------------------------------------------------

#delay_inicio(3)
#inicia_Timer()
#copiar_e_formatar_excel
#main_Item()
#main_Item_Secundario()
#salvar_db()
#finaliza_Timer()
#exibir_caixa_mensagem()

# Combinacao de planilhas -----------------------------------------------------

# exemplo = 10-03-2025
#print('Informa o nome_base')
#nome_base = input()
#nome_base = "10-03-2025"  
#pasta_entrada = "db_sheets/{}".format(nome_base)
#combinar_planilhas(pasta_entrada, nome_base)

# Atualizacao preco TC---------------------------------------------------------

atualiza_Preco_TC_arquivo_combined(r'D:\Programming\tibia_search\db_sheets\10-03-2025\10-03-2025_combined.xlsx')


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
* Atualizar o db após o lançamento oficial da nova classe Monk, adicionando novos items como poções, colares, aneis, armas e equipamentos 
"""
