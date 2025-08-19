Projeto teste para coleta automatizada de preço de compra e venda de itens no tibia utilizando o sistema de print dentro do próprio jogo, uma vez que o jogo possui um sistema anti-bots que impede  o print ou gravação de tela por meios fora do jogo.

Com os valores salvos em um banco de dados será possível comparar valores de itens entre servidores através da própria moeda do jogo (Tibia coin) e assim fazer comparações, com o intuito de comprar em servidores onde determinado item é mais barato e revender onde o mesmo for mais caro.

O projeto envolve coleta de dados, armanezamento em um banco de dados e limpeza para ser utilizado em forma de dashboards através de programas como PowerBI. 

Esse projeto tem o intuito de deixar a visualização de dados mais transparente e facilitar trocas entre jogadores.

Bibliotecas utilizadas:
    - pyautogui (manipulação do mouse)
    - openpyxl (manipulação de arquivo excel para armazenamento de dados)
    - time (medição de tempo de execução do programa)
    - pytesseract (fazer reconhecimento visual dos valores em tela e convertê-los em texto)
    - os (interface gráfica)
    - tkinter (abrir caixa de textos/alertas)

A adicionar:
    - PIL (manipulação de imagens)
    - xlsxwriter (manipulação de arquivo excel como editar colunas já existentes)
    - streamlit (criação de dashboards)
    - plotly (criação de gráficos)

Após a instalação das bibliotecas, basta o usuário editar os items que queira buscar, adicionando ou removendo do array items e items_grand_sanguine.

Com o jogo aberto e o personagem posicionado na frente do depot o usuário deverá fazer um teste para verificar a posição do mouse onde será necessário executar parte do código e salvar os valores da posicao_mouse, esse valor irá variar devido a: tamanho de monitor, tamanho da janela do jogo e tamanho da resolução da tela.

Com o teste feito o programa principal poderá ser executado, tendo que informar o servidor para o programa. Feito isso deixe a janela do jogo no monitor e o programa executará as ações de coleta e armazenamento de valores.

Ao encerrar o programa irá informar o tempo decorrido.

O arquivo .pbix é um teste feito em PowerBI para demonstração e afins de estudos utilizando a plataforma

# No excel:
- B = Server name
- C = Item name
- D = Categoria
- E = Buy Offer
- F = Sell Offer

# configuracoes da janela do tibia
- apenas local chat e log abertos
- 1 barra lateral na esquerda e uma na direita
- Barra de vida no topo com barra de XP
- Control buttons minimizado
- personagem parado com o "DP" na sua frente (acima do personagem)