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
    - PIL (manipulação de imagens)