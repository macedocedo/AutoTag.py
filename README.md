# Projeto AutoTag

Este código Python cria uma GUI usando o módulo tkinter. Ele permite ao usuário selecionar um arquivo Excel, filtrar dados de uma coluna desse arquivo e colar parte desses dados em outra parte da tela usando pyautogui. O usuário pode interagir com a interface gráfica para realizar essas ações de forma intuitiva.

# Como funciona?

-Seleção de Arquivo Excel:
O usuário pode selecionar um arquivo Excel (.xlsx ou .xls) usando um botão de seleção.
O caminho do arquivo selecionado é exibido em um campo de entrada na interface.

-Filtragem de Dados:
O usuário pode digitar um filtro para filtrar os dados de uma coluna do arquivo Excel.
A lista de opções na interface é atualizada dinamicamente para exibir apenas os valores que correspondem ao filtro.

-Seleção de Coluna:
Uma combobox exibe os valores da primeira linha do arquivo Excel selecionado.
O usuário pode selecionar uma coluna a partir desses valores.

-Colagem de Dados:
Ao clicar em um botão, o programa cola os dados da coluna selecionada em outra parte da tela.
Isso é feito utilizando a biblioteca pyautogui para interagir com a interface gráfica.
Feedback de Status:

Um campo de rótulo na interface fornece feedback sobre o status das operações, como sucesso ou erro.

# Para instalar as bibliotecas necessárias

1-tkinter
>pip install tk

2-openpyxl:
>pip install openpyxl

3-pyautogui:
>pip install pyautogui

4-Código para auxiliar na localização de coordenadas.
>Rastromouse




