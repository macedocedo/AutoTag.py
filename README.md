![Python](https://img.shields.io/badge/Python-3776AB?style=for-the-badge&logo=python&logoColor=white)

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

Encontre o caminho para instalação das bibliotecas Python 
C:\Users\mateus\AppData\Local\Programs\Python\Python312\Scripts

![image](https://github.com/macedocedo/AutoTag.py/assets/84480587/b24c7c37-0c6a-4b3b-bce4-6d06b7c11a28)

Na barra de pesquisa do mesmo campo onde o Python foi instalado, dentro da pasta "Scripts", digite "cmd" e aguarde a abertura do prompt de comando.

![image](https://github.com/macedocedo/AutoTag.py/assets/84480587/54d5ee5e-527a-4abf-bf23-70af1c954682)

CMD
Em seguida instale as seguintes bibliotecas no prompt de comando.

1-tkinter
>pip install tkinter

2-openpyxl:
>pip install openpyxl

3-pyautogui:
>pip install pyautogui

4-Código para auxiliar na localização de coordenadas.
>Rastromouse

![image](https://github.com/macedocedo/AutoTag/assets/84480587/f6ad5602-516d-4d68-999f-cb31bd83dd63)




