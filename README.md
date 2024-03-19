# Projeto AutoTag

O usuário pode clicar no botão "Selecionar" para abrir uma janela de diálogo e escolher um arquivo Excel (.xlsx ou .xls).

Entrada da Coluna:

O usuário pode inserir o número da coluna ou sua letra na entrada designada.

Extração e Colagem dos Dados:

Os dados da coluna selecionada são extraídos do arquivo Excel escolhido.

Usando a biblioteca openpyxl, os dados são lidos da planilha ativa.

Os dados são colados em uma área específica da tela usando a biblioteca pyautogui.

O processo de colagem é repetido para as primeiras 8 células da coluna selecionada.

Além disso, para facilitar a localização, foi implementado um rastreador de coordenadas do mouse. 

Código para auxiliar na localização de coordenadas.
>Rastromouse

# Para instalar as bibliotecas necessárias

1-tkinter
>pip install tk

2-openpyxl:
>pip install openpyxl

3-pyautogui:
>pip install pyautogui




