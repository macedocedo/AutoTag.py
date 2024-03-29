import pyautogui
import tkinter as tk

def iniciar_selecao():
    label_instrucao.config(text="Posicione o cursor...")
    botao_selecionar.config(state=tk.DISABLED)
    janela.after(3000, capturar_posicao)  # Aguarda 3 segundos antes de capturar a posição
    

def capturar_posicao():
    posicao = pyautogui.position()
    label_posicao.config(text="Posição do cursor: {}".format(posicao))
    label_instrucao.config(text="")
    botao_selecionar.config(state=tk.NORMAL)

# Criar janela
janela = tk.Tk()
janela.title("Seleção de Posição do Cursor")

# Criar rótulo de instrução
label_instrucao = tk.Label(janela, text="Clique em 'Selecionar Posição' e posicione o cursor.")
label_instrucao.pack(pady=5)

# Botão para iniciar seleção de posição
botao_selecionar = tk.Button(janela, text="Selecionar Posição", command=iniciar_selecao)
botao_selecionar.pack(pady=5)

# Rótulo para exibir a posição
label_posicao = tk.Label(janela, text="")
label_posicao.pack(pady=10, padx=20)

# Loop principal da janela
janela.mainloop()
