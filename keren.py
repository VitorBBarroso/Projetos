import pyautogui
import time
import keyboard
import threading
import tkinter as tk
import tkinter.messagebox as messagebox

ativo = False
executando = True

def loop_automacao():
    global ativo, executando
    time.sleep(20)
    while executando:
        if ativo:
            pyautogui.hotkey('ctrl', 'c')
            time.sleep(0.3)

            pyautogui.hotkey('alt', 'tab')
            time.sleep(0.3)

            pyautogui.hotkey('ctrl', 'v')
            time.sleep(0.3)

            pyautogui.press('down')
            time.sleep(0.3)

            pyautogui.hotkey('alt', 'tab')
            time.sleep(0.3)

            pyautogui.press('down')
            time.sleep(0.3)
        else:
            time.sleep(0.1)

def iniciar_loop():
    global ativo
    ativo = True
    status_var.set("Status: EM EXECUÇÃO")

def parar_loop():
    global ativo
    ativo = False
    status_var.set("Status: PARADO")

def sair():
    global executando
    resposta = messagebox.askyesno("Sair", "Deseja realmente encerrar o programa?")
    if resposta:
        executando = False
        root.destroy()

def sobre():
    messagebox.showinfo("Sobre", "Automação de cópia entre Excel, para o HK ou outro app.\nDesenvolvido em Python.")

# GUI principal
root = tk.Tk()
root.title("Automação Excel")
root.geometry("400x250")

status_var = tk.StringVar()
status_var.set("Status: PARADO")

# Barra de menus
menu_bar = tk.Menu(root)

menu_arquivo = tk.Menu(menu_bar, tearoff=0)
menu_arquivo.add_command(label="Iniciar Loop", command=iniciar_loop)
menu_arquivo.add_command(label="Parar Loop", command=parar_loop)
menu_arquivo.add_separator()
menu_arquivo.add_command(label="Sair", command=sair)

menu_ajuda = tk.Menu(menu_bar, tearoff=0)
menu_ajuda.add_command(label="Sobre", command=sobre)

menu_bar.add_cascade(label="Arquivo", menu=menu_arquivo)
menu_bar.add_cascade(label="Ajuda", menu=menu_ajuda)
root.config(menu=menu_bar)

tk.Label(root, text="Automação de Transferência", font=("Arial", 16, "bold")).pack(pady=10)
tk.Label(root, textvariable=status_var, font=("Arial", 12)).pack(pady=10)

tk.Button(root, text="Iniciar", command=iniciar_loop, bg="green", fg="white", width=20, height=2).pack(pady=5)
tk.Button(root, text="Parar", command=parar_loop, width=20, height=2).pack(pady=5)
tk.Button(root, text="Sair", command=sair, bg="red", fg="white", width=20, height=2).pack(pady=5)


thread = threading.Thread(target=loop_automacao)
thread.daemon = True
thread.start()

root.mainloop()
