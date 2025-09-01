import pyautogui
import time
import threading
import tkinter as tk
import tkinter.messagebox as messagebox
from PIL import Image, ImageTk
import pygetwindow as gw
import sys
import ctypes
import platform

ativo = False
executando = True
flash_ativo = False

def janela_ativa_e_chrome():
    try:
        janela = gw.getActiveWindow()
        return janela and "Chrome" in janela.title
    except:
        return False

def reset_flash():
    global flash_ativo
    flash_ativo = False

def piscar_janela():
    global flash_ativo
    if platform.system() == "Windows" and not flash_ativo:
        flash_ativo = True
        root.deiconify()
        root.lift()
        root.attributes('-topmost', True)

        hwnd = ctypes.windll.user32.GetForegroundWindow()
        FLASHW_ALL = 3

        class FLASHWINFO(ctypes.Structure):
            _fields_ = [('cbSize', ctypes.c_uint),
                        ('hwnd', ctypes.c_void_p),
                        ('dwFlags', ctypes.c_uint),
                        ('uCount', ctypes.c_uint),
                        ('dwTimeout', ctypes.c_uint)]

        flash_info = FLASHWINFO(ctypes.sizeof(FLASHWINFO),
                                hwnd,
                                FLASHW_ALL,
                                5,
                                0)
        ctypes.windll.user32.FlashWindowEx(ctypes.byref(flash_info))
        root.after(5000, reset_flash)

def loop_func():
    global ativo, executando, flash_ativo
    while executando:
        if ativo:
            if janela_ativa_e_chrome():
                aviso_var.set("")
                root.attributes('-topmost', False)
                root.iconify()
                flash_ativo = False 
                pyautogui.hotkey('ctrl', 'tab')
            else:
                aviso_var.set("Chrome não está em foco!")
                piscar_janela()
        time.sleep(1) # Ajuste o tempo de espera.

def ativar():
    global ativo
    ativo = True
    status_var.set("Status: ATIVADO")
    print("\033[1m\nStatus: Ativado\033[0m")

def desativar():
    global ativo
    ativo = False
    aviso_var.set("")
    status_var.set("Status: DESATIVADO")
    root.attributes('-topmost', False)
    print("\033[1m\nStatus: Desativado\033[0m")

def encerrar():
    global executando
    resposta = messagebox.askyesno("Confirmação", "Deseja realmente encerrar o programa?")
    if resposta:
        executando = False              
        root.destroy()          
        print("\033[1m\nStatus: Encerrado\033[0m")

root = tk.Tk()
root.title("Loop de Páginas")
root.geometry("350x380")

try:
    img_icon = Image.open("works.jpg").resize((32, 32))
    icon = ImageTk.PhotoImage(img_icon)
    root.iconphoto(True, icon)
except Exception as e:
    print(f"Erro ao definir ícone: {e}")

thread_loop = threading.Thread(target=loop_func)
thread_loop.daemon = True
thread_loop.start()

try:
    imagem = Image.open("works.jpg").resize((80, 80))
    imagem_tk = ImageTk.PhotoImage(imagem)
    tk.Label(root, image=imagem_tk).pack(pady=5)
except Exception as e:
    print(f"Erro ao carregar imagem no corpo: {e}")

status_var = tk.StringVar()
status_var.set("Status: DESATIVADO")
tk.Label(root, textvariable=status_var, font=("Arial", 14)).pack(pady=10)

aviso_var = tk.StringVar()
tk.Label(root, textvariable=aviso_var, font=("Arial", 13), fg="red").pack(pady=5)

tk.Button(root, text="Ativar", command=ativar, fg="white", bg="blue", height=2).pack(fill='x', expand=True, padx=20, pady=5)
tk.Button(root, text="Desativar", command=desativar, height=2).pack(fill='x', expand=True, padx=20, pady=5)
tk.Button(root, text="Encerrar", command=encerrar, fg="white", bg="red", height=2).pack(fill='x', expand=True, padx=20, pady=5)

root.mainloop()
