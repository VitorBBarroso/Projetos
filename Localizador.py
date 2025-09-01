import os
import shutil
import tkinter as tk
from tkinter import filedialog, ttk, messagebox

PASTAS_MAE = {
    "Works":        [   r'T:\01 - ARQUIVO DIGITALIZADO\01 - DIVERSOS\WORKS',
                        r'T:\01 - ARQUIVO DIGITALIZADO\03 - PRONTUARIO DIGITALIZADO\02 - WORKS',
                        r'T:\01 - ARQUIVO DIGITALIZADO\02 - FOLHAS DE PONTO\WORKS',
                        r'S:\02 - WORKS',
                        r'R:\WORKS',
                        r'V:\03.RESCISAO'],
   
    "Partner":      [   r'T:\01 - ARQUIVO DIGITALIZADO\01 - DIVERSOS\PARTNER',
                        r'T:\01 - ARQUIVO DIGITALIZADO\03 - PRONTUARIO DIGITALIZADO\03 - PARTNER',
                        r'T:\01 - ARQUIVO DIGITALIZADO\02 - FOLHAS DE PONTO\PARTNER',
                        r'S:\04 - PARTNER',
                        r'R:\PARTNER SECURITY',
                        r'V:\PARTNER\PARTNER SECURITY\TERMO DE RESCISAO'],
   
    "Pressseg":     [   r'T:\01 - ARQUIVO DIGITALIZADO\01 - DIVERSOS\PRESSSEG',
                        r'T:\01 - ARQUIVO DIGITALIZADO\03 - PRONTUARIO DIGITALIZADO\01 - PRESSSEG',
                        r'T:\01 - ARQUIVO DIGITALIZADO\02 - FOLHAS DE PONTO\PRESSSEG',
                        r'S:\01 - PRESSSEG',
                        r'R:\PRESSSEG',
                        r'V:\RH-PRESSSEG\05 RESCISAO'],
    
    "Qualitech":    [   r'T:\01 - ARQUIVO DIGITALIZADO\01 - DIVERSOS\QUALITECH',
                        r'T:\01 - ARQUIVO DIGITALIZADO\03 - PRONTUARIO DIGITALIZADO\04 - QUALITECH',
                        r'T:\01 - ARQUIVO DIGITALIZADO\02 - FOLHAS DE PONTO\QUALITECH',
                        r'S:\03 - QUALITECH',
                        r'R:\QUALITECH',
                        r'V:\WORKS-QUALITECH']
}



def buscar_arquivos(empresa, re_busca, nome_funcionario):
    pasta_mae = PASTAS_MAE.get(empresa)
    if not pasta_mae or not os.path.exists(pasta_mae):
        messagebox.showerror("Erro", f"Pasta da empresa '{empresa}' não foi encontrada.")
        return

    encontrados = []
    for root, dirs, files in os.walk(pasta_mae):
        for arquivo in files:
            nome_sem_ext = os.path.splitext(arquivo)[0]
            if re_busca in nome_sem_ext:
                caminho_completo = os.path.join(root, arquivo)
                encontrados.append(caminho_completo)

    if not encontrados:
        messagebox.showinfo("Busca finalizada", "Nenhum arquivo com o RE informado foi encontrado.")
        return

    destino = os.path.join(os.getcwd(), nome_funcionario)
    os.makedirs(destino, exist_ok=True)

    for caminho in encontrados:
        try:
            shutil.copy(caminho, destino)
        except Exception as e:
            print(f"[ERRO] Falha ao copiar {caminho}: {e}")

    messagebox.showinfo("Sucesso", f"{len(encontrados)} arquivo(s) copiado(s) para a pasta: {destino}")
    
#Visualização
janela = tk.Tk()
janela.title("Localizador de Arquivos por RE")
janela.geometry("400x250")
janela.resizable(False, False)

fonte = ("Arial", 11)

tk.Label(janela, text="Empresa:", font=fonte).pack(pady=5)
combo_empresa = ttk.Combobox(janela, values=list(PASTAS_MAE.keys()), state="readonly", font=fonte)
combo_empresa.pack()

tk.Label(janela, text="RE (somente números):", font=fonte).pack(pady=5)
entry_re = tk.Entry(janela, font=fonte)
entry_re.pack()

tk.Label(janela, text="Nome do Funcionário:", font=fonte).pack(pady=5)
entry_nome = tk.Entry(janela, font=fonte)
entry_nome.pack()

def iniciar_busca():
    empresa = combo_empresa.get()
    re = entry_re.get().strip()
    nome = entry_nome.get().strip()

    if not (empresa and re and nome):
        messagebox.showwarning("Campos obrigatórios", "Preencha todos os campos.")
        return

    buscar_arquivos(empresa, re, nome)

tk.Button(janela, text="Buscar Arquivos", font=fonte, bg="#4CAF50", fg="white", command=iniciar_busca).pack(pady=15)

janela.mainloop()
