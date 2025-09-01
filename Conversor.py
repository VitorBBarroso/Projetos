from dbfread import DBF
import csv
import tkinter as tk
from tkinter import filedialog, ttk, messagebox

def converter():

    arquivo_dbf = filedialog.askopenfilename(
        title="Selecione o arquivo DBF",
        filetypes=[("Arquivos DBF", "*.dbf"), ("Todos os arquivos", "*.*")]
    )
    if not arquivo_dbf:
        messagebox.showerror("Erro", "Nenhum arquivo selecionado!")
        return

    # Selecionar local para salvar CSV
    arquivo_csv = filedialog.asksaveasfilename(
        title="Salvar arquivo CSV como",
        defaultextension=".csv",
        filetypes=[("Arquivos CSV", "*.csv")]
    )
    if not arquivo_csv:
        messagebox.showerror("Erro", "Nenhum local de salvamento selecionado!")
        return

    try:
        tabela = DBF(arquivo_dbf, encoding='latin1')

        total_registros = len(tabela)  # usado para progresso
        progresso["maximum"] = total_registros
        progresso["value"] = 0
        root.update_idletasks()
        
        with open(arquivo_csv, 'w', newline='', encoding='utf-8-sig') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(tabela.field_names)

            for i, registro in enumerate(tabela, start=1):
                writer.writerow([registro[campo] for campo in tabela.field_names])
                progresso["value"] = i
                root.update_idletasks() 

        messagebox.showinfo("Sucesso", f"Conversão concluída!\nArquivo salvo em:\n{arquivo_csv}")

    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro: {e}")


# --- Interface gráfica ---
root = tk.Tk()
root.title("Conversor DBF para CSV")
root.geometry("400x200")

frame = ttk.Frame(root, padding=20)
frame.pack(expand=True, fill="both")

titulo = ttk.Label(frame, text="Conversor DBF → CSV", font=("Arial", 14, "bold"))
titulo.pack(pady=10)

botao = ttk.Button(frame, text="Selecionar e Converter", command=converter)
botao.pack(pady=10)

progresso = ttk.Progressbar(frame, orient="horizontal", length=300, mode="determinate")
progresso.pack(pady=20)

root.mainloop()
