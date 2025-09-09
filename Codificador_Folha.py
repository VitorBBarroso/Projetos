import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, simpledialog
import os
import string
import datetime

MESES = {
    'Janeiro': '01', 'Fevereiro': '02', 'Março': '03', 'Abril': '04',
    'Maio': '05', 'Junho': '06', 'Julho': '07', 'Agosto': '08',
    'Setembro': '09', 'Outubro': '10', 'Novembro': '11', 'Dezembro': '12'
}
ANOS = [str(ano) for ano in range(2020, 2026)]

def completar_zeros(valor, tamanho):
    return str(valor).zfill(tamanho)

def gerar_codigo(filial, re, mes, ano, sufixo):
    return (
        completar_zeros(filial, 4)  +
        completar_zeros(re, 7)      +
        completar_zeros(mes, 2)     +
        completar_zeros(ano, 4)     +
        completar_zeros(sufixo, 4)
    )

def letra_para_indice(letra):
    return string.ascii_uppercase.index(letra.upper())

def selecionar_arquivo_excel():
    return filedialog.askopenfilename(title="Selecione a planilha Excel", filetypes=[("Excel", "*.xlsx")])

def selecionar_pasta_destino():
    return filedialog.askdirectory(title="Selecione a pasta de destino")

def processar_planilha():
    caminho_excel = selecionar_arquivo_excel()
    if not caminho_excel:
        return

    pasta_destino = selecionar_pasta_destino()
    if not pasta_destino:
        return

    try:
        aba = simpledialog.askstring("Aba", "Digite o nome da aba onde estão os dados:")
        if not aba:
            messagebox.showwarning("Aviso", "Nome da aba não informado.")
            return

        df = pd.read_excel(caminho_excel, sheet_name=aba, header=None)

        letras_colunas = [f"{letra} → Índice {i}" for i, letra in enumerate(string.ascii_uppercase[:df.shape[1]])]
        colunas_texto = "\n".join(letras_colunas)
        messagebox.showinfo("Colunas disponíveis", f"As colunas encontradas foram:\n\n{colunas_texto}")

        letra_filial = simpledialog.askstring("Coluna Filial", "Digite a letra da coluna da Filial (ex: B):")
        letra_re = simpledialog.askstring("Coluna RE", "Digite a letra da coluna do RE (ex: C):")

        if not (letra_filial and letra_re):
            messagebox.showwarning("Aviso", "As letras das colunas de Filial e RE devem ser informadas.")
            return

        idx_filial = letra_para_indice(letra_filial)
        idx_re = letra_para_indice(letra_re)

    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao abrir planilha: {e}")
        return

    mes_nome = combo_mes.get()
    ano = combo_ano.get()
    sufixo = entrada_sufixo.get()

    if mes_nome not in MESES:
        messagebox.showwarning("Aviso", "Selecione um mês válido.")
        return
    if ano not in ANOS:
        messagebox.showwarning("Aviso", "Selecione um ano válido.")
        return
    if not sufixo.isdigit() or len(sufixo) > 4:
        messagebox.showwarning("Aviso", "Sufixo deve conter até 4 dígitos numéricos.")
        return

    mes_codigo = MESES[mes_nome]
    codigos = []

    try:
        for _, row in df.iterrows():
            filial = row[idx_filial]
            re = row[idx_re]
            codigo = gerar_codigo(filial, re, mes_codigo, ano, sufixo)
            codigos.append(codigo)
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao gerar códigos: {e}")
        return

    df_resultado = pd.DataFrame({'Código Gerado': codigos})
    
    # Adiciona timestamp ao nome do arquivo
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    nome_arquivo = f"codigos_gerados_{timestamp}.xlsx"
    caminho_saida = os.path.join(pasta_destino, nome_arquivo)

    try:
        df_resultado.to_excel(caminho_saida, index=False)
        messagebox.showinfo("Concluído", f"Arquivo salvo com sucesso:\n{caminho_saida}")
    except PermissionError:
        messagebox.showerror("Erro", f"Permissão negada ao salvar o arquivo:\n{caminho_saida}\n\nVerifique se o arquivo está aberto ou se você tem permissão para gravar na pasta.")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao salvar arquivo:\n{e}")

# Interface Tkinter
janela = tk.Tk()
janela.title("Gerador de Códigos")
janela.geometry("300x250")
janela.configure(bg="black")
fonte = ("Arial", 10)

tk.Label(janela, text="Mês:", bg="black", fg="white", font=fonte).pack(pady=2)
combo_mes = ttk.Combobox(janela, values=list(MESES.keys()), state="readonly", font=fonte)
combo_mes.pack()
combo_mes.set("Julho")

tk.Label(janela, text="Ano:", bg="black", fg="white", font=fonte).pack(pady=2)
combo_ano = ttk.Combobox(janela, values=ANOS, state="readonly", font=fonte)
combo_ano.pack()
combo_ano.set("2025")

tk.Label(janela, text="Sufixo:", bg="black", fg="white", font=fonte).pack(pady=2)
entrada_sufixo = tk.Entry(janela, font=fonte)
entrada_sufixo.pack()
entrada_sufixo.insert(0, "0001")

tk.Button(
    janela,
    text="Selecionar Planilha e Gerar",
    command=processar_planilha,
    bg="#333333", fg="white", font=fonte
).pack(pady=10)

janela.mainloop()