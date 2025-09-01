import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import re
from openpyxl import load_workbook
from openpyxl.styles import Font, Border, Side, Alignment
from openpyxl.drawing.image import Image

# Caminho da logo
logo = r'Y:\Power BI\Scritps\python\logo.jpg'

def gerar_relatorio():
    try:
        mes = combo_mes.get()
        ano = entry_ano.get()

        if not mes or not ano.isdigit():
            messagebox.showerror("Erro", "Informe corretamente o mês e o ano!")
            return

        arquivo_excel = filedialog.askopenfilename(
            title="Selecione a base de dados",
            filetypes=[("Planilhas", "*.xlsx *.xls *.ods *.csv")]
        )
        if not arquivo_excel:
            return

        ext = os.path.splitext(arquivo_excel)[1].lower()
        if ext == ".xlsx":
            df = pd.read_excel(arquivo_excel, engine="openpyxl")
        elif ext == ".xls":
            df = pd.read_excel(arquivo_excel, engine="xlrd")
        elif ext == ".ods":
            df = pd.read_excel(arquivo_excel, engine="odf")
        elif ext == ".csv":
            df = pd.read_csv(arquivo_excel, sep=";", encoding="latin1")
        else:
            messagebox.showerror("Erro", "Formato de arquivo não suportado!")
            return

        df.columns = df.columns.str.lower().str.strip()

        coluna_posto = "posto"
        coluna_re = "re"
        coluna_nome = "nome"
        coluna_funcao = "desc_cargo"
        coluna_csituacao = "csituacao"
        coluna_csithoje = "csithoje"

        # Filtragem
        df = df[df[coluna_csituacao].isin([10, 11])]
        df = df[~df[coluna_csithoje].isin([2, 13, 14])]

        def limpar_nome_aba(nome):
            nome_limpo = re.sub(r'[:\\/*?\[\]]', '_', str(nome))
            return nome_limpo[:31]

        saida = filedialog.asksaveasfilename(
            title="Salvar Relatório",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile="relatorio_por_posto.xlsx"
        )
        if not saida:
            return

        with pd.ExcelWriter(saida, engine="openpyxl") as writer:
            abas_criadas = 0
            for posto, bloco_posto in df.groupby(coluna_posto):
                if bloco_posto.empty:
                    continue

                resultado = (
                    bloco_posto.groupby([coluna_re, coluna_nome, coluna_funcao])
                    .size()
                    .reset_index()
                    .drop(columns=0, errors="ignore")
                )

                nome_aba = "Indefinido" if pd.isna(posto) else limpar_nome_aba(posto)
                resultado.to_excel(writer, sheet_name=nome_aba, index=False, startrow=14, startcol=0)  # startcol=0
                abas_criadas += 1

            if abas_criadas == 0:
                resumo = pd.DataFrame({"Mensagem": ["Nenhum dado válido encontrado para gerar o relatório."]})
                resumo.to_excel(writer, sheet_name="Resumo", index=False)

        # --- Formatação Excel ---
        wb = load_workbook(saida)
        fonte = Font(name="Arial", size=9)
        fonte_cnpj = Font(name="Arial", size=8)
        fonte_negrito = Font(name="Arial", size=10, bold=True)
        fonte_titulo = Font(name="Arial", size=12, bold=True)
        alinhamento_centro = Alignment(horizontal="center", vertical="center")
        borda = Border(
            left=Side(border_style="thin"),
            right=Side(border_style="thin"),
            top=Side(border_style="thin"),
            bottom=Side(border_style="thin")
        )

        for aba in wb.sheetnames:
            ws = wb[aba]

            # Logo centralizado
            if os.path.exists(logo):
                img = Image(logo)
                img.width, img.height = 134, 104  # dimensões da imagem
                coluna_b = ws.column_dimensions['B'].width
                if not coluna_b:
                    coluna_b = 32.43
                ws.add_image(img, "B1")


            # Função para aplicar borda em células mescladas
            def aplicar_borda(range_str):
                start_cell, end_cell = range_str.split(":")
                start_col, start_row = ws[start_cell].column, ws[start_cell].row
                end_col, end_row = ws[end_cell].column, ws[end_cell].row
                for row in ws.iter_rows(min_row=start_row, max_row=end_row,
                                        min_col=start_col, max_col=end_col):
                    for cell in row:
                        cell.border = borda

            # Cabeçalho
            cabecalhos = [
                ("A7:C7", "Endereço: R. Conselheiro Ribas, 297 - Vila Anastácio, São Paulo - SP, 05093-060", fonte),
                ("A8:C8", "CNPJ: 56.419.492/0001-09     Telefone: (11) 4563-9017", fonte),
                ("A9:C9", "CONTRATO Nº 83/SME/2024 - P.E. 23/SME/2023     Processo Administrativo 6016.2024/0043414-0", fonte_cnpj)
            ]

            for range_str, texto, fnt in cabecalhos:
                start_cell = range_str.split(":")[0]
                ws.merge_cells(range_str)
                ws[start_cell].value = texto
                ws[start_cell].font = fnt
                ws[start_cell].alignment = alinhamento_centro
                aplicar_borda(range_str)

            # Nome do posto
            ws.merge_cells("A10:C10")
            ws["A10"].value = aba
            ws["A10"].font = fonte_titulo
            ws["A10"].alignment = alinhamento_centro
            aplicar_borda("A10:C10")

            # Mês de referência
            ws.merge_cells("A11:C11")
            ws["A11"].value = f"Mês de Referência: 01 a 31 de {mes} de {ano}"
            ws["A11"].font = fonte_negrito
            ws["A11"].alignment = alinhamento_centro
            aplicar_borda("A11:C11")

            # Cabeçalho da tabela
            for col in range(1, 4):  # col 1 a 3 (A a C)
                cell = ws.cell(row=15, column=col)
                cell.font = fonte_negrito
                cell.alignment = alinhamento_centro
                cell.border = borda

            # Dados da tabela
            for row in ws.iter_rows(min_row=15, max_row=ws.max_row, min_col=1, max_col=3):
                for cell in row:
                    if cell.value is not None:
                        cell.font = fonte
                        cell.alignment = alinhamento_centro
                        cell.border = borda

            ultima_linha = ws.max_row + 2

            # Rodapé
 # --- Rodapé com mês +1 ---
            meses = [
                "janeiro","fevereiro","março","abril","maio","junho",
                "julho","agosto","setembro","outubro","novembro","dezembro"
            ]

            indice_mes = meses.index(mes)  # pega índice do mês selecionado
            mes_rodape = meses[(indice_mes + 1) % 12]  # próximo mês, volta para janeiro se dezembro

            ultima_linha = ws.max_row + 2
            ws.merge_cells(start_row=ultima_linha, start_column=1, end_row=ultima_linha, end_column=3)
            ws.cell(row=ultima_linha, column=1).value = f"São Paulo, 01 de {mes_rodape} de {ano}"
            ws.cell(row=ultima_linha, column=1).font = fonte_negrito
            ws.cell(row=ultima_linha, column=1).alignment = alinhamento_centro


            # Assinaturas
            linha_assinatura = ultima_linha + 3
            # Fiscal de Contrato (à esquerda)
            ws.merge_cells(start_row=linha_assinatura, start_column=1, end_row=linha_assinatura, end_column=2)
            ws.cell(row=linha_assinatura, column=1).value = "________________________"
            alinhamento_esquerda = Alignment(horizontal="left", vertical="center")
            ws.cell(row=linha_assinatura, column=1).alignment = alinhamento_esquerda

            # Supervisão (mantém à direita)
            ws.merge_cells(start_row=linha_assinatura, start_column=3, end_row=linha_assinatura, end_column=4)
            ws.cell(row=linha_assinatura, column=3).value = "________________________"
            ws.cell(row=linha_assinatura, column=3).alignment = alinhamento_centro

            # Títulos abaixo das assinaturas
            ws.cell(row=linha_assinatura + 1, column=1).value = "Fiscal de Contrato"
            ws.cell(row=linha_assinatura + 1, column=1).font = fonte_negrito
            ws.cell(row=linha_assinatura + 1, column=1).alignment = alinhamento_esquerda

            ws.cell(row=linha_assinatura + 1, column=3).value = "Supervisão"
            ws.cell(row=linha_assinatura + 1, column=3).font = fonte_negrito
            ws.cell(row=linha_assinatura + 1, column=3).alignment = alinhamento_centro

        wb.save(saida)
        messagebox.showinfo("Sucesso", f"Relatório salvo em:\n{saida}")

    except Exception as e:
        messagebox.showerror("Erro", str(e))


# --- INTERFACE TKINTER ---
root = tk.Tk()
root.title("📊 Gerador de Relatório")
root.geometry("500x350")
root.configure(bg="#f4f4f4")

style = ttk.Style()
style.theme_use("clam")
style.configure("TButton", font=("Segoe UI", 12, "bold"), padding=10,
                background="#4CAF50", foreground="white")
style.map("TButton", background=[("active", "#45a049")])

frame = ttk.Frame(root, padding=20)
frame.pack(expand=True)

titulo = tk.Label(frame, text="Gerador de Relatório por Posto",
                  font=("Segoe UI", 16, "bold"), bg="#f4f4f4", fg="#333")
titulo.pack(pady=10)

frame_data = ttk.Frame(frame)
frame_data.pack(pady=5)

tk.Label(frame_data, text="Mês:", font=("Segoe UI", 11), background="#f4f4f4").grid(row=0, column=0, padx=5)
combo_mes = ttk.Combobox(frame_data, values=[
    "janeiro","fevereiro","março","abril","maio","junho",
    "julho","agosto","setembro","outubro","novembro","dezembro"
], width=12)
combo_mes.grid(row=0, column=1, padx=5)

tk.Label(frame_data, text="Ano:", font=("Segoe UI", 11), background="#f4f4f4").grid(row=0, column=2, padx=5)
entry_ano = tk.Entry(frame_data, width=8)
entry_ano.grid(row=0, column=3, padx=5)

btn = ttk.Button(frame, text="📂 Selecionar e Gerar Relatório", command=gerar_relatorio)
btn.pack(pady=20)

root.mainloop()
