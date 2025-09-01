import os
import re
import shutil
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
from PIL import Image, ImageEnhance, ImageFilter
import fitz
import pytesseract

# Configuração
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
OCR_CONFIG = r'--oem 3 --psm 4'

EMPRESAS_DISPONIVEIS = {'Works': '02', 'Qualitech': '01', 'Partner': '02', 'Presseg': '01'}
TIPO_DOCS = {'Folha de Ponto': '01', 'FT': '84'}
ANO = {str(y): str(y) for y in range(2020, 2026)}
MES = {'Janeiro': '01', 'Fevereiro': '02', 'Março': '03', 'Abril': '04',
       'Maio': '05', 'Junho': '06', 'Julho': '07', 'Agosto': '08',
       'Setembro': '09', 'Outubro': '10', 'Novembro': '11', 'Dezembro': '12'}

parar_processamento = False

def pdf_para_imagens(caminho_pdf):
    imagens = []
    try:
        doc = fitz.open(caminho_pdf)
        for page in doc:
            pix = page.get_pixmap(dpi=300)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            imagens.append(img)
    except Exception as e:
        print(f"[ERRO ABRIR PDF] {caminho_pdf}: {e}")
    return imagens

def preprocessar_imagem(img):
    try:
        img_gray = img.convert("L")
        osd = pytesseract.image_to_osd(img_gray)
        rotate_match = re.search(r'Rotate: (\d+)', osd)
        if rotate_match:
            angle = int(rotate_match.group(1))
            if angle != 0:
                img = img.rotate(-angle, expand=True)
    except Exception as e:
        print(f"[AVISO] Falha ao detectar rotação automática: {e}")

    img = img.convert("L")
    img = img.filter(ImageFilter.MedianFilter(size=3))
    enhancer = ImageEnhance.Contrast(img)
    img = enhancer.enhance(2.5)
    return img

def extrair_re(texto):
    print("---- EXTRAINDO RE ----")

    # 1. Tenta por 'Registro'
    match = re.search(r'Registro\s*[:\-]?\s*(\d{3,6})', texto, re.IGNORECASE)
    if not match:
        match = re.search(r'\b(?:registro|reg)\s*[:\-]?\s*(\d{5,})', texto, re.IGNORECASE)
    if match:
        return match.group(1).zfill(5), False

    # 2. Rodapé com " - Nome"
    linhas = texto.strip().splitlines()
    for linha in reversed(linhas):
        if " - Nome" in linha:
            tentativa = re.search(r'(\d{3,6})\s*-\s*Nome', linha)
            if tentativa:
                return tentativa.group(1).zfill(5), False

    # 3. Tenta por 'Crachá'
    cracha_match = re.search(r'crach[aá][: ]+\s*(\d{3,6})', texto, re.IGNORECASE)
    if cracha_match:
        return cracha_match.group(1).zfill(5), True

    return "00000", False

def limpar_nome(nome):
    return re.sub(r'[<>:"/\\|?*\n\r\t]', '_', nome)



def gerar_nome_unico(nome_base, destino):
    novo_nome = os.path.join(destino, f"{nome_base}.pdf")
    contador = 1
    while os.path.exists(novo_nome):
        novo_nome = os.path.join(destino, f"{nome_base}_{contador}.pdf")
        contador += 1
    return novo_nome

def renomear_pdfs(pasta_entrada, pasta_saida, empresa_codigo, mes_codigo, ano_completo, tipodoc_nome):
    global parar_processamento
    parar_processamento = False

    if not os.path.exists(pasta_saida):
        os.makedirs(pasta_saida)

    if tipodoc_nome == "FT":
        empresa_nome = combo_empresa.get()
        tipodoc = "90" if empresa_nome in ("Qualitech", "Partner") else "84"
    else:
        tipodoc = TIPO_DOCS.get(tipodoc_nome, "01")

    for nome_arquivo in os.listdir(pasta_entrada):
        if parar_processamento:
            print("[PROCESSO INTERROMPIDO PELO USUÁRIO]")
            messagebox.showwarning("Interrompido", "Renomeação interrompida pelo usuário.")
            break

        if nome_arquivo.lower().endswith(".pdf"):
            caminho_pdf = os.path.join(pasta_entrada, nome_arquivo)
            imagens = pdf_para_imagens(caminho_pdf)

            if not imagens:
                print(f"[IGNORADO] {nome_arquivo} não pôde ser processado.")
                continue

            texto_completo = ""
            for img in imagens:
                img_pre = preprocessar_imagem(img)
                texto = pytesseract.image_to_string(img_pre, lang='por+eng', config=OCR_CONFIG)
                texto_completo += texto

            registro, veio_do_cracha = extrair_re(texto_completo)
            print(f"Arquivo: {nome_arquivo} ➜ RE original: {registro}")

            if not veio_do_cracha and len(registro) > 1:
                registro = registro[:-1].zfill(5)

            print(f"➜ RE ajustado: {registro}")

            paginas = str(len(imagens)).zfill(2)
            codigo = f"{empresa_codigo}{mes_codigo}{ano_completo}{paginas}{tipodoc}{registro}"
            codigo_limpo = limpar_nome(codigo)
            novo_nome = gerar_nome_unico(codigo_limpo, pasta_saida)

            try:
                shutil.move(caminho_pdf, novo_nome)
                print(f"Renomeado para: {novo_nome}")
            except Exception as e:
                print(f"[ERRO RENOMEAR] {e}")

    if not parar_processamento:
        messagebox.showinfo("Concluído", "Renomeação finalizada com sucesso.")

def selecionar_pasta():
    return filedialog.askdirectory(title="Selecione a pasta com PDFs")

def selecionar_pasta_destino():
    return filedialog.askdirectory(title="Selecione a pasta DESTINO para os PDFs renomeados")

def iniciar_script():
    empresa_nome = combo_empresa.get()
    tipodoc_nome = combo_tipodoc.get()
    ano_interface = combo_ano.get()
    mes_nome = combo_mes.get()

    if not (empresa_nome and mes_nome and ano_interface and tipodoc_nome):
        messagebox.showwarning("Atenção", "Preencha todos os campos obrigatórios.")
        return

    empresa_codigo = EMPRESAS_DISPONIVEIS.get(empresa_nome, "0")
    mes_codigo = MES.get(mes_nome, "01")
    ano_completo = ANO.get(ano_interface, "2024")

    pasta_entrada = selecionar_pasta()
    if not pasta_entrada:
        return

    pasta_saida = selecionar_pasta_destino()
    if not pasta_saida:
        return

    renomear_pdfs(pasta_entrada, pasta_saida, empresa_codigo, mes_codigo, ano_completo, tipodoc_nome)

def parar_execucao():
    global parar_processamento
    parar_processamento = True

# Interface
janela = tk.Tk()
janela.title("Codificador de PDFs")
janela.geometry("330x230")
janela.configure(bg="#FF4141")

fonte_padrao = ("Arial", 11)
bg_padrao = "#FF4141"

style = ttk.Style()
style.theme_use("default")
style.configure("TCombobox", fieldbackground="white", background="#e0e0e0", font=fonte_padrao)
style.map("TCombobox", fieldbackground=[('readonly', 'white')])

tk.Label(janela, text="Empresa:", bg=bg_padrao, font=fonte_padrao).grid(row=0, column=0, padx=5, pady=5, sticky="w")
combo_empresa = ttk.Combobox(janela, values=list(EMPRESAS_DISPONIVEIS.keys()), state="readonly", font=fonte_padrao)
combo_empresa.grid(row=0, column=1)
combo_empresa.set(list(EMPRESAS_DISPONIVEIS.keys())[0])

tk.Label(janela, text="Tipo de Documento:", bg=bg_padrao, font=fonte_padrao).grid(row=1, column=0, padx=5, pady=5, sticky="w")
combo_tipodoc = ttk.Combobox(janela, values=list(TIPO_DOCS.keys()), state="readonly", font=fonte_padrao)
combo_tipodoc.grid(row=1, column=1)
combo_tipodoc.set(list(TIPO_DOCS.keys())[0])

tk.Label(janela, text="Ano:", bg=bg_padrao, font=fonte_padrao).grid(row=2, column=0, padx=5, pady=5, sticky="w")
combo_ano = ttk.Combobox(janela, values=list(ANO.keys()), state="readonly", font=fonte_padrao)
combo_ano.grid(row=2, column=1)
combo_ano.set("2024")

tk.Label(janela, text="Mês:", bg=bg_padrao, font=fonte_padrao).grid(row=3, column=0, padx=5, pady=5, sticky="w")
combo_mes = ttk.Combobox(janela, values=list(MES.keys()), state="readonly", font=fonte_padrao)
combo_mes.grid(row=3, column=1)
combo_mes.set(list(MES.keys())[0])

tk.Button(janela, text="Selecionar Pastas e Iniciar", command=iniciar_script, font=fonte_padrao, bg="#ffffff").grid(row=4, column=0, columnspan=2, pady=10)
tk.Button(janela, text="Parar Execução", command=parar_execucao, font=fonte_padrao, bg="#ffcccc").grid(row=5, column=0, columnspan=2, pady=5)

janela.mainloop()