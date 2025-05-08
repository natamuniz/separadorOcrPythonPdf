import os
import re
import fitz
import pytesseract
from PIL import Image
from pdf2image import convert_from_path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path

# Caminhos do Tesseract e Poppler
pytesseract.pytesseract.tesseract_cmd = r'C:\Users\USUARIO\Desktop\Separador 2\Tesseract-OCR\tesseract.exe'
poppler_path = 'C:\\Users\\USUARIO\\Desktop\\Separador 2\\poppler-24.08.0\\Library\\bin'

def extrair_nome_paciente(imagem):
    try:
        texto = pytesseract.image_to_string(imagem, lang='por')
        # Remove quebras de linha e espaços duplicados
        texto = texto.replace('\n', ' ').strip()

        match = re.search(r"(?:Nome|Paciente)[\s:\-\"“”|]*([A-ZÀ-Ú][A-Za-zÀ-ÿ\s]{3,})", texto, re.IGNORECASE)
        if match:
            nome = match.group(1).strip()
            nome = re.sub(r'\s+', ' ', nome)  # remove espaços repetidos
            nome = re.sub(r'[^A-Za-zÀ-ÿ\s]', '', nome)  # remove números e símbolos
            if len(nome.split()) >= 2:
                return nome
    except Exception:
        pass
    return "Paciente_Desconhecido"


def processar_pdf(pdf_path, saida_pasta_base):
    nome_pdf = Path(pdf_path).stem
    saida_path = os.path.join(saida_pasta_base, nome_pdf)
    os.makedirs(saida_path, exist_ok=True)

    doc = fitz.open(pdf_path)
    grupos = []
    grupo_atual = []
    nome_atual = None

    for i in range(len(doc)):
        imagem = convert_from_path(
            pdf_path,
            dpi=150,
            first_page=i + 1,
            last_page=i + 1,
            poppler_path=poppler_path
        )[0]

        nome_encontrado = extrair_nome_paciente(imagem)
        print(f"[Página {i + 1}] Nome detectado: {nome_encontrado}")

        if nome_encontrado == "Paciente_Desconhecido" or nome_encontrado == nome_atual:
            grupo_atual.append(i)
        else:
            if grupo_atual:
                grupos.append((nome_atual, grupo_atual))
            grupo_atual = [i]
            nome_atual = nome_encontrado

    if grupo_atual:
        grupos.append((nome_atual, grupo_atual))

    for nome, paginas in grupos:
        nome_final = nome or "Paciente_Desconhecido"
        nome_limpo = re.sub(r'[\\/*?:"<>|\n\r]', "_", nome_final)[:100]
        contador = 1
        caminho_saida = os.path.join(saida_path, f"{nome_limpo}.pdf")
        while os.path.exists(caminho_saida):
            caminho_saida = os.path.join(saida_path, f"{nome_limpo}_{contador}.pdf")
            contador += 1

        novo_pdf = fitz.open()
        for num in paginas:
            novo_pdf.insert_pdf(doc, from_page=num, to_page=num)
        novo_pdf.save(caminho_saida)
        novo_pdf.close()

    doc.close()

def iniciar_interface():
    def selecionar_pasta_pdfs():
        pasta = filedialog.askdirectory(title="Selecione a pasta com PDFs")
        if pasta:
            entry_pdf_dir.delete(0, tk.END)
            entry_pdf_dir.insert(0, pasta)

    def selecionar_pasta_saida():
        pasta = filedialog.askdirectory(title="Selecione a pasta de saída")
        if pasta:
            entry_saida.delete(0, tk.END)
            entry_saida.insert(0, pasta)

    def iniciar_processamento():
        pasta_pdfs = entry_pdf_dir.get()
        pasta_saida = entry_saida.get()

        if not pasta_pdfs or not pasta_saida:
            messagebox.showerror("Erro", "Selecione ambas as pastas.")
            return

        pdfs = [os.path.join(pasta_pdfs, f) for f in os.listdir(pasta_pdfs) if f.lower().endswith(".pdf")]
        if not pdfs:
            messagebox.showinfo("Aviso", "Nenhum PDF encontrado na pasta.")
            return

        progress["maximum"] = len(pdfs)
        progress["value"] = 0
        root.update_idletasks()

        btn_iniciar.config(state="disabled")

        for i, pdf in enumerate(pdfs, 1):
            processar_pdf(pdf, pasta_saida)
            progress["value"] = i
            root.update_idletasks()

        btn_iniciar.config(state="normal")
        messagebox.showinfo("Concluído", "Todos os PDFs foram processados com sucesso.")

    root = tk.Tk()
    root.title("Separador de Prontuários")
    root.geometry("500x250")

    tk.Label(root, text="Pasta com PDFs:").pack(pady=(10, 0))
    entry_pdf_dir = tk.Entry(root, width=60)
    entry_pdf_dir.pack()
    tk.Button(root, text="Selecionar", command=selecionar_pasta_pdfs).pack(pady=(0, 10))

    tk.Label(root, text="Pasta de saída:").pack()
    entry_saida = tk.Entry(root, width=60)
    entry_saida.pack()
    tk.Button(root, text="Selecionar", command=selecionar_pasta_saida).pack(pady=(0, 10))

    global btn_iniciar
    btn_iniciar = tk.Button(root, text="Iniciar", command=iniciar_processamento)
    btn_iniciar.pack(pady=(10, 0))

    global progress
    progress = ttk.Progressbar(root, orient="horizontal", length=400, mode="determinate")
    progress.pack(pady=(10, 0))

    root.mainloop()

if __name__ == "__main__":
    iniciar_interface()
