import os
import subprocess
import shutil
import tempfile
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import PhotoImage
from PyPDF2 import PdfReader, PdfWriter

def unir_pdfs(lista_arquivos, pdf_saida):
    pdf_writer = PdfWriter()
    arquivos_pdf = [f for f in lista_arquivos if f.lower().endswith(".pdf")]

    if not arquivos_pdf:
        messagebox.showerror("Erro", "Nenhum PDF selecionado.")
        return False

    for caminho_pdf in arquivos_pdf:
        try:
            leitor = PdfReader(caminho_pdf)
            for pagina in leitor.pages:
                pdf_writer.add_page(pagina)
        except Exception as e:
            print(f"Erro ao ler {caminho_pdf}: {e}")

    try:
        with open(pdf_saida, 'wb') as saida:
            pdf_writer.write(saida)
        return True
    except Exception as e:
        messagebox.showerror("Erro ao salvar PDF unido", str(e))
        return False

def reduzir_pdf(input_pdf, output_pdf, tamanho_max_mb=5):
    qualidade = "/ebook"
    pasta_temp = tempfile.mkdtemp(prefix="temp_gs_")
    nome_temp = os.path.basename(output_pdf).replace(".pdf", "_temp.pdf")
    temp_output = os.path.join(pasta_temp, nome_temp)

    gs_path = None
    possiveis_caminhos = [
        r"C:\Program Files\gs\gs10.05.1\bin\gswin64c.exe",
        r"C:\Program Files (x86)\gs\gs10.05.1\bin\gswin32c.exe"
    ]
    for caminho in possiveis_caminhos:
        if os.path.exists(caminho):
            gs_path = caminho
            break

    if not gs_path:
        messagebox.showerror("Erro", "Ghostscript não encontrado.")
        return

    comandos = [
        f'"{gs_path}"',
        "-sDEVICE=pdfwrite",
        "-dCompatibilityLevel=1.4",
        f"-dPDFSETTINGS={qualidade}",
        "-dNOPAUSE",
        "-dBATCH",
        "-dQUIET",
        "-dDownsampleColorImages=true",
        "-dColorImageDownsampleType=/Bicubic",
        "-dColorImageResolution=110",
        "-dDownsampleGrayImages=true",
        "-dGrayImageDownsampleType=/Bicubic",
        "-dGrayImageResolution=110",
        "-dDownsampleMonoImages=true",
        "-dMonoImageDownsampleType=/Subsample",
        "-dMonoImageResolution=110",
        f'-sOutputFile="{temp_output}"',
        f'"{input_pdf}"'
    ]

    try:
        subprocess.run(" ".join(comandos), shell=True, check=True)

        if not os.path.isfile(temp_output):
            messagebox.showerror("Erro", "Arquivo comprimido não foi gerado.")
            return

        os.replace(temp_output, output_pdf)

        tamanho_mb = os.path.getsize(output_pdf) / (1024 * 1024)
        msg = f"✅ PDF comprimido salvo com {tamanho_mb:.2f} MB\n\n{output_pdf}"
        messagebox.showinfo("Sucesso", msg)

    except subprocess.CalledProcessError as e:
        messagebox.showerror("Erro Ghostscript", e.stderr)

    finally:
        try:
            shutil.rmtree(pasta_temp)
        except Exception as e:
            print(f"⚠️ Falha ao remover pasta temporária: {e}")

def selecionar_arquivos():
    arquivos = filedialog.askopenfilenames(
        title="Selecione os arquivos PDF",
        filetypes=[("Arquivos PDF", "*.pdf")]
    )

    if not arquivos:
        return

    destino = filedialog.asksaveasfilename(
        defaultextension=".pdf",
        filetypes=[("PDF", "*.pdf")],
        title="Salvar PDF unido e comprimido como"
    )

    if not destino:
        return

    pdf_unido_temp = os.path.join(tempfile.gettempdir(), "temp_pdf_unido.pdf")

    sucesso = unir_pdfs(arquivos, pdf_unido_temp)
    if sucesso:
        reduzir_pdf(pdf_unido_temp, destino)

# === Interface Gráfica ===
root = tk.Tk()
root.title("Compressor de PDFs")
root.geometry("420x220")
root.resizable(False, False)

# Centralizar janela na tela
root.update_idletasks()
w = 420
h = 220
x = (root.winfo_screenwidth() // 2) - (w // 2)
y = (root.winfo_screenheight() // 2) - (h // 2)
root.geometry(f"{w}x{h}+{x}+{y}")

# Definir ícone da janela
try:
    root.iconbitmap("OIP.ico")
except:
    pass  # Ignorar se o ícone falhar

frame = tk.Frame(root, padx=20, pady=20)
frame.pack()

label = tk.Label(frame, text="Selecione os arquivos PDF para unir e comprimir:")
label.pack(pady=(0, 15))

botao = tk.Button(frame, text="Selecionar PDFs", command=selecionar_arquivos, height=2, width=30)
botao.pack()

sair = tk.Button(frame, text="Sair", command=root.destroy)
sair.pack(pady=(20, 0))

root.mainloop()
