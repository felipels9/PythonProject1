import os
import subprocess
from pypdf import PdfReader, PdfWriter

def unir_pdfs(lista_arquivos, pdf_saida):
    pdf_writer = PdfWriter()
    for pdf in lista_arquivos:
        leitor = PdfReader(pdf)
        for pagina in leitor.pages:
            pdf_writer.add_page(pagina)
    with open(pdf_saida, 'wb') as saida:
        pdf_writer.write(saida)
    print(f"✅ PDFs unidos em: {pdf_saida}")

def reduzir_pdf(input_pdf, output_pdf, gs_path, tamanho_max_mb=5):
    qualidade = "/ebook"

    # Criar pasta temporária em C:\temp_gs
    pasta_temp = r"C:\temp_gs"
    try:
        os.makedirs(pasta_temp, exist_ok=True)
        print(f"✅ Pasta temporária criada ou já existe: {pasta_temp}")
    except Exception as e:
        print(f"❌ Erro ao criar pasta temporária: {e}")
        return

    nome_temp = os.path.basename(output_pdf).replace(".pdf", "_temp.pdf")
    temp_output = os.path.join(pasta_temp, nome_temp)

    comandos = [
        gs_path,
        "-sDEVICE=pdfwrite",
        "-dCompatibilityLevel=1.4",
        f"-dPDFSETTINGS={qualidade}",
        "-dNOPAUSE",
        "-dBATCH",
        "-dDownsampleColorImages=true",
        "-dColorImageDownsampleType=/Bicubic",
        "-dColorImageResolution=110",
        "-dDownsampleGrayImages=true",
        "-dGrayImageDownsampleType=/Bicubic",
        "-dGrayImageResolution=110",
        "-dDownsampleMonoImages=true",
        "-dMonoImageDownsampleType=/Subsample",
        "-dMonoImageResolution=110",
        f"-sOutputFile={temp_output}",
        input_pdf
    ]

    print(f"📌 Comando Ghostscript:\n{' '.join(comandos)}")
    print(f"🔧 Executando Ghostscript, arquivo temporário: {temp_output}")

    try:
        resultado = subprocess.run(comandos, check=True, capture_output=True, text=True)
        print("✅ Ghostscript executado com sucesso.")
        print("📝 Saída do Ghostscript:")
        print(resultado.stdout)
        print("🛑 Erros do Ghostscript (se houver):")
        print(resultado.stderr)
    except subprocess.CalledProcessError as e:
        print(f"❌ Erro ao executar Ghostscript:\n{e.stderr}")
        return

    if not os.path.isfile(temp_output):
        print("❌ O arquivo comprimido NÃO foi gerado. Verifique erros acima.")
        return

    tamanho_bytes = os.path.getsize(temp_output)
    tamanho_mb = tamanho_bytes / (1024 * 1024)

    if tamanho_mb > tamanho_max_mb:
        print(f"⚠️ Arquivo ainda tem {tamanho_mb:.2f}MB, acima do limite de {tamanho_max_mb}MB.")
    else:
        print(f"✅ Arquivo reduzido para {tamanho_mb:.2f}MB e está dentro do limite.")

    try:
        os.replace(temp_output, output_pdf)
        print(f"📄 Arquivo final salvo em: {output_pdf}")
    except Exception as e:
        print(f"❌ Erro ao mover arquivo temporário para destino final: {e}")

# Caminho Ghostscript 32 bits - verifique se está correto!
gs_path = r"C:\Program Files (x86)\gs\gs10.05.1\bin\gswin32c.exe"

lista_pdfs = [
    r"C:\Users\lealf\OneDrive\Área de trabalho - atual\ARQUIVOS ESCRITÓRIO DE EMANUEL\PEDRO FERRAZ DA SILVA\DOCS E PROCURACAO\img20250623_10080013.pdf",
    r"C:\Users\lealf\OneDrive\Área de trabalho - atual\ARQUIVOS ESCRITÓRIO DE EMANUEL\PEDRO FERRAZ DA SILVA\DOCS E PROCURACAO\img20250623_10085154.pdf",
    r"C:\Users\lealf\OneDrive\Área de trabalho - atual\ARQUIVOS ESCRITÓRIO DE EMANUEL\PEDRO FERRAZ DA SILVA\DOCS E PROCURACAO\img20250623_10094637.pdf",
    r"C:\Users\lealf\OneDrive\Área de trabalho - atual\ARQUIVOS ESCRITÓRIO DE EMANUEL\PEDRO FERRAZ DA SILVA\DOCS E PROCURACAO\img20250623_10103832.pdf",
    r"C:\Users\lealf\OneDrive\Área de trabalho - atual\ARQUIVOS ESCRITÓRIO DE EMANUEL\PEDRO FERRAZ DA SILVA\DOCS E PROCURACAO\img20250623_10113608.pdf",
    r"C:\Users\lealf\OneDrive\Área de trabalho - atual\ARQUIVOS ESCRITÓRIO DE EMANUEL\PEDRO FERRAZ DA SILVA\DOCS E PROCURACAO\img20250623_10123578.pdf",
    r"C:\Users\lealf\OneDrive\Área de trabalho - atual\ARQUIVOS ESCRITÓRIO DE EMANUEL\PEDRO FERRAZ DA SILVA\DOCS E PROCURACAO\img20250623_10133281.pdf",
    r"C:\Users\lealf\OneDrive\Área de trabalho - atual\ARQUIVOS ESCRITÓRIO DE EMANUEL\PEDRO FERRAZ DA SILVA\DOCS E PROCURACAO\img20250623_10141848.pdf",
    r"C:\Users\lealf\OneDrive\Área de trabalho - atual\ARQUIVOS ESCRITÓRIO DE EMANUEL\PEDRO FERRAZ DA SILVA\DOCS E PROCURACAO\img20250623_10151031.pdf",
    r"C:\Users\lealf\OneDrive\Área de trabalho - atual\ARQUIVOS ESCRITÓRIO DE EMANUEL\PEDRO FERRAZ DA SILVA\DOCS E PROCURACAO\img20250623_10162929.pdf"
]

pdf_unido = r"C:\Users\lealf\OneDrive\Área de trabalho - atual\ARQUIVOS ESCRITÓRIO DE EMANUEL\PEDRO FERRAZ DA SILVA\DOCS E PROCURACAO\unid.pdf"
pdf_final = r"C:\Users\lealf\OneDrive\Área de trabalho - atual\ARQUIVOS ESCRITÓRIO DE EMANUEL\PEDRO FERRAZ DA SILVA\DOCS E PROCURACAO\comprim.pdf"

unir_pdfs(lista_pdfs, pdf_unido)
reduzir_pdf(pdf_unido, pdf_final, gs_path)
