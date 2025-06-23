from docx2pdf import convert

arquivos = [
    r"C:\Users\lealf\OneDrive\Área de trabalho - atual\ARQUIVOS ESCRITÓRIO DE EMANUEL\SEPARAÇÃO CONCENSUAL\REGINALDO ANTAS\PROCURACAO DOCS E PETICAO INICIAL\PROCURACAO REGINALDO ANTAS.docx"

]

for arquivo in arquivos:
    convert(arquivo)
    print(f"{arquivo} convertido com sucesso!")

print("Todos os arquivos foram convertidos!")


