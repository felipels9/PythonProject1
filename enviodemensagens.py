import pandas as pd
import pyautogui
import webbrowser
import time
import urllib

# Caminho do arquivo PDF do convite
arquivo_pdf = r"Venha celebrar conosco!.pdf"

# Ler a planilha Excel com convidados e contatos
tabela = pd.read_excel(r"C:\Users\lealf\OneDrive\Área de trabalho - atual\CONVIDADOS.xlsx")

for i, linha in tabela.iterrows():
    nome = str(linha['Nome']).strip()
    telefone = str(linha['Contato']).strip()

    if not telefone.startswith("55"):
        telefone = "55" + telefone

    # Mensagem personalizada
    if (' e ' in nome.lower()) or (' & ' in nome.lower()):
        mensagem = f"""Para: {nome} 

Está chegando o momento mais esperado do ano: o nosso casamento! ❤️

Estamos muito felizes e contamos com vossas presenças em nosso grande dia!

O ícone do convite é clicável e te levará diretamente para:

✅ Local da cerimônia e da recepção;  
✅ Nossa lista de presentes (opcional);  
✅ Confirmação de presença.

✨ É muito importante que vocês confirmem presença até o dia 05/12/2025.

Para isso, siga este passo a passo:  

1) Abra o pdf e clique no ícone de carta no convite. 
2) Role até o final da página, na seção “Confirme sua presença”.  
3) Preencha com seu nome.  
4) Clique no convite para selecionar e digite a senha: 0298.  
5) Finalize sua confirmação e pronto! 🎉

Se desejarem nos presentear, fiquem à vontade para escolher algum dos itens da "Lista de Presentes" ou qualquer outra forma que lhes convier.

A compra é segura e pode ser feita por cartão de crédito (parcelado), Pix ou boleto bancário. 😊

📝 Atenção: No resumo da compra aparece um custo referente ao "Cartão Postal", que é opcional. Você pode removê-lo clicando em “Não gostaria de enviar um cartão”, pagando apenas pelo presente escolhido.

Aguardamos vocês no nosso grande dia! 💖
"""
    else:
        mensagem = f"""Para: {nome} 

Está chegando o momento mais esperado do ano: o nosso casamento! ❤️

Estamos muito felizes e contamos com sua presença em nosso grande dia!

O ícone do convite é clicável e te levará diretamente para:

✅ Local da cerimônia e da recepção;  
✅ Nossa lista de presentes (opcional);  
✅ Confirmação de presença.

✨ É muito importante que você confirme sua presença até o dia 05/12/2025.

Para isso, siga este passo a passo:  

1) Abra o pdf e clique no ícone de carta no convite.  
2) Role até o final da página, na seção “Confirme sua presença”.  
3) Preencha com seu nome.  
4) Clique no convite para selecionar e digite a senha: 0298.  
5) Finalize sua confirmação e pronto! 🎉

Se desejar nos presentear, fique à vontade para escolher algum dos itens da "Lista de Presentes" ou qualquer outra forma que lhe convier.

A compra é segura e pode ser feita por cartão de crédito (parcelado), Pix ou boleto bancário. 😊

📝 Atenção: No resumo da compra aparece um custo referente ao "Cartão Postal", que é opcional. Você pode removê-lo clicando em “Não gostaria de enviar um cartão”, pagando apenas pelo presente escolhido.

Aguardamos você no nosso grande dia! 💖
"""

    # Codifica a mensagem para URL
    link = f"https://web.whatsapp.com/send?phone={telefone}&text={urllib.parse.quote(mensagem)}"

    webbrowser.open(link)
    time.sleep(15)  # Espera carregar o chat com a mensagem já no campo

    # Envia a mensagem (texto) que está no campo de texto
    pyautogui.press('enter')
    time.sleep(5)

    # Anexar o arquivo PDF (sem digitar o caminho no chat!)
    pyautogui.click(x=709, y=966)  # Ícone de anexar (clipe)
    time.sleep(2)

    pyautogui.click(x=686, y=536)  # Opção Documento
    time.sleep(2)

    # Digita o caminho do arquivo PDF direto na janela do arquivo (dialog do SO)
    pyautogui.write(arquivo_pdf)
    time.sleep(1)
    pyautogui.press('enter')  # Confirmar seleção do arquivo
    time.sleep(5)

    # Envia o arquivo PDF
    pyautogui.press('enter')
    time.sleep(8)

    # Fecha a aba do navegador
    pyautogui.hotkey('ctrl', 'w')
    time.sleep(3)

print("✅ Mensagens e convites enviados com sucesso!")
