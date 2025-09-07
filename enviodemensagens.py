import os
import re
import time
import random
import urllib
import pandas as pd
import pyautogui
import webbrowser
import pyperclip
import logging
from datetime import datetime

# ======= REPORTLAB (PDF) =======
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm

# ======= CONFIGURAÇÕES =======

# Caminho do arquivo PDF do convite
ARQUIVO_PDF = r"C:\Users\lealf\OneDrive\Área de trabalho - atual\FELIPE LEAL\CONVITE DE CASAMENTO\venhacelebrarconosco.pdf"

# Caminho da planilha Excel (colunas obrigatórias: Nome, Contato)
EXCEL_PATH = r"C:\Users\lealf\OneDrive\Área de trabalho - atual\FELIPE LEAL\CONVITE DE CASAMENTO\teste.xlsx"

# Saídas

NOME_PDF_RELATORIO = "relatorio_envio_whatsapp.pdf"
NOME_CSV_RELATORIO = "relatorio_envio_whatsapp.csv"   # histórico (habilite via flag)
NOME_LOG = None  # se None, cria automático

# Modo de operação
MODO_SOMENTE_PDF_DE_CSV = False   # True: não envia nada; lê CSV histórico e gera PDF
CSV_RESULTADOS_FONTE = r""        # caminho do CSV a usar no modo acima (se vazio, usa CSV padrão ao lado do Excel)

# Comportamento de envio

ABRIR_NAVEGADOR = True
FECHAR_ABA_APOS_ENVIO = True
ENVIAR_ANEXO = True
DRY_RUN = False                    # True = simula (não clica, não envia)

# CSV histórico desativado por padrão
GERAR_CSV_HISTORICO = False

CALIBRAR_COORDENADAS = False       # True = roda rotina de calibração antes da execução
CHECAR_LOGIN_INICIAL = True        # abre web.whatsapp.com e espera para login

# Tempos base (segundos)
TEMPO_CARREGAR_CHAT = 22
TEMPO_APOS_ENVIAR_TEXTO = 3
TEMPO_APOS_CLIP = 2
TEMPO_APOS_DOC = 2
TEMPO_APOS_COLAR_CAMINHO = 1
TEMPO_UPLOAD = 5
TEMPO_APOS_ENVIAR_ARQUIVO = 8
TEMPO_APOS_FECHAR_ABA = 2
INTERVALO_ENTRE_CONTATOS = (1.5, 4.0)

# Variação aleatória (jitter)
VARIACAO = {'carregar': 4, 'apos_texto': 1, 'apos_clip': 1, 'apos_doc': 1, 'upload': 3, 'apos_arquivo': 2, 'final': 1}

# Coordenadas padrão da sua tela (clipe e opção Documento) — serão sobrescritas se calibrar
CLIP_X, CLIP_Y = 709, 966
DOC_X, DOC_Y   = 686, 536

# Segurança PyAutoGUI
pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.5


# ======= LOGGING (silenciado) =======
def preparar_logging(pasta_relatorios: str):
    """
    Configurado para silenciar logs (apenas CRITICAL).
    """
    logging.basicConfig(
        level=logging.CRITICAL,  # <<< silencia INFO/ERROR/EXCEPTION
        format="%(asctime)s [%(levelname)s] %(message)s"
    )
    return "console"


# ======= UTILS =======
def jitter(base: float, key: str) -> float:
    delta = VARIACAO.get(key, 0)
    return max(0, base + random.uniform(-delta, delta)) if delta > 0 else base

def somente_digitos(s: str) -> str:
    return re.sub(r"\D", "", (s or "").strip())

def normalizar_telefone(telefone: str) -> tuple[str, str]:
    d = somente_digitos(telefone)
    if not d:
        return "", "vazio"
    d = d.lstrip('0')
    if not d.startswith('55'):
        d = '55' + d
    sem_ddi = d[2:]
    if len(sem_ddi) < 10 or len(sem_ddi) > 11:
        return "", f"tamanho inválido após DDI (esperado 10-11, obtido {len(sem_ddi)})"
    ddd = sem_ddi[:2]
    numero = sem_ddi[2:]
    if not (len(ddd) == 2 and ddd.isdigit() and 11 <= int(ddd) <= 99):
        return "", f"DDD inválido: {ddd}"
    if len(numero) not in (8, 9):
        return "", f"número local inválido: {numero}"
    return d, ""

def montar_mensagem(nome: str) -> tuple[str, str]:
    # Plural heurístico: contém ' e ' ou ' & ' no nome
    plural = (' e ' in nome.lower()) or (' & ' in nome.lower())
    if plural:
        msg = f"""Para: {nome} 

Estamos muito felizes e contamos com a presença de vocês em nosso grande dia!

O ícone de carta no convite é clicável e te levará diretamente para um site contendo todas as informações sobre o casamento, como:

✅ Local da cerimônia e da recepção;  
✅ Nossa lista de presentes (opcional);  
✅ Confirmação de presença.

✨ É importante que vocês confirmem presença até o dia 05/12/2025.

Para isso, siga este passo a passo:  

1) Abra o pdf e clique no ícone de carta no convite. 
2) Role até o final da página, na seção “Confirme sua presença”.  
3) Preencha com seu nome e clique em "Pesquisar".  
4) Clique em "Selecionar" e digite a senha: 0298.  
5) Finalize sua confirmação clicando em "Validar" e pronto! 🎉

Se desejarem nos presentear, fiquem à vontade para escolher algum dos itens da "Lista de Presentes" ou qualquer outra forma que preferirem.

A compra é segura e pode ser feita por cartão de crédito (parcelado), Pix ou boleto bancário. 😊

📝 Atenção: No resumo da compra aparece um custo referente ao "Cartão Postal", que é opcional. Você pode removê-lo clicando em “Não gostaria de enviar um cartão”, pagando apenas pelo presente escolhido.

Aguardamos vocês no nosso grande dia! 💖
"""
        tipo = "Plural"
    else:
        msg = f"""Para: {nome} 

Estamos muito felizes e contamos com sua presença em nosso grande dia!

O ícone de carta no convite é clicável e te levará diretamente para um site contendo todas as informações sobre o casamento, como:

✅ Local da cerimônia e da recepção;  
✅ Nossa lista de presentes (opcional);  
✅ Confirmação de presença.

✨ É importante que você confirme sua presença até o dia 05/12/2025.

Para isso, siga este passo a passo:  

1) Abra o pdf e clique no ícone de carta no convite.  
2) Role até o final da página, na seção “Confirme sua presença”.  
3) Preencha com seu nome e clique em "Pesquisar".   
4) Clique em "Selecionar" e digite a senha: 0298.  
5) Finalize sua confirmação clicando em "Validar" e pronto! 🎉

Se desejar nos presentear, fique à vontade para escolher algum dos itens da "Lista de Presentes" ou qualquer outra forma que preferir.

A compra é segura e pode ser feita por cartão de crédito (parcelado), Pix ou boleto bancário. 😊

📝 Atenção: No resumo da compra aparece um custo referente ao "Cartão Postal", que é opcional. Você pode removê-lo clicando em “Não gostaria de enviar um cartão”, pagando apenas pelo presente escolhido.

Aguardamos você no nosso grande dia! 💖
"""
        tipo = "Singular"
    return msg, tipo

def abrir_chat(telefone_normalizado: str, mensagem: str):
    link = f"https://web.whatsapp.com/send?phone={telefone_normalizado}&text={urllib.parse.quote(mensagem)}"
    webbrowser.open(link)

def enviar_texto():
    pyautogui.press('enter')

def anexar_pdf(caminho_pdf: str):
    pyautogui.click(x=CLIP_X, y=CLIP_Y)
    time.sleep(jitter(TEMPO_APOS_CLIP, 'apos_clip'))
    pyautogui.click(x=DOC_X, y=DOC_Y)
    time.sleep(jitter(TEMPO_APOS_DOC, 'apos_doc'))
    pyperclip.copy(caminho_pdf)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(jitter(TEMPO_APOS_COLAR_CAMINHO, 'apos_doc'))
    pyautogui.press('enter')
    time.sleep(jitter(TEMPO_UPLOAD, 'upload'))
    pyautogui.press('enter')

def fechar_aba():
    pyautogui.hotkey('ctrl', 'w')

def checar_login_whatsapp():
    if DRY_RUN or not CHECAR_LOGIN_INICIAL or not ABRIR_NAVEGADOR:
        return
    webbrowser.open("https://web.whatsapp.com")
    time.sleep(12)

# ======= Calibração/Coordenadas =======
def caminho_coords() -> str:
    pasta = os.path.dirname(EXCEL_PATH) or os.getcwd()
    return os.path.join(pasta, "whatsapp.coords")

def calibrar_coordenadas():
    global CLIP_X, CLIP_Y, DOC_X, DOC_Y
    print("\n=== CALIBRAÇÃO DE COORDENADAS ===")
    print("1) Abra o WhatsApp Web e posicione o mouse sobre o ÍCONE DE CLIPE por 3 segundos…")
    time.sleep(3)
    p1 = pyautogui.position()
    print(f"   CLIP: {p1}")
    print("2) Abra o menu do clipe e posicione o mouse sobre 'Documento' por 3 segundos…")
    time.sleep(3)
    p2 = pyautogui.position()
    print(f"   DOCUMENTO: {p2}")
    CLIP_X, CLIP_Y = p1.x, p1.y
    DOC_X, DOC_Y = p2.x, p2.y
    with open(caminho_coords(), "w", encoding="utf-8") as f:
        f.write(f"{CLIP_X},{CLIP_Y}\n{DOC_X},{DOC_Y}\n")
    logging.critical(f"Coordenadas salvas em {caminho_coords()} -> CLIP=({CLIP_X},{CLIP_Y}), DOC=({DOC_X},{DOC_Y})")

def carregar_coordenadas_salvas():
    global CLIP_X, CLIP_Y, DOC_X, DOC_Y
    fp = caminho_coords()
    if os.path.isfile(fp):
        try:
            with open(fp, "r", encoding="utf-8") as f:
                linhas = [l.strip() for l in f if l.strip()]
            if len(linhas) >= 2:
                cx, cy = linhas[0].split(",")
                dx, dy = linhas[1].split(",")
                CLIP_X, CLIP_Y = int(cx), int(cy)
                DOC_X, DOC_Y = int(dx), int(dy)
        except Exception:
            pass


# ======= CSV HISTÓRICO (retomada) =======
def caminho_csv() -> str:
    pasta = os.path.dirname(EXCEL_PATH) or os.getcwd()
    return os.path.join(pasta, NOME_CSV_RELATORIO)

def carregar_ok_anteriores() -> set[int]:
    enviados = set()
    if GERAR_CSV_HISTORICO and os.path.isfile(caminho_csv()):
        try:
            prev = pd.read_csv(caminho_csv(), dtype={'indice': int, 'status': str})
            for _, r in prev.iterrows():
                if str(r.get('status', '')).upper() == 'OK':
                    enviados.add(int(r.get('indice')))
        except Exception:
            pass
    return enviados

def acrescentar_csv(resultados: list[dict]):
    if not GERAR_CSV_HISTORICO:
        return
    csv_path = caminho_csv()
    try:
        if os.path.isfile(csv_path):
            antigo = pd.read_csv(csv_path)
            novo = pd.DataFrame(resultados)
            df = pd.concat([antigo, novo], ignore_index=True)
        else:
            df = pd.DataFrame(resultados)
        df.to_csv(csv_path, index=False, encoding='utf-8-sig')
    except Exception:
        pass

# ======= PDF =======
def soft_wrap(s: str) -> str:
    if not s:
        return ""
    return (s.replace("\\", "\\\u200b")
             .replace("/", "/\u200b")
             .replace(".", ".\u200b")
             .replace("_", "_\u200b")
             .replace("-", "-\u200b")
             .replace(":", ":\u200b"))

def gerar_relatorio_pdf(pasta: str, registros: list[dict], resumo_texto: str, nome_pdf: str) -> str:
    """
    Gera o PDF SEM a coluna 'Prévia da Mensagem'.
    Colunas no PDF: Índice | Nome | Telefone Original | Telefone Normalizado | Tipo Msg | Status | Motivo/Obs. | Data/Hora
    """
    pdf_path = os.path.join(pasta, nome_pdf)
    pagesize = landscape(A4)
    left = right = top = bottom = 1.0 * cm
    page_w, _ = pagesize
    frame_w = page_w - (left + right)
    doc = SimpleDocTemplate(pdf_path, pagesize=pagesize,
                            leftMargin=left, rightMargin=right,
                            topMargin=top, bottomMargin=bottom)

    styles = getSampleStyleSheet()
    body = ParagraphStyle("BodySmall", parent=styles["BodyText"], fontName="Helvetica", fontSize=8.5, leading=11, wordWrap="CJK")
    head = ParagraphStyle("HeadSmall", parent=styles["BodyText"], fontName="Helvetica-Bold", fontSize=9, leading=11, alignment=1)

    story = []
    story.append(Paragraph("<b>Relatório de Envio – WhatsApp</b>", styles["Title"]))
    story.append(Spacer(1, 0.2*cm))
    story.append(Paragraph(resumo_texto, body))
    story.append(Spacer(1, 0.5*cm))

    colunas = [
        ("indice", "Índice"),
        ("nome", "Nome"),
        ("telefone_original", "Telefone Original"),
        ("telefone_normalizado", "Telefone Normalizado"),
        ("tipo_msg", "Tipo Msg"),
        ("status", "Status"),
        ("motivo", "Motivo/Obs."),
        ("quando", "Data/Hora"),
    ]

    header = [Paragraph(f"<b>{rot}</b>", head) for _, rot in colunas]
    data = [header]

    for r in registros or []:
        row = []
        for key, _rot in colunas:
            val = "" if r.get(key) is None else str(r.get(key))
            row.append(Paragraph(soft_wrap(val), body))
        data.append(row)

    fractions = {
        "indice": 0.06, "nome": 0.18, "telefone_original": 0.12, "telefone_normalizado": 0.12,
        "tipo_msg": 0.08, "status": 0.08, "motivo": 0.20, "quando": 0.16
    }
    total_frac = sum(fractions.values())
    for k in fractions: fractions[k] = fractions[k]/total_frac
    col_widths = [fractions[key]*frame_w for key, _ in colunas]

    table = Table(data, colWidths=col_widths, repeatRows=1)
    table.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#E8EEF7")),
        ("TEXTCOLOR", (0,0), (-1,0), colors.HexColor("#0B2239")),
        ("ALIGN", (0,0), (-1,0), "CENTER"),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("VALIGN", (0,0), (-1,-1), "TOP"),
        ("FONTSIZE", (0,1), (-1,-1), 8.5),
        ("ALIGN", (0,1), (0,-1), "RIGHT"),
        ("ALIGN", (5,1), (5,-1), "CENTER"),  # Status centralizado
        ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.white, colors.HexColor("#F7F9FC")]),
        ("BOX", (0,0), (-1,-1), 0.5, colors.HexColor("#B3C1D1")),
        ("GRID", (0,0), (-1,-1), 0.25, colors.HexColor("#D3DDE8")),
        ("LEFTPADDING", (0,0), (-1,-1), 4),
        ("RIGHTPADDING", (0,0), (-1,-1), 4),
        ("TOPPADDING", (0,0), (-1,-1), 3),
        ("BOTTOMPADDING", (0,0), (-1,-1), 3),
    ]))

    story.append(table)
    doc.build(story)
    return pdf_path


# ======= PIPELINES =======
def gerar_pdf_de_csv():
    """
    Gera o PDF a partir de um CSV existente (se MODO_SOMENTE_PDF_DE_CSV=True).
    """
    csv_path = CSV_RESULTADOS_FONTE or caminho_csv()
    if not os.path.isfile(csv_path):
        raise FileNotFoundError(f"CSV não encontrado: {csv_path}")
    df = pd.read_csv(csv_path, dtype=str).fillna("")
    esperadas = ["indice","nome","telefone_original","telefone_normalizado","tipo_msg","status","motivo","quando"]
    for c in esperadas:
        if c not in df.columns: df[c] = ""
    registros = df[esperadas].to_dict(orient="records")
    pasta = os.path.dirname(csv_path) or os.getcwd()
    ok = sum(1 for r in registros if str(r.get("status","")).upper()=="OK")
    erro = sum(1 for r in registros if str(r.get("status","")).upper()=="ERRO")
    resumo = f"Total: {len(registros)} | Sucesso: {ok} | Erros: {erro} | Fonte: {os.path.basename(csv_path)}"
    gerar_relatorio_pdf(pasta, registros, resumo, NOME_PDF_RELATORIO)
    # Apenas mensagem final curta:
    print("✅ Relatório PDF salvo com sucesso!")

def enviar_e_gerar_relatorio():
    # validações iniciais
    if not os.path.isfile(EXCEL_PATH):
        raise FileNotFoundError(f"Planilha não encontrada: {EXCEL_PATH}")
    if ENVIAR_ANEXO and not os.path.isfile(ARQUIVO_PDF):
        raise FileNotFoundError(f"Arquivo PDF não encontrado: {ARQUIVO_PDF}")

    pasta = os.path.dirname(EXCEL_PATH) or os.getcwd()
    preparar_logging(pasta)
    carregar_coordenadas_salvas()

    if CALIBRAR_COORDENADAS and not DRY_RUN:
        checar_login_whatsapp()
        calibrar_coordenadas()

    if CHECAR_LOGIN_INICIAL and not DRY_RUN:
        checar_login_whatsapp()

    # carregar planilha
    df = pd.read_excel(EXCEL_PATH, dtype={'Contato': str}).fillna("")
    for col in ["Nome", "Contato"]:
        if col not in df.columns:
            raise ValueError(f"Coluna obrigatória ausente na planilha: {col}")

    ja_ok = carregar_ok_anteriores() if GERAR_CSV_HISTORICO else set()

    resultados = []

    for i, linha in df.iterrows():
        if i in ja_ok:
            continue

        nome = str(linha["Nome"]).strip()
        tel_original = str(linha["Contato"]).strip()
        tel_normalizado, motivo_tel = normalizar_telefone(tel_original)
        mensagem, tipo_msg = montar_mensagem(nome)

        # Registro base
        base_reg = {
            "indice": i,
            "nome": nome,
            "telefone_original": tel_original,
            "telefone_normalizado": tel_normalizado if tel_normalizado else "",
            "tipo_msg": tipo_msg,
            "quando": datetime.now().isoformat()
        }

        if not tel_normalizado:
            reg = {**base_reg, "status": "ERRO", "motivo": f"telefone inválido: {motivo_tel}"}
            resultados.append(reg)
            continue

        try:
            if ABRIR_NAVEGADOR and not DRY_RUN:
                abrir_chat(tel_normalizado, mensagem)
                time.sleep(jitter(TEMPO_CARREGAR_CHAT, 'carregar'))
                enviar_texto()
                time.sleep(jitter(TEMPO_APOS_ENVIAR_TEXTO, 'apos_texto'))
                if ENVIAR_ANEXO:
                    anexar_pdf(ARQUIVO_PDF)
                    time.sleep(jitter(TEMPO_APOS_ENVIAR_ARQUIVO, 'apos_arquivo'))
                if FECHAR_ABA_APOS_ENVIO:
                    fechar_aba()
                    time.sleep(jitter(TEMPO_APOS_FECHAR_ABA, 'final'))

            reg = {**base_reg, "status": "OK", "motivo": ""}
            resultados.append(reg)
            time.sleep(random.uniform(*INTERVALO_ENTRE_CONTATOS))

        except Exception as e:
            try:
                if FECHAR_ABA_APOS_ENVIO and not DRY_RUN:
                    fechar_aba()
            except Exception:
                pass
            reg = {**base_reg, "status": "ERRO", "motivo": str(e)}
            resultados.append(reg)

    # Atualiza CSV histórico (se habilitado)
    acrescentar_csv(resultados)

    # Gera PDF consolidado
    ok = sum(1 for r in resultados if r['status'] == 'OK')
    erro = sum(1 for r in resultados if r['status'] == 'ERRO')
    resumo = f"✅ Concluído. Sucesso: {ok} | Erros: {erro} | Fonte: {os.path.basename(EXCEL_PATH)}"
    gerar_relatorio_pdf(pasta, resultados, resumo, NOME_PDF_RELATORIO)

    # Apenas mensagem final curta (sem caminho, sem resumo):
    print("✅ Relatório PDF salvo com sucesso!")


# ======= MAIN =======
if __name__ == "__main__":
    if MODO_SOMENTE_PDF_DE_CSV:
        gerar_pdf_de_csv()
    else:
        enviar_e_gerar_relatorio()
