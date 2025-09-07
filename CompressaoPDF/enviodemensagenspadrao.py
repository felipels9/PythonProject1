import os
import re
import time
import random
import urllib.parse
import pandas as pd
import pyautogui
import webbrowser
import pyperclip
from datetime import datetime

# ====== REPORTLAB (PDF) ======
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm

# ============ FLAGS ============
MODO_SOMENTE_PDF_DE_CSV = False   # True = lê CSV e gera PDF; False = executa envios e gera PDF no fim
CSV_RESULTADOS = r"C:\CAMINHO\para\relatorio_envio_whatsapp.csv"  # usado só se MODO_SOMENTE_PDF_DE_CSV=True

# ============ CONFIG ENVIO (usado apenas no modo 2) ============
ARQUIVO_PDF = r"C:\Users\lealf\OneDrive\Área de trabalho - atual\FELIPE LEAL\CONVITE DE CASAMENTO\CONVITEDECASAMENTO.pdf"
PLANILHA_XLSX = r"C:\Users\lealf\OneDrive\Área de trabalho - atual\FELIPE LEAL\CONVITE DE CASAMENTO\teste.xlsx"

ABRIR_NAVEGADOR = True
FECHAR_ABA_APOS_ENVIO = True
SEM_ANEXO = False
DRY_RUN = False
CALIBRAR_COORDENADAS = False

TEMPO_CARREGAR_CHAT = 22
TEMPO_APOS_ENVIAR_TEXTO = 3
TEMPO_APOS_CLIP = 2
TEMPO_APOS_DOC = 2
TEMPO_APOS_COLAR_CAMINHO = 1
TEMPO_UPLOAD = 5
TEMPO_APOS_ENVIAR_ARQUIVO = 8
TEMPO_APOS_FECHAR_ABA = 2
INTERVALO_ENTRE_CONTATOS = (1.5, 4.0)

VARIACAO = {'carregar': 4,'apos_texto': 1,'apos_clip': 1,'apos_doc': 1,'upload': 3,'apos_arquivo': 2,'final': 1}

# Coordenadas padrão (calibre se precisar)
CLIP_X, CLIP_Y = 709, 966
DOC_X, DOC_Y   = 686, 536

# Segurança
pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.5

# ============ UTIL ============
def jitter(base: float, key: str) -> float:
    delta = VARIACAO.get(key, 0)
    return max(0, base + random.uniform(-delta, delta)) if delta > 0 else base

def somente_digitos(s: str) -> str:
    return re.sub(r"\D", "", (s or "").strip())

def normalizar_telefone_br(telefone: str) -> tuple[str, str]:
    d = somente_digitos(telefone)
    if not d:
        return "", "vazio"
    d = d.lstrip('0')
    if not d.startswith('55'):
        d = '55' + d
    sem_ddi = d[2:]
    if len(sem_ddi) < 10 or len(sem_ddi) > 11:
        return "", f"tamanho inválido: {len(sem_ddi)} dígitos após DDI"
    ddd = sem_ddi[:2]
    numero = sem_ddi[2:]
    if not (len(ddd) == 2 and ddd.isdigit() and 11 <= int(ddd) <= 99):
        return "", f"DDD inválido: {ddd}"
    if len(numero) not in (8, 9):
        return "", f"número local inválido: {numero}"
    return d, ""

def montar_mensagem(nome: str) -> str:
    """
    Define a mensagem conforme o 'nome':
    - Se o nome sugerir plural (contém ' e ' ou ' & ') -> 'presença de vocês'
    - Senão -> 'sua presença'
    """
    nome_lower = (nome or "").lower()
    if " e " in nome_lower or " & " in nome_lower:
        return f"""Para: {nome}

Estamos muito felizes e contamos com a presença de vocês em nosso grande dia!"""
    else:
        return f"""Para: {nome}

Estamos muito felizes e contamos com sua presença em nosso grande dia!"""

def abrir_chat(telefone_normalizado: str, mensagem: str):
    link = (f"https://web.whatsapp.com/send?phone={telefone_normalizado}"
            f"&text={urllib.parse.quote(mensagem)}")
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

def carregar_tabela(caminho_xlsx: str) -> pd.DataFrame:
    df = pd.read_excel(caminho_xlsx, dtype={'Contato': str})
    for col in ['Nome', 'Contato']:
        if col not in df.columns:
            raise ValueError(f"Coluna obrigatória ausente na planilha: {col}")
    df = df.dropna(subset=['Contato'])
    return df

def checar_login_whatsapp():
    if DRY_RUN or not ABRIR_NAVEGADOR: return
    webbrowser.open("https://web.whatsapp.com")
    time.sleep(12)

# ============ PDF ============
def gerar_relatorio_pdf(pasta: str, resultados: list[dict], resumo_texto: str) -> str:
    """
    PDF A4 paisagem com tabela ajustada (não corta conteúdo).
    """
    pdf_path = os.path.join(pasta, "relatorio_envio_whatsapp.pdf")

    pagesize = landscape(A4)
    left = right = top = bottom = 1.0 * cm
    page_w, _ = pagesize
    frame_w = page_w - (left + right)

    doc = SimpleDocTemplate(
        pdf_path, pagesize=pagesize,
        leftMargin=left, rightMargin=right, topMargin=top, bottomMargin=bottom
    )

    styles = getSampleStyleSheet()
    body = ParagraphStyle(
        "BodySmall", parent=styles["BodyText"],
        fontName="Helvetica", fontSize=8.5, leading=11, wordWrap="CJK"
    )
    head = ParagraphStyle(
        "HeadSmall", parent=styles["BodyText"],
        fontName="Helvetica-Bold", fontSize=9, leading=11, alignment=1
    )

    def soft_wrap(s: str) -> str:
        if not s: return ""
        return (s.replace("\\", "\\\u200b")
                 .replace("/", "/\u200b")
                 .replace(".", ".\u200b")
                 .replace("_", "_\u200b")
                 .replace("-", "-\u200b")
                 .replace(":", ":\u200b"))

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
        ("status", "Status"),
        ("motivo", "Motivo/Observação"),
        ("quando", "Data/Hora"),
    ]

    header = [Paragraph(f"<b>{rot}</b>", head) for _, rot in colunas]
    data = [header]

    for r in resultados or []:
        row = []
        for key, _rotulo in colunas:
            val = "" if r.get(key) is None else str(r.get(key))
            row.append(Paragraph(soft_wrap(val), body))
        data.append(row)

    fractions = {
        "indice": 0.05, "nome": 0.18, "telefone_original": 0.12,
        "telefone_normalizado": 0.12, "status": 0.08,
        "motivo": 0.33, "quando": 0.12
    }
    total_frac = sum(fractions.values())
    for k in fractions: fractions[k] = fractions[k] / total_frac
    col_widths = [fractions[key] * frame_w for key, _ in colunas]

    table = Table(data, colWidths=col_widths, repeatRows=1)
    table.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#E8EEF7")),
        ("TEXTCOLOR", (0,0), (-1,0), colors.HexColor("#0B2239")),
        ("ALIGN", (0,0), (-1,0), "CENTER"),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("VALIGN", (0,0), (-1,-1), "TOP"),
        ("FONTSIZE", (0,1), (-1,-1), 8.5),
        ("ALIGN", (0,1), (0,-1), "RIGHT"),
        ("ALIGN", (4,1), (4,-1), "CENTER"),
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

# ============ PIPELINES ============
def gerar_pdf_de_csv(caminho_csv: str) -> str:
    if not os.path.isfile(caminho_csv):
        raise FileNotFoundError(f"CSV não encontrado: {caminho_csv}")
    df = pd.read_csv(caminho_csv, dtype=str).fillna("")
    # Normaliza nomes de colunas esperadas; se não tiver, cria vazias
    esperadas = ["indice","nome","telefone_original","telefone_normalizado","status","motivo","quando"]
    for col in esperadas:
        if col not in df.columns:
            df[col] = ""
    resultados = df[esperadas].to_dict(orient="records")
    pasta = os.path.dirname(caminho_csv) or os.getcwd()
    ok = sum(1 for r in resultados if str(r.get("status","")).upper()=="OK")
    erro = sum(1 for r in resultados if str(r.get("status","")).upper()=="ERRO")
    resumo = f"Total: {len(resultados)} | Sucesso: {ok} | Erros: {erro} | Fonte: {os.path.basename(caminho_csv)}"
    pdf_path = gerar_relatorio_pdf(pasta, resultados, resumo)
    print(f"PDF gerado: {pdf_path}")
    return pdf_path

def processar_e_gerar_pdf():
    # validações
    if not os.path.isfile(PLANILHA_XLSX):
        raise FileNotFoundError(f"Planilha não encontrada: {PLANILHA_XLSX}")
    if not SEM_ANEXO and not os.path.isfile(ARQUIVO_PDF):
        raise FileNotFoundError(f"PDF não encontrado: {ARQUIVO_PDF}")

    pasta = os.path.dirname(PLANILHA_XLSX) or os.getcwd()
    if not DRY_RUN:
        checar_login_whatsapp()

    df = carregar_tabela(PLANILHA_XLSX)
    resultados = []

    for idx, linha in df.iterrows():
        nome = str(linha.get('Nome', '')).strip()
        telefone_raw = str(linha.get('Contato', '')).strip()
        tel_norm, motivo = normalizar_telefone_br(telefone_raw)

        if not tel_norm:
            resultados.append({'indice': idx,'nome': nome,'telefone_original': telefone_raw,
                               'telefone_normalizado': '', 'status': 'ERRO',
                               'motivo': f'telefone inválido: {motivo}',
                               'quando': datetime.now().isoformat()})
            continue

        # >>> usa a nova lógica pedida (singular/plural)
        mensagem = montar_mensagem(nome)

        try:
            if ABRIR_NAVEGADOR and not DRY_RUN:
                abrir_chat(tel_norm, mensagem)
                time.sleep(jitter(TEMPO_CARREGAR_CHAT, 'carregar'))
                enviar_texto()
                time.sleep(jitter(TEMPO_APOS_ENVIAR_TEXTO, 'apos_texto'))
                if not SEM_ANEXO:
                    anexar_pdf(ARQUIVO_PDF)
                    time.sleep(jitter(TEMPO_APOS_ENVIAR_ARQUIVO, 'apos_arquivo'))
                if FECHAR_ABA_APOS_ENVIO:
                    fechar_aba()
                    time.sleep(jitter(TEMPO_APOS_FECHAR_ABA, 'final'))

            resultados.append({'indice': idx,'nome': nome,'telefone_original': telefone_raw,
                               'telefone_normalizado': tel_norm, 'status': 'OK',
                               'motivo': '', 'quando': datetime.now().isoformat()})
            time.sleep(random.uniform(*INTERVALO_ENTRE_CONTATOS))

        except Exception as e:
            try:
                if FECHAR_ABA_APOS_ENVIO and not DRY_RUN:
                    fechar_aba()
            except Exception:
                pass
            resultados.append({'indice': idx,'nome': nome,'telefone_original': telefone_raw,
                               'telefone_normalizado': tel_norm, 'status': 'ERRO',
                               'motivo': str(e), 'quando': datetime.now().isoformat()})

    ok = sum(1 for r in resultados if r['status'] == 'OK')
    erro = sum(1 for r in resultados if r['status'] == 'ERRO')
    resumo = f"✅ Concluído. Sucesso: {ok} | Erros: {erro} | Fonte: {os.path.basename(PLANILHA_XLSX)}"
    pdf_path = gerar_relatorio_pdf(pasta, resultados, resumo)
    print(f"PDF gerado: {pdf_path}")

# ============ MAIN ============
if __name__ == "__main__":
    if MODO_SOMENTE_PDF_DE_CSV:
        gerar_pdf_de_csv(CSV_RESULTADOS)
    else:
        processar_e_gerar_pdf()
