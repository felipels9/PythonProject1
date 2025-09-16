import os
import shutil
import subprocess
import tempfile
import threading
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
from typing import List, Dict, Optional, Tuple
from concurrent.futures import ThreadPoolExecutor, as_completed
from queue import Queue, Empty

from PyPDF2 import PdfReader, PdfWriter

# ===================== Parâmetros =====================
LIMITE_MB = 5.0                     # limite PJe por arquivo
QUALIDADES_GS = ("/screen", "/ebook", "/printer", "/prepress")
GS_TIMEOUT_SEC = 120                # timeout base por execução do Ghostscript
MAX_WORKERS_DEFAULT = max(2, (os.cpu_count() or 4) - 1)  # paralelismo da pré-compressão
MARGEM_MB = 0.12                    # margem para overhead ao unir sem recomprimir (~120KB)

# ===================== Utils =====================

def _mb(path: str) -> float:
    try:
        return os.path.getsize(path) / (1024 * 1024)
    except Exception:
        return 0.0

def _nome_parte(base_saida: str, indice: int) -> str:
    base, ext = os.path.splitext(base_saida)
    return f"{base}_parte_{indice:02d}{ext}"

def _norm(p: str) -> str:
    return os.path.abspath(p).replace("\\", "/")

def _is_windows() -> bool:
    return os.name == "nt"

# ===================== Canal de notificação (thread-safe) =====================

class Notifier:
    def __init__(self, queue: Queue):
        self.q = queue
        self.msgs: List[Tuple[str, str]] = []

    def text(self, s: str): self.q.put(("text", s))
    def subtext(self, s: str): self.q.put(("subtext", s))
    def set_total(self, n: int): self.q.put(("set_total", n))
    def step_to(self, i: int): self.q.put(("step", i))

    def done(self, result): self.q.put(("done", result))
    def close(self): self.q.put(("close", None))

    def info(self, s: str): self.msgs.append(("info", s))
    def warn(self, s: str): self.msgs.append(("warn", s))
    def error(self, s: str): self.msgs.append(("error", s))

# ===================== Ghostscript =====================

def _encontrar_ghostscript() -> Optional[str]:
    candidatos = [
        r"C:/Program Files/gs/gs10.05.1/bin/gswin64c.exe",
        r"C:/Program Files (x86)/gs/gs10.05.1/bin/gswin32c.exe",
        shutil.which("gswin64c"),
        shutil.which("gswin32c"),
        shutil.which("gs"),
    ]
    for p in candidatos:
        if p and os.path.exists(p):
            return p
    return None

# --------- STAGING ASCII + response file (cwd no stage) ---------

def _stage_inputs_for_gs(input_files: List[str]) -> Tuple[str, List[str]]:
    """
    Copia os PDFs para um diretório temporário único (por chamada),
    com nomes ASCII sequenciais 00001.pdf, 00002.pdf...
    """
    stage_dir = tempfile.mkdtemp(prefix=f"gs_stage_{os.getpid()}_")
    staged = []
    for idx, src in enumerate(input_files, start=1):
        dst = os.path.join(stage_dir, f"{idx:05d}.pdf")
        shutil.copyfile(os.path.abspath(src), dst)
        staged.append(dst)
    return stage_dir, staged

def _write_gs_listfile_basename(files_in_stage: List[str], stage_dir: str) -> str:
    """
    Cria um response file .lst contendo APENAS os nomes (basename) dos PDFs, um por linha.
    """
    list_path = os.path.join(stage_dir, f"gs_inputs_{os.getpid()}_{threading.get_ident()}.lst")
    with open(list_path, "w", encoding="ascii", errors="strict") as f:
        for p in files_in_stage:
            f.write(os.path.basename(p) + "\n")
    return list_path

def _run_gs(args: List[str], timeout: int, cwd: Optional[str], notify: Notifier) -> Tuple[bool, str]:
    """
    Executa o Ghostscript capturando stderr/stdout. Retorna (ok, stderr_text).
    """
    popen_kwargs = dict(
        shell=False, timeout=timeout, cwd=cwd,
        stdout=subprocess.PIPE, stderr=subprocess.PIPE
    )
    if _is_windows():
        popen_kwargs["creationflags"] = subprocess.CREATE_NO_WINDOW
    try:
        proc = subprocess.run(args, check=False, **popen_kwargs)
        err = (proc.stderr or b"").decode(errors="replace").strip()
        out = (proc.stdout or b"").decode(errors="replace").strip()
        if proc.returncode == 0:
            return True, err
        msg = f"Ghostscript retornou código {proc.returncode}."
        if err: msg += f"\n\n[stderr]\n{err}"
        if out: msg += f"\n\n[stdout]\n{out}"
        if "undefinedfilename" in err.lower() or "cannot find" in err.lower():
            msg += "\n\nDica: verifique se algum PDF foi movido/renomeado, se há bloqueio do OneDrive, ou se há caracteres incomuns no nome ORIGINAL."
        notify.warn(msg)
        return False, err
    except subprocess.TimeoutExpired:
        notify.error(f"Tempo esgotado ({timeout}s) ao executar o Ghostscript.")
        return False, "timeout"
    except Exception as e:
        notify.error(f"Falha ao executar o Ghostscript: {e}")
        return False, str(e)

def gs_comprimir_para_pdf(input_files: List[str], output_pdf: str, qualidade: str, notify: Notifier) -> bool:
    """
    Integração robusta:
    - staging com nomes ASCII e cwd no stage
    - response file com basenames
    - se falhar com múltiplos, testa arquivos 1 a 1 (GS) para isolar e excluir problemáticos
    """
    gs_path = _encontrar_ghostscript()
    if not gs_path:
        notify.error("Ghostscript não encontrado. Instale o Ghostscript ou adicione-o ao PATH.")
        return False

    out_norm = _norm(output_pdf)
    os.makedirs(os.path.dirname(out_norm) or ".", exist_ok=True)

    dyn_timeout = min(900, max(GS_TIMEOUT_SEC, 90 + 3 * max(1, len(input_files))))

    # 1) STAGING
    try:
        stage_dir, staged_files = _stage_inputs_for_gs(input_files)
    except Exception as e:
        notify.error(str(e))
        return False

    def _call_gs_with_list(files_in_stage: List[str], out_path: str) -> Tuple[bool, str]:
        listfile = _write_gs_listfile_basename(files_in_stage, stage_dir)
        args = [
            gs_path,
            "-sDEVICE=pdfwrite",
            "-dCompatibilityLevel=1.4",
            f"-dPDFSETTINGS={qualidade}",
            "-dNOPAUSE",
            "-dBATCH",
            "-dQUIET",
            "-dDetectDuplicateImages=true",
            "-dDownsampleColorImages=true",
            "-dColorImageDownsampleType=/Bicubic",
            "-dColorImageResolution=110",
            "-dDownsampleGrayImages=true",
            "-dGrayImageDownsampleType=/Bicubic",
            "-dGrayImageResolution=110",
            "-dDownsampleMonoImages=true",
            "-dMonoImageDownsampleType=/Subsample",
            "-dMonoImageResolution=110",
            f"-sOutputFile={_norm(out_path)}",
            "-f",
            f"@{os.path.basename(listfile)}",
        ]
        try:
            ok, err = _run_gs(args, timeout=dyn_timeout, cwd=stage_dir, notify=notify)
            return ok and os.path.isfile(out_path), err
        finally:
            try: os.remove(listfile)
            except Exception: pass

    try:
        # 2) Tenta direto com todos
        ok, err = _call_gs_with_list(staged_files, out_norm)
        if ok:
            return True

        # 3) Se falhou e havia múltiplos arquivos, isola culpados SEM usar PyPDF2
        if len(staged_files) > 1:
            notify.subtext("Isolando PDFs problemáticos…")
            bons: List[str] = []
            ruins: List[str] = []
            for p in staged_files:
                tmp_out = os.path.join(stage_dir, f"test_{os.path.basename(p)}")
                ok_one, _ = _call_gs_with_list([p], tmp_out)
                try:
                    if os.path.exists(tmp_out):
                        os.remove(tmp_out)
                except Exception:
                    pass
                (bons if ok_one else ruins).append(p)

            if ruins:
                base_names = [os.path.basename(x) for x in ruins]
                notify.warn("Alguns PDFs estão corrompidos/incompatíveis e foram ignorados nesta etapa:\n\n" +
                            "\n".join(base_names))

            if not bons:
                return False

            ok2, _ = _call_gs_with_list(bons, out_norm)
            return ok2

        return False

    finally:
        # limpeza do stage
        try:
            for name in os.listdir(stage_dir):
                try: os.remove(os.path.join(stage_dir, name))
                except Exception: pass
            os.rmdir(stage_dir)
        except Exception:
            pass

# ===================== Remoção de páginas em branco =====================

def _page_has_xobject_or_annots(page) -> bool:
    try:
        ann = page.get("/Annots")
        if ann and len(ann) > 0:
            return True
    except Exception:
        pass
    try:
        resources = page.get("/Resources")
        if resources:
            xobj = resources.get("/XObject")
            if xobj:
                for obj in xobj.values():
                    try:
                        o = obj.get_object()
                        subtype = o.get("/Subtype")
                        if subtype in ["/Image", "/Form"]:
                            return True
                    except Exception:
                        return True
    except Exception:
        return True
    return False

def is_page_blank_fast(page) -> bool:
    if _page_has_xobject_or_annots(page):
        return False
    try:
        contents = page.get_contents()
        if contents is None:
            return True
        if isinstance(contents, list):
            data = b"".join([c.get_data() for c in contents if hasattr(c, "get_data")])
        else:
            data = contents.get_data()
        return len(data.strip()) == 0
    except Exception:
        return False

def limpar_brancos_para_arquivo(entrada_pdf: str) -> str:
    try:
        reader = PdfReader(entrada_pdf)
    except Exception:
        return entrada_pdf

    writer = PdfWriter()
    removidos = 0
    for p in reader.pages:
        if is_page_blank_fast(p):
            removidos += 1
            continue
        writer.add_page(p)

    if removidos == 0:
        return entrada_pdf

    tmp_out = os.path.join(tempfile.gettempdir(), f"temp_clean_{os.getpid()}_{os.path.basename(entrada_pdf)}")
    with open(tmp_out, "wb") as f:
        writer.write(f)
    return tmp_out

# ===================== Unir sem recomprimir =====================

def unir_pdfs_sem_recomprimir(pdf_paths: List[str], out_path: str) -> bool:
    """
    Une PDFs já comprimidos (saída do GS) copiando páginas com PyPDF2.
    É bem rápido e não reprocessa imagens.
    """
    writer = PdfWriter()
    total = 0
    for p in pdf_paths:
        r = PdfReader(p)
        for pg in r.pages:
            writer.add_page(pg)
            total += 1
    if total == 0:
        return False
    with open(out_path, "wb") as fo:
        writer.write(fo)
    return True

# ===================== Escrita parcial (PyPDF2) =====================

def escrever_intervalo_de_paginas(input_pdf: str, start_idx: int, end_idx: int, out_path: str, remover_brancos: bool = True):
    reader = PdfReader(input_pdf)
    writer = PdfWriter()
    for i in range(start_idx, end_idx):
        p = reader.pages[i]
        if remover_brancos and is_page_blank_fast(p):
            continue
        writer.add_page(p)
    with open(out_path, "wb") as f:
        writer.write(f)

# ===================== Cache de compressão (paralelo) =====================

class CompressCache:
    """
    Pré-comprime cada arquivo individualmente (em paralelo) e mede o tamanho.
    caminho_original -> {"cleaned": caminho_entrada, "compressed": caminho_comp, "mb": tamanho}
    """
    def __init__(self, qualidade: str, remover_brancos: bool, max_workers: Optional[int], notify: Notifier):
        self.qualidade = qualidade
        self.remover_brancos = remover_brancos
        self.cache: Dict[str, Dict] = {}
        self.max_workers = max_workers or max(2, min(MAX_WORKERS_DEFAULT, (os.cpu_count() or 4)))
        self.notify = notify

    def _clean_if_needed(self, caminho: str) -> str:
        # IMPORTANTE: remover páginas em branco é custoso; use com parcimônia
        return limpar_brancos_para_arquivo(caminho) if self.remover_brancos else caminho

    def _build_one(self, caminho: str) -> Optional[Tuple[str, Dict]]:
        try:
            cleaned = self._clean_if_needed(caminho)
            temp_out = os.path.join(tempfile.gettempdir(), f"temp_comp_ind_{os.getpid()}_{threading.get_ident()}_{os.path.basename(cleaned)}")
            ok = gs_comprimir_para_pdf([cleaned], temp_out, qualidade=self.qualidade, notify=self.notify)
            if not ok:
                return None
            info = {"cleaned": cleaned, "compressed": temp_out, "mb": _mb(temp_out)}
            return caminho, info
        except Exception as e:
            print(f"[CompressCache] Falha em {caminho}: {e}")
            return None

    def build_many(self, caminhos: List[str]) -> List[Dict]:
        self.notify.text("Pré-compressão paralela (modo preciso)")
        self.notify.set_total(len(caminhos))
        self.notify.step_to(0)

        faltam = [c for c in caminhos if c not in self.cache]
        done = 0
        if faltam:
            with ThreadPoolExecutor(max_workers=self.max_workers) as ex:
                futs = {ex.submit(self._build_one, c): c for c in faltam}
                for fut in as_completed(futs):
                    res = fut.result()
                    if res:
                        key, info = res
                        self.cache[key] = info
                    else:
                        c = futs[fut]
                        self.cache[c] = {"cleaned": c, "compressed": None, "mb": float("inf")}
                    done += 1
                    self.notify.step_to(done)
        if done < len(caminhos):
            self.notify.step_to(len(caminhos))
        return [self.cache.get(c, {"cleaned": c, "compressed": None, "mb": float("inf")}) for c in caminhos]

# ===================== Lógica principal (worker) =====================

def _split_single_pdf_por_paginas(input_pdf: str, base_saida: str, qualidade: str, start_ind: int,
                                  remover_brancos: bool, notify: Notifier, cancel_flag: threading.Event) -> List[str]:
    """
    Divide um único PDF em pedaços ≤ LIMITE_MB (saltos exponenciais + busca binária).
    """
    out_paths: List[str] = []
    reader = PdfReader(input_pdf)
    N = len(reader.pages)
    i = 0
    indice_parte = start_ind

    notify.text("Dividindo PDF grande por páginas")
    notify.subtext(os.path.basename(input_pdf))

    def comp_range(k: int) -> Tuple[Optional[str], float]:
        tmp_unido = os.path.join(tempfile.gettempdir(), f"temp_chunk_{os.getpid()}_{threading.get_ident()}.pdf")
        escrever_intervalo_de_paginas(input_pdf, i, min(i + k, N), tmp_unido, remover_brancos=remover_brancos)
        tmp_out = os.path.join(tempfile.gettempdir(), f"temp_chunk_comp_{os.getpid()}_{threading.get_ident()}.pdf")
        ok = gs_comprimir_para_pdf([tmp_unido], tmp_out, qualidade=qualidade, notify=notify)
        try: os.remove(tmp_unido)
        except Exception: pass
        if not ok:
            return None, 0.0
        return tmp_out, _mb(tmp_out)

    while i < N and not cancel_flag.is_set():
        best_ok = 0
        high = 1

        # crescimento exponencial
        while not cancel_flag.is_set():
            tmp_path, size_mb = comp_range(high)
            if tmp_path is None:
                break
            if size_mb <= LIMITE_MB:
                best_ok = high
                try: os.remove(tmp_path)
                except Exception: pass
                high *= 2
                if i + best_ok >= N:
                    break
            else:
                try: os.remove(tmp_path)
                except Exception: pass
                break

        if best_ok == 0:
            tmp_single, _ = comp_range(1)
            saida = _nome_parte(base_saida, indice_parte)
            shutil.move(tmp_single, saida)
            out_paths.append(saida)
            indice_parte += 1
            i += 1
            notify.subtext(f"Gerada parte {indice_parte - start_ind}")
            continue

        lo, hi = best_ok, min(high, N - i)
        while lo < hi and not cancel_flag.is_set():
            mid = (lo + hi + 1) // 2
            tmp_mid, size_mb = comp_range(mid)
            if tmp_mid is None:
                break
            if size_mb <= LIMITE_MB:
                try: os.remove(tmp_mid)
                except Exception: pass
                lo = mid
            else:
                try: os.remove(tmp_mid)
                except Exception: pass
                hi = mid - 1

        tmp_ok, _ = comp_range(lo)
        saida = _nome_parte(base_saida, indice_parte)
        shutil.move(tmp_ok, saida)
        out_paths.append(saida)
        indice_parte += 1
        i += lo
        notify.subtext(f"Gerada parte {indice_parte - start_ind}")

    return out_paths

def processar_com_limite_worker(arquivos_ordenados: List[str], destino_final: str, qualidade: str,
                                remover_brancos: bool, modo_turbo: bool,
                                notify: Notifier, cancel_flag: threading.Event) -> List[str]:
    if cancel_flag.is_set():
        return []

    # 0) Remoção de brancos (opcional – custa I/O/CPU)
    notify.text("Preparando documentos")
    cleaned_inputs = [limpar_brancos_para_arquivo(p) for p in arquivos_ordenados] if remover_brancos else list(arquivos_ordenados)

    # 1) Comprime tudo em um só (sempre fazemos isso — é 1 chamada ao GS)
    notify.text("Comprimindo conjunto inteiro")
    notify.subtext("Passo 1/2")
    notify.set_total(1); notify.step_to(0)

    if not gs_comprimir_para_pdf(cleaned_inputs, destino_final, qualidade=qualidade, notify=notify):
        notify.error("Falha ao comprimir PDF único.")
        return []
    notify.step_to(1)
    tam = _mb(destino_final)
    if tam <= LIMITE_MB:
        notify.info(f"✅ PDF comprimido salvo com {tam:.2f} MB\n\n{destino_final}")
        return [destino_final]

    # 2A) MODO TURBO: dividir o PDF já comprimido por páginas (bem rápido, pouquíssimos GS)
    if modo_turbo:
        notify.text("Modo Turbo: dividindo por páginas")
        partes = _split_single_pdf_por_paginas(destino_final, destino_final, qualidade,
                                               start_ind=1, remover_brancos=False, notify=notify, cancel_flag=cancel_flag)
        # opcional: remover o "inteiro" que estourou
        try:
            if os.path.exists(destino_final) and partes:
                os.remove(destino_final)
        except Exception:
            pass
        return _finalizar_partes_existentes(partes, notify)

    # 2B) MODO PRECISO: pré-comprime cada PDF em paralelo e une sem recomprimir
    notify.text("Modo Preciso: pré-compressão paralela")
    cache = CompressCache(qualidade=qualidade, remover_brancos=False, max_workers=None, notify=notify)
    infos = cache.build_many(cleaned_inputs)

    # filtra falhas
    validos: List[Dict] = []
    grandoes: List[str] = []
    for info in infos:
        if info["compressed"] is None or info["mb"] == float("inf"):
            continue
        # se um arquivo já comprimido continuar > 5MB, marca para dividir por páginas
        if info["mb"] > LIMITE_MB:
            grandoes.append(info["cleaned"])
        else:
            validos.append(info)

    if not validos and not grandoes:
        notify.error("Nenhum PDF pôde ser processado.")
        return []

    # trata arquivos que ainda estão > 5MB individualmente
    partes_saida: List[str] = []
    for src in grandoes:
        notify.subtext(f"Dividindo documento grande: {os.path.basename(src)}")
        partes = _split_single_pdf_por_paginas(src, destino_final, qualidade,
                                               start_ind=len(partes_saida) + 1,
                                               remover_brancos=False, notify=notify, cancel_flag=cancel_flag)
        partes_saida.extend(partes)

    # agrupa os demais (já comprimidos) somando tamanhos
    # usa margem para evitar que o arquivo final ultrapasse por overhead da junção
    budget = LIMITE_MB - MARGEM_MB
    grupos: List[List[str]] = []
    atual: List[str] = []
    soma = 0.0
    for info in validos:
        if soma + info["mb"] <= budget:
            atual.append(info["compressed"]); soma += info["mb"]
        else:
            if atual:
                grupos.append(atual)
            atual = [info["compressed"]]; soma = info["mb"]
    if atual:
        grupos.append(atual)

    # gera cada parte unindo sem recomprimir
    if grupos:
        notify.text("Gerando partes (sem recompressão)")
        notify.set_total(len(grupos)); notify.step_to(0)
        for i, grupo in enumerate(grupos, start=1):
            saida = _nome_parte(destino_final, len(partes_saida) + 1)
            ok = unir_pdfs_sem_recomprimir(grupo, saida)
            if not ok:
                # se por algum motivo ultrapassar, faz “refino” dividindo o grupo
                notify.warn(f"Parte {i}: união direta falhou — tentando dividir grupo.")
                # divide pela metade até dar
                left = grupo
                while left:
                    chunk = left[:max(1, len(left)//2)]
                    left = left[len(chunk):]
                    saida = _nome_parte(destino_final, len(partes_saida) + 1)
                    if unir_pdfs_sem_recomprimir(chunk, saida) and _mb(saida) <= LIMITE_MB:
                        partes_saida.append(saida)
                    else:
                        # fallback final: recomprime esse pedaço via GS
                        if os.path.exists(saida):
                            try: os.remove(saida)
                            except Exception: pass
                        if not gs_comprimir_para_pdf(chunk, saida, qualidade=qualidade, notify=notify):
                            notify.error("Falha ao gerar parte mesmo após divisão/refino.")
                            return []
                        partes_saida.append(saida)
                continue

            # checa tamanho final; se por algum motivo passou, recomprime só esta parte
            if _mb(saida) > LIMITE_MB:
                notify.subtext(f"Ajustando parte {len(partes_saida) + 1}…")
                if not gs_comprimir_para_pdf(grupo, saida, qualidade=qualidade, notify=notify):
                    notify.error("Falha ao ajustar parte.")
                    return []
            partes_saida.append(saida)
            notify.step_to(i)

    return _finalizar_partes_existentes(partes_saida, notify)

def _finalizar_partes_existentes(partes_geradas: List[str], notify: Notifier) -> List[str]:
    if partes_geradas:
        lista_str = "\n".join(f"- {os.path.basename(p)} ({_mb(p):.2f} MB)" for p in partes_geradas)
        notify.info(f"✅ Excedeu {LIMITE_MB:.0f} MB e foi fragmentado em {len(partes_geradas)} parte(s):\n\n{lista_str}")
    return partes_geradas

# ===================== Diálogo de Progresso (UI, com %) =====================

class ProgressDialog:
    def __init__(self, master, title="Processando..."):
        self.top = tk.Toplevel(master)
        self.top.title(title)
        self.top.resizable(False, False)
        self.top.grab_set()  # modal

        frm = tk.Frame(self.top, padx=14, pady=12); frm.pack(fill="both", expand=True)
        self.lbl = tk.Label(frm, text="Iniciando…"); self.lbl.pack(anchor="w")
        self.pb = ttk.Progressbar(frm, orient="horizontal", length=360, mode="determinate", maximum=100); self.pb.pack(pady=(8,2))
        self.lbl_pct = tk.Label(frm, text="0%"); self.lbl_pct.pack(anchor="e", padx=2)
        self.lbl2 = tk.Label(frm, text=""); self.lbl2.pack(anchor="w", pady=(6,0))

        btn_frame = tk.Frame(frm); btn_frame.pack(fill="x", pady=(10, 0))
        self.btn_cancel = tk.Button(btn_frame, text="Cancelar", width=10, command=self._on_cancel); self.btn_cancel.pack(side="right")

        self.queue = Queue()
        self.notifier = Notifier(self.queue)
        self.cancel_flag = threading.Event()
        self.result = None

        self._max = 100
        self._val = 0

        self.top.protocol("WM_DELETE_WINDOW", self._on_cancel)
        self._poll()

    def _on_cancel(self):
        self.cancel_flag.set()
        self.lbl2.config(text="Cancelando...")

    def _render_pct(self):
        try:
            pct = 0 if self._max <= 0 else int((self._val / self._max) * 100)
            pct = max(0, min(100, pct))
            self.lbl_pct.config(text=f"{pct}%")
        except Exception:
            self.lbl_pct.config(text="")

    def _poll(self):
        try:
            while True:
                kind, payload = self.queue.get_nowait()
                if kind == "set_total":
                    total = max(1, int(payload))
                    self._max = total
                    self._val = 0
                    self.pb["maximum"] = total
                    self.pb["value"] = 0
                    self._render_pct()
                elif kind == "step":
                    self._val = int(payload)
                    self.pb["value"] = self._val
                    self._render_pct()
                elif kind == "text":
                    self.lbl.config(text=str(payload))
                elif kind == "subtext":
                    self.lbl2.config(text=str(payload))
                elif kind == "done":
                    self.result = payload
                elif kind == "close":
                    self._finalize()
                    return
        except Empty:
            pass
        self.top.after(80, self._poll)

    def _finalize(self):
        try:
            for kind, msg in self.notifier.msgs:
                if kind == "info":
                    messagebox.showinfo("Sucesso", msg, parent=self.top)
                elif kind == "warn":
                    messagebox.showwarning("Aviso", msg, parent=self.top)
                elif kind == "error":
                    messagebox.showerror("Erro", msg, parent=self.top)
        except Exception:
            pass
        try:
            self.top.grab_release()
        except Exception:
            pass
        self.top.destroy()

# ===================== Interface Gráfica =====================

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Compressor de PDFs")
        self.root.geometry("680x520")
        self.root.resizable(False, False)

        # Centralizar
        self.root.update_idletasks()
        w, h = 680, 520
        x = (self.root.winfo_screenwidth() // 2) - (w // 2)
        y = (self.root.winfo_screenheight() // 2) - (h // 2)
        self.root.geometry(f"{w}x{h}+{x}+{y}")

        try:
            self.root.iconbitmap("OIP.ico")
        except Exception:
            pass

        self.destino = tk.StringVar(value="")
        self.qualidade = tk.StringVar(value="/ebook")
        self.remover_brancos = tk.BooleanVar(value=False)   # DESLIGADO por padrão (mais rápido)
        self.modo_turbo = tk.BooleanVar(value=True)         # LIGADO por padrão (bem mais rápido)

        self._construir_layout()

    def _construir_layout(self):
        frame = tk.Frame(self.root, padx=14, pady=12); frame.pack(fill="both", expand=True)

        tk.Label(frame, text="Arquivos PDF (reordene/remova antes de unir):").grid(row=0, column=0, columnspan=6, sticky="w")

        self.lista = tk.Listbox(frame, selectmode=tk.EXTENDED, width=72, height=16)
        self.lista.grid(row=1, column=0, columnspan=5, sticky="nsew", padx=(0, 8))

        scroll = tk.Scrollbar(frame, orient="vertical", command=self.lista.yview)
        scroll.grid(row=1, column=5, sticky="ns")
        self.lista.config(yscrollcommand=scroll.set)

        col_btns = tk.Frame(frame); col_btns.grid(row=1, column=6, sticky="ns")
        tk.Button(col_btns, text="Adicionar PDFs", width=18, command=self.adicionar_pdfs).pack(pady=(0, 6))
        tk.Button(col_btns, text="Remover sel.", width=18, command=self.remover_selecionados).pack(pady=3)
        tk.Button(col_btns, text="Mover ↑", width=18, command=self.mover_para_cima).pack(pady=3)
        tk.Button(col_btns, text="Mover ↓", width=18, command=self.mover_para_baixo).pack(pady=3)
        tk.Button(col_btns, text="Limpar lista", width=18, command=self.limpar_lista).pack(pady=3)

        # Destino
        destino_frame = tk.Frame(frame); destino_frame.grid(row=2, column=0, columnspan=7, sticky="we", pady=(10, 0))
        destino_frame.columnconfigure(1, weight=1)
        tk.Label(destino_frame, text="Salvar como:").grid(row=0, column=0, sticky="w")
        tk.Entry(destino_frame, textvariable=self.destino).grid(row=0, column=1, sticky="we", padx=6)
        tk.Button(destino_frame, text="Escolher…", command=self.escolher_destino).grid(row=0, column=2)

        # Opções
        opts_frame = tk.Frame(frame); opts_frame.grid(row=3, column=0, columnspan=7, sticky="w", pady=(8, 0))
        tk.Label(opts_frame, text="Qualidade:").pack(side="left")
        tk.OptionMenu(opts_frame, self.qualidade, *QUALIDADES_GS).pack(side="left", padx=6)
        tk.Checkbutton(opts_frame, text="Remover páginas em branco (lento)", variable=self.remover_brancos).pack(side="left", padx=16)
        tk.Checkbutton(opts_frame, text="Modo turbo (mais rápido)", variable=self.modo_turbo).pack(side="left", padx=16)

        # Rodapé
        rodape = tk.Frame(frame); rodape.grid(row=4, column=0, columnspan=7, sticky="e", pady=(16, 0))
        tk.Button(rodape, text="Unir & Comprimir", width=18, command=self.unir_e_comprimir).pack(side="left", padx=(0, 8))
        tk.Button(rodape, text="Somente Unir", width=14, command=self.somente_unir).pack(side="left", padx=(0, 8))
        tk.Button(rodape, text="Sair", width=10, command=self.root.destroy).pack(side="left")

    # --------- Ações Lista ---------
    def adicionar_pdfs(self):
        arquivos = filedialog.askopenfilenames(title="Selecione os arquivos PDF", filetypes=[("Arquivos PDF", "*.pdf")])
        for a in arquivos:
            if a and a.lower().endswith(".pdf"):
                self.lista.insert(tk.END, a)

    def remover_selecionados(self):
        sel = list(self.lista.curselection())
        if not sel: return
        for idx in reversed(sel):
            self.lista.delete(idx)

    def mover_para_cima(self):
        sel = list(self.lista.curselection())
        if not sel or sel[0] == 0: return
        for idx in sel:
            texto = self.lista.get(idx)
            self.lista.delete(idx)
            self.lista.insert(idx - 1, texto)
        self.lista.selection_clear(0, tk.END)
        for idx in [i - 1 for i in sel]:
            self.lista.selection_set(idx)

    def mover_para_baixo(self):
        sel = list(self.lista.curselection())
        if not sel: return
        max_idx = self.lista.size() - 1
        if sel[-1] == max_idx: return
        for idx in reversed(sel):
            texto = self.lista.get(idx)
            self.lista.delete(idx)
            self.lista.insert(idx + 1, texto)
        self.lista.selection_clear(0, tk.END)
        for idx in [i + 1 for i in sel]:
            self.lista.selection_set(idx)

    def limpar_lista(self):
        self.lista.delete(0, tk.END)

    # --------- Destino / Execução ---------
    def escolher_destino(self):
        destino = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF", "*.pdf")], title="Salvar PDF de saída como")
        if destino:
            self.destino.set(destino)

    def _obter_arquivos_da_lista(self) -> List[str]:
        return [self.lista.get(i) for i in range(self.lista.size())]

    def _validar_pronto(self, exige_destino: bool = True) -> bool:
        if self.lista.size() == 0:
            messagebox.showwarning("Atenção", "Adicione pelo menos um PDF à lista.")
            return False
        if exige_destino and not self.destino.get():
            messagebox.showwarning("Atenção", "Escolha o arquivo de destino.")
            return False
        return True

    def somente_unir(self):
        if not self._validar_pronto(exige_destino=True):
            return
        arquivos = self._obter_arquivos_da_lista()
        out = self.destino.get()

        writer = PdfWriter()
        total = 0
        for f in arquivos:
            path_use = limpar_brancos_para_arquivo(f) if self.remover_brancos.get() else f
            try:
                r = PdfReader(path_use)
                for p in r.pages:
                    writer.add_page(p)
                    total += 1
            except Exception as e:
                messagebox.showerror("Erro", f"Falha ao ler {os.path.basename(f)}:\n{e}")
                return
        if total == 0:
            messagebox.showwarning("Aviso", "Nenhuma página encontrada.")
            return
        try:
            with open(out, "wb") as fo:
                writer.write(fo)
            messagebox.showinfo("Sucesso", f"✅ PDF unido salvo em:\n\n{out}")
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao salvar:\n{e}")

    def unir_e_comprimir(self):
        if not self._validar_pronto(exige_destino=True):
            return
        arquivos = self._obter_arquivos_da_lista()
        destino_final = self.destino.get()
        qualidade = self.qualidade.get()
        remover = self.remover_brancos.get()
        turbo = self.modo_turbo.get()

        prog = ProgressDialog(self.root, title="Processando PDFs")
        notify = prog.notifier
        cancel_flag = prog.cancel_flag

        def run_worker():
            try:
                result = processar_com_limite_worker(
                    arquivos, destino_final, qualidade, remover, turbo, notify, cancel_flag
                )
                notify.done(result)
            finally:
                notify.close()

        threading.Thread(target=run_worker, daemon=True).start()

# ===================== Execução =====================

if __name__ == "__main__":
    root = tk.Tk()
    App(root)
    root.mainloop()
