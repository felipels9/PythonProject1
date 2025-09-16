import os
import sys
from tkinter import (
    Tk, Frame, Button, Listbox, SINGLE, EXTENDED, END, filedialog,
    messagebox, Label, Entry, StringVar, Spinbox, Checkbutton, IntVar, BOTH, LEFT, RIGHT, X, Y, Scrollbar, VERTICAL
)
from PIL import Image

# ------------------------ Utilidades ------------------------

IMG_EXTS = (".jpg", ".jpeg", ".png", ".bmp", ".tif", ".tiff", ".webp")


def is_image(path: str) -> bool:
    return path.lower().endswith(IMG_EXTS)


def natural_sort_key(s):
    # Mantém ordens humanas (001 < 10 < 100, etc.)
    import re
    return [int(text) if text.isdigit() else text.lower() for text in re.split(r'(\d+)', s)]


def mm_to_px(mm: float, dpi: int) -> int:
    # 25.4 mm = 1 polegada
    return int(round((mm / 25.4) * dpi))


def build_a4_canvas(img: Image.Image, dpi: int, margin_mm: float, auto_orient: bool) -> Image.Image:
    """
    Cria uma "folha" A4 branca e centraliza a imagem ajustando por margem.
    Se auto_orient=True, decide entre A4 retrato (8.27x11.69) e paisagem (11.69x8.27) para melhor aproveitamento.
    """
    # Tamanhos A4 em polegadas
    A4_W_IN, A4_H_IN = 8.27, 11.69

    # Se auto_orient, escolhe retrato ou paisagem conforme aspecto da imagem
    if auto_orient:
        img_w, img_h = img.size
        if img_w > img_h:
            # paisagem
            page_w_px = int(round(11.69 * dpi))
            page_h_px = int(round(8.27 * dpi))
        else:
            # retrato
            page_w_px = int(round(8.27 * dpi))
            page_h_px = int(round(11.69 * dpi))
    else:
        # padrão: retrato
        page_w_px = int(round(A4_W_IN * dpi))
        page_h_px = int(round(A4_H_IN * dpi))

    margin_px = mm_to_px(margin_mm, dpi)
    box_w = max(1, page_w_px - 2 * margin_px)
    box_h = max(1, page_h_px - 2 * margin_px)

    # Ajusta tamanho da imagem para caber no box mantendo proporção
    img = img.convert("RGB")
    img_w, img_h = img.size
    scale = min(box_w / img_w, box_h / img_h)
    new_w = max(1, int(img_w * scale))
    new_h = max(1, int(img_h * scale))
    img_resized = img.resize((new_w, new_h), Image.LANCZOS)

    # Cria folha branca
    page = Image.new("RGB", (page_w_px, page_h_px), "white")
    offset_x = (page_w_px - new_w) // 2
    offset_y = (page_h_px - new_h) // 2
    page.paste(img_resized, (offset_x, offset_y))
    return page


def open_image_safe(path: str) -> Image.Image:
    img = Image.open(path)
    # Algumas imagens têm modo P/LA/RGBA; padronizamos mais adiante
    return img


def save_multipage_pdf(images, out_path: str, resolution: int = 300):
    """
    Salva lista de PIL.Image em um PDF multipágina.
    """
    if not images:
        raise ValueError("Nenhuma imagem para salvar.")
    first, rest = images[0], images[1:]
    first.save(out_path, "PDF", resolution=resolution, save_all=True, append_images=rest)


# ------------------------ Interface ------------------------

class App:
    def __init__(self, master: Tk):
        self.master = master
        master.title("Imagens → PDF (A4 opcional)")

        # Estado
        self.images = []  # lista de caminhos
        self.out_name_var = StringVar(value="saida.pdf")
        self.dpi_var = StringVar(value="300")
        self.margin_var = StringVar(value="10")  # mm
        self.fit_a4_var = IntVar(value=1)
        self.auto_orient_var = IntVar(value=1)

        # Topo: botões de arquivos
        top = Frame(master)
        top.pack(fill=X, padx=8, pady=(8, 4))

        Button(top, text="Adicionar imagens…", command=self.add_images).pack(side=LEFT, padx=(0, 6))
        Button(top, text="Remover selecionadas", command=self.remove_selected).pack(side=LEFT, padx=6)
        Button(top, text="Limpar lista", command=self.clear_list).pack(side=LEFT, padx=6)

        # Lista com scrollbar
        mid = Frame(master)
        mid.pack(fill=BOTH, expand=True, padx=8, pady=4)

        self.listbox = Listbox(mid, selectmode=EXTENDED)
        self.listbox.pack(side=LEFT, fill=BOTH, expand=True)

        sb = Scrollbar(mid, orient=VERTICAL, command=self.listbox.yview)
        sb.pack(side=RIGHT, fill=Y)
        self.listbox.config(yscrollcommand=sb.set)

        # Controles de ordem
        order = Frame(master)
        order.pack(fill=X, padx=8, pady=4)
        Button(order, text="↑ Subir", command=self.move_up).pack(side=LEFT, padx=(0, 6))
        Button(order, text="↓ Descer", command=self.move_down).pack(side=LEFT, padx=6)

        # Opções
        opts = Frame(master, bd=1, relief="groove")
        opts.pack(fill=X, padx=8, pady=(6, 4))

        Label(opts, text="Nome do PDF:").pack(side=LEFT, padx=(6, 4))
        Entry(opts, textvariable=self.out_name_var, width=28).pack(side=LEFT, padx=(0, 12))

        Label(opts, text="DPI:").pack(side=LEFT, padx=(0, 4))
        Spinbox(opts, from_=72, to=600, increment=1, width=6, textvariable=self.dpi_var).pack(side=LEFT, padx=(0, 12))

        Label(opts, text="Margem (mm):").pack(side=LEFT, padx=(0, 4))
        Spinbox(opts, from_=0, to=50, increment=1, width=6, textvariable=self.margin_var).pack(side=LEFT, padx=(0, 12))

        Checkbutton(opts, text="Ajustar para A4", variable=self.fit_a4_var).pack(side=LEFT, padx=(0, 12))
        Checkbutton(opts, text="Auto-orientação", variable=self.auto_orient_var).pack(side=LEFT, padx=(0, 12))

        # Ações finais
        bottom = Frame(master)
        bottom.pack(fill=X, padx=8, pady=(4, 8))

        Button(bottom, text="Salvar como ÚNICO PDF…", command=self.export_single_pdf, height=2).pack(side=LEFT, padx=(0, 6), fill=X, expand=True)
        Button(bottom, text="Salvar UM PDF por imagem…", command=self.export_many_pdfs, height=2).pack(side=LEFT, padx=6, fill=X, expand=True)

        # Dica de uso
        tip = Frame(master)
        tip.pack(fill=X, padx=8, pady=(0, 10))
        Label(tip, text="Dica: você pode reordenar as imagens com ↑/↓. O PDF segue a ordem da lista.").pack(side=LEFT)

    # ---------- Handlers ----------

    def add_images(self):
        paths = filedialog.askopenfilenames(
            title="Selecione as imagens",
            filetypes=[
                ("Imagens", "*.jpg;*.jpeg;*.png;*.bmp;*.tif;*.tiff;*.webp"),
                ("Todos os arquivos", "*.*"),
            ]
        )
        if not paths:
            return

        # Filtra imagens válidas e remove duplicadas
        new_paths = [p for p in paths if is_image(p)]
        # Ordena naturalmente por nome (para conveniência)
        new_paths.sort(key=natural_sort_key)

        # Sugere nome de saída com base na primeira imagem
        if self.out_name_var.get() == "saida.pdf" and new_paths:
            base = os.path.splitext(os.path.basename(new_paths[0]))[0]
            self.out_name_var.set(f"{base}_convertido.pdf")

        added = 0
        for p in new_paths:
            if p not in self.images:
                self.images.append(p)
                self.listbox.insert(END, p)
                added += 1

        if added == 0 and new_paths:
            messagebox.showinfo("Aviso", "As imagens selecionadas já estavam na lista.")

    def remove_selected(self):
        selection = list(self.listbox.curselection())
        if not selection:
            messagebox.showinfo("Remover", "Nenhum item selecionado.")
            return
        # Remover de baixo para cima para não deslocar índices
        for idx in reversed(selection):
            self.listbox.delete(idx)
            del self.images[idx]

    def clear_list(self):
        if not self.images:
            return
        if messagebox.askyesno("Confirmar", "Limpar TODA a lista de imagens?"):
            self.listbox.delete(0, END)
            self.images.clear()

    def move_up(self):
        selection = list(self.listbox.curselection())
        if not selection:
            return
        if selection[0] == 0:
            return  # topo

        for idx in selection:
            # swap idx e idx-1
            self.images[idx-1], self.images[idx] = self.images[idx], self.images[idx-1]

        # Atualiza listbox preservando seleção
        self.refresh_listbox(select=[i-1 for i in selection])

    def move_down(self):
        selection = list(self.listbox.curselection())
        if not selection:
            return
        if selection[-1] == len(self.images) - 1:
            return  # final

        for idx in reversed(selection):
            self.images[idx+1], self.images[idx] = self.images[idx], self.images[idx+1]

        # Atualiza listbox preservando seleção
        self.refresh_listbox(select=[i+1 for i in selection])

    def refresh_listbox(self, select=None):
        self.listbox.delete(0, END)
        for p in self.images:
            self.listbox.insert(END, p)
        if select:
            for i in select:
                self.listbox.selection_set(i)
            self.listbox.see(select[0])

    def get_settings(self):
        try:
            dpi = int(self.dpi_var.get())
            if dpi < 72 or dpi > 1200:
                raise ValueError
        except Exception:
            raise ValueError("DPI inválido. Use um número entre 72 e 1200.")

        try:
            margin = float(self.margin_var.get())
            if margin < 0 or margin > 100:
                raise ValueError
        except Exception:
            raise ValueError("Margem inválida (mm). Use um número entre 0 e 100.")

        fit_a4 = bool(self.fit_a4_var.get())
        auto_orient = bool(self.auto_orient_var.get())
        return dpi, margin, fit_a4, auto_orient

    def export_single_pdf(self):
        if not self.images:
            messagebox.showwarning("Atenção", "Adicione ao menos uma imagem.")
            return

        default_name = self.out_name_var.get().strip() or "saida.pdf"
        if not default_name.lower().endswith(".pdf"):
            default_name += ".pdf"

        out_path = filedialog.asksaveasfilename(
            title="Salvar PDF",
            defaultextension=".pdf",
            initialfile=default_name,
            filetypes=[("PDF", "*.pdf")]
        )
        if not out_path:
            return

        try:
            dpi, margin_mm, fit_a4, auto_orient = self.get_settings()
        except ValueError as e:
            messagebox.showerror("Erro", str(e))
            return

        try:
            pages = []
            for p in self.images:
                img = open_image_safe(p)
                if fit_a4:
                    page = build_a4_canvas(img, dpi=dpi, margin_mm=margin_mm, auto_orient=auto_orient)
                    pages.append(page)
                else:
                    pages.append(img.convert("RGB"))

            save_multipage_pdf(pages, out_path, resolution=dpi)
            messagebox.showinfo("Sucesso", f"PDF criado:\n{out_path}")
        except Exception as e:
            messagebox.showerror("Erro ao gerar PDF", f"Ocorreu um erro:\n{e}")

    def export_many_pdfs(self):
        if not self.images:
            messagebox.showwarning("Atenção", "Adicione ao menos uma imagem.")
            return

        out_dir = filedialog.askdirectory(title="Selecione a pasta para salvar os PDFs")
        if not out_dir:
            return

        try:
            dpi, margin_mm, fit_a4, auto_orient = self.get_settings()
        except ValueError as e:
            messagebox.showerror("Erro", str(e))
            return

        ok, fail = 0, 0
        for p in self.images:
            try:
                base = os.path.splitext(os.path.basename(p))[0]
                out_path = os.path.join(out_dir, f"{base}.pdf")
                img = open_image_safe(p)
                if fit_a4:
                    page = build_a4_canvas(img, dpi=dpi, margin_mm=margin_mm, auto_orient=auto_orient)
                else:
                    page = img.convert("RGB")
                page.save(out_path, "PDF", resolution=dpi)
                ok += 1
            except Exception:
                fail += 1

        if fail == 0:
            messagebox.showinfo("Sucesso", f"PDF(s) criado(s) com sucesso na pasta:\n{out_dir}")
        else:
            messagebox.showwarning("Concluído com avisos", f"Concluído: {ok} sucesso(s), {fail} falha(s).\nPasta: {out_dir}")


def main():
    # Correção de alta DPI no Windows (opcional, não quebra em outros SOs)
    if sys.platform.startswith("win"):
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass

    root = Tk()
    root.geometry("820x520")
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
