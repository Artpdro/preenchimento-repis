"""Interface desktop - tema escuro estilo Proton Drive, janela 1366x768.

Abas: Novo Certificado, Pesquisa, Importar Excel, Preenchimento em Lote,
Historico e Sobre. Banco de dados SQLite (database.py) e importacao
Excel (excel_importer.py) em modulos separados.
"""
import os
import subprocess
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from config import LOGO_PNG, OUTPUT_DIR, WINDOW_HEIGHT, WINDOW_WIDTH
import database
import excel_importer
import history
from pdf_handler import fill_pdf
from repis_generator import generate_repis_code
from validators import format_cnpj, is_valid_cnpj, validate_empresa

THEME = {"bg": "#191A23", "sidebar": "#12131B", "card": "#22242F",
         "border": "#34374A", "text": "#E8EAF2", "muted": "#9AA1B5",
         "entry": "#2B2D3C", "accent": "#1E5B96", "accent_h": "#3D76AC",
         "accent_d": "#123B61", "gold": "#C2964B", "gold_h": "#D2B075",
         "ok": "#4CAF7D", "err": "#E05A5A"}
FONT = "Segoe UI"


def open_path(path):
    try:
        if sys.platform.startswith("win"):
            os.startfile(str(path))
        elif sys.platform == "darwin":
            subprocess.Popen(["open", str(path)])
        else:
            subprocess.Popen(["xdg-open", str(path)])
    except Exception:
        pass


class NavButton(tk.Button):
    def __init__(self, master, text, command, **kw):
        super().__init__(master, text=text, command=command, font=(FONT, 11),
                         anchor="w", padx=18, pady=9, bd=0, relief="flat",
                         cursor="hand2", bg=THEME["sidebar"], fg=THEME["muted"],
                         activebackground=THEME["card"],
                         activeforeground=THEME["text"], **kw)
        self._active = False
        self.bind("<Enter>", lambda e: self.config(bg=self._hover()))
        self.bind("<Leave>", lambda e: self.config(bg=self._bg()))

    def _bg(self):
        return THEME["accent_d"] if self._active else THEME["sidebar"]

    def _hover(self):
        return THEME["card"] if not self._active else THEME["accent_d"]

    def set_active(self, a):
        self._active = a
        self.config(bg=self._bg(), fg=THEME["gold"] if a else THEME["muted"])


class Header(tk.Frame):
    def __init__(self, master, title, subtitle=""):
        super().__init__(master, bg=THEME["bg"])
        tk.Label(self, text=title, font=(FONT, 20, "bold"),
                 bg=THEME["bg"], fg=THEME["text"]).pack(anchor="w")
        if subtitle:
            tk.Label(self, text=subtitle, font=(FONT, 11),
                     bg=THEME["bg"], fg=THEME["muted"]).pack(anchor="w", pady=(4, 0))
        tk.Frame(self, bg=THEME["border"], height=1).pack(fill="x", pady=(14, 0))


class Card(tk.Frame):
    def __init__(self, master, **kw):
        super().__init__(master, bg=THEME["card"],
                         highlightbackground=THEME["border"],
                         highlightthickness=1, **kw)


def style_entry(e):
    e.config(bg=THEME["entry"], fg=THEME["text"], insertbackground=THEME["text"],
             relief="flat", font=(FONT, 11), highlightthickness=1,
             highlightbackground=THEME["border"], highlightcolor=THEME["accent"])


def action_btn(master, text, cmd, accent=False):
    cor = THEME["accent"] if accent else THEME["card"]
    cor_h = THEME["accent_h"] if accent else THEME["accent_d"]
    b = tk.Button(master, text=text, command=cmd, font=(FONT, 10, "bold" if accent else "normal"),
                  bg=cor, fg="#FFFFFF" if accent else THEME["text"],
                  activebackground=cor_h, activeforeground="#FFFFFF" if accent else THEME["gold"],
                  bd=0, cursor="hand2", padx=18, pady=8,
                  highlightbackground=THEME["border"], highlightthickness=0 if accent else 1)
    b.bind("<Enter>", lambda e: b.config(bg=cor_h))
    b.bind("<Leave>", lambda e: b.config(bg=cor))
    return b


def make_tree(master, columns):
    st = ttk.Style(master)
    st.theme_use("clam")
    st.configure("Dark.Treeview", background=THEME["card"],
                 fieldbackground=THEME["card"], foreground=THEME["text"],
                 rowheight=30, bordercolor=THEME["border"], font=(FONT, 10))
    st.configure("Dark.Treeview.Heading", background=THEME["sidebar"],
                 foreground=THEME["muted"], relief="flat",
                 font=(FONT, 9, "bold"), padding=6)
    st.map("Dark.Treeview", background=[("selected", THEME["accent_d"])],
           foreground=[("selected", THEME["gold"])])
    tree = ttk.Treeview(master, style="Dark.Treeview",
                        columns=[c[0] for c in columns], show="headings")
    for cid, label, w in columns:
        tree.heading(cid, text=label)
        tree.column(cid, width=w, anchor="w")
    tree.pack(side="left", fill="both", expand=True, padx=(8, 0), pady=8)
    sc = ttk.Scrollbar(master, orient="vertical", command=tree.yview)
    sc.pack(side="right", fill="y", pady=8, padx=(0, 8))
    tree.configure(yscrollcommand=sc.set)
    return tree


# ============================================================ NOVO CERTIFICADO
class NovoCertificadoView(tk.Frame):
    def __init__(self, master, app):
        super().__init__(master, bg=THEME["bg"])
        self.app = app
        Header(self, "Novo Certificado",
               "Preencha os dados para gerar o PDF de adesao ao REPIS").pack(
            fill="x", padx=40, pady=(32, 24))
        form = Card(self, padx=32, pady=28)
        form.pack(fill="x", padx=40)

        tk.Label(form, text="NOME DA EMPRESA", font=(FONT, 9, "bold"),
                 bg=THEME["card"], fg=THEME["muted"]).pack(anchor="w")
        self.entry_nome = tk.Entry(form, width=70)
        style_entry(self.entry_nome)
        self.entry_nome.pack(anchor="w", pady=(6, 20), ipady=6)

        tk.Label(form, text="CNPJ", font=(FONT, 9, "bold"),
                 bg=THEME["card"], fg=THEME["muted"]).pack(anchor="w")
        vcmd = (self.register(lambda p: len("".join(ch for ch in p if ch.isdigit())) <= 14
                              or p == ""), "%P")
        self.entry_cnpj = tk.Entry(form, width=25, validate="key", validatecommand=vcmd)
        style_entry(self.entry_cnpj)
        self.entry_cnpj.pack(anchor="w", pady=(6, 26), ipady=6)
        self.entry_cnpj.bind("<FocusOut>", self._mask)
        self.entry_cnpj.bind("<KeyRelease>", self._mask_live)

        action_btn(form, "GERAR PDF", self.gerar, accent=True).pack(anchor="w")

        self.lbl_result = tk.Label(form, text="", font=(FONT, 10),
                                   bg=THEME["card"], fg=THEME["gold"],
                                   justify="left", wraplength=820)
        self.lbl_result.pack(anchor="w", pady=(18, 0))
        for e in (self.entry_nome, self.entry_cnpj):
            e.bind("<Return>", lambda _ev: self.gerar())

    def preencher(self, nome, cnpj):
        """Preenche o formulario a partir de dados externos (ex.: Pesquisa)."""
        self.entry_nome.delete(0, tk.END)
        self.entry_nome.insert(0, nome or "")
        self.entry_cnpj.delete(0, tk.END)
        self.entry_cnpj.insert(0, cnpj or "")

    def _mask(self, _evt=None):
        d = "".join(ch for ch in self.entry_cnpj.get() if ch.isdigit())[:14]
        self.entry_cnpj.delete(0, tk.END)
        self.entry_cnpj.insert(0, format_cnpj(d) if len(d) == 14 else d)

    def _mask_live(self, _evt=None):
        d = "".join(ch for ch in self.entry_cnpj.get() if ch.isdigit())[:14]
        if d and len(d) in (2, 5, 8, 12):
            self._mask()

    def gerar(self):
        nome = self.entry_nome.get().strip()
        cnpj_raw = self.entry_cnpj.get().strip()
        erro = validate_empresa(nome)
        if erro:
            messagebox.showwarning("Validacao", erro)
            return
        if not is_valid_cnpj(cnpj_raw):
            messagebox.showwarning("Validacao", "CNPJ invalido.")
            return
        path = gerar_e_registrar(nome, format_cnpj(cnpj_raw), self.app)
        if path:
            self.lbl_result.config(text=f"Codigo REPIS: {self.app.ultimo_codigo}"
                                        f"\nArquivo: {path}")
            if messagebox.askyesno("Sucesso",
                    f"PDF gerado!\nCodigo REPIS: {self.app.ultimo_codigo}\n\n"
                    f"Deseja abrir a pasta de destino?"):
                open_path(OUTPUT_DIR)


def gerar_e_registrar(nome: str, cnpj: str, app) -> "Path | None":
    """Fluxo comum: gera codigo + PDF, grava historico e atualiza cadastro."""
    from pathlib import Path
    codigo = generate_repis_code()
    try:
        path = fill_pdf(nome_empresa=nome, cnpj=cnpj, codigo_repis=codigo)
    except Exception as exc:
        messagebox.showerror("Erro", f"Falha ao gerar o PDF:\n{exc}")
        return None
    history.add_record(nome, cnpj, codigo, path)
    try:
        database.upsert_empresa(nome, cnpj, codigo)
    except ValueError:
        pass  # CNPJ fora do padrao nao bloqueia a emissao
    database.update_codigo_repis(cnpj, codigo)
    app.ultimo_codigo = codigo
    app.set_status(f"Certificado gerado: {codigo}")
    app.refresh_lists()
    return path


# ============================================================ PESQUISA
class PesquisaView(tk.Frame):
    COLS = (("nome", "Empresa", 330), ("cnpj", "CNPJ", 160),
            ("codigo", "Codigo REPIS", 140))

    def __init__(self, master, app):
        super().__init__(master, bg=THEME["bg"])
        self.app = app
        Header(self, "Pesquisa",
               "Localize empresas por CNPJ ou Codigo REPIS (busca parcial)").pack(
            fill="x", padx=40, pady=(32, 24))

        bar = Card(self, padx=20, pady=14)
        bar.pack(fill="x", padx=40, pady=(0, 12))
        self.entry_termo = tk.Entry(bar, width=45)
        style_entry(self.entry_termo)
        self.entry_termo.pack(side="left", ipady=6, padx=(4, 10))
        self.entry_termo.bind("<Return>", lambda _e: self.buscar())
        action_btn(bar, "Buscar", self.buscar, accent=True).pack(side="left")
        self.lbl_contagem = tk.Label(bar, text="", font=(FONT, 10),
                                     bg=THEME["card"], fg=THEME["muted"])
        self.lbl_contagem.pack(side="left", padx=14)

        card = Card(self)
        card.pack(fill="both", expand=True, padx=40, pady=(0, 12))
        self.tree = make_tree(card, self.COLS)
        self.tree.bind("<Double-1>", lambda _e: self.selecionar())

        acoes = tk.Frame(self, bg=THEME["bg"])
        acoes.pack(fill="x", padx=40, pady=(0, 0))
        action_btn(acoes, "Selecionar (usar no formulario)", self.selecionar).pack(side="left", padx=(0, 10))
        action_btn(acoes, "Preencher PDF direto", self.preencher_pdf).pack(side="left", padx=(0, 10))
        action_btn(acoes, "Limpar", self.limpar).pack(side="left")
        self.resultados = []

    def buscar(self):
        termo = self.entry_termo.get().strip()
        self.resultados = database.search_empresas(termo) if termo else database.list_empresas()
        self._render()

    def _render(self):
        self.tree.delete(*self.tree.get_children())
        for r in self.resultados:
            self.tree.insert("", "end", values=(
                r.get("nome", ""), r.get("cnpj", ""), r.get("codigo_repis", "")))
        self.lbl_contagem.config(text=f"{len(self.resultados)} resultado(s)")

    def _selecionado(self):
        sel = self.tree.selection()
        if not sel:
            return None
        i = self.tree.index(sel[0])
        return self.resultados[i] if i < len(self.resultados) else None

    def selecionar(self):
        reg = self._selecionado()
        if not reg:
            messagebox.showinfo("Pesquisa", "Selecione uma empresa na lista.")
            return
        self.app.views["novo"].preencher(reg.get("nome", ""), reg.get("cnpj", ""))
        self.app.set_status(f"Empresa selecionada: {reg.get('nome', '')}")
        self.app.show("novo")

    def preencher_pdf(self):
        reg = self._selecionado()
        if not reg:
            messagebox.showinfo("Pesquisa", "Selecione uma empresa na lista.")
            return
        path = gerar_e_registrar(reg.get("nome", ""), reg.get("cnpj", ""), self.app)
        if path and messagebox.askyesno("Sucesso",
                f"PDF gerado!\nCodigo REPIS: {self.app.ultimo_codigo}\n\nAbrir pasta?"):
            open_path(OUTPUT_DIR)

    def limpar(self):
        self.entry_termo.delete(0, tk.END)
        self.resultados = []
        self._render()

    def refresh(self):
        self.buscar()


# ============================================================ IMPORTAR EXCEL
class ImportarView(tk.Frame):
    COLS = (("linha", "Linha", 60), ("nome", "Empresa", 330),
            ("cnpj", "CNPJ", 160), ("codigo", "Codigo REPIS", 130))

    def __init__(self, master, app):
        super().__init__(master, bg=THEME["bg"])
        self.app = app
        self.arquivo = None
        self.preview = None
        Header(self, "Importar Empresas (Excel)",
               "Selecione um arquivo .xlsx com as colunas: CNPJ e Nome da Empresa").pack(
            fill="x", padx=40, pady=(32, 24))

        bar = Card(self, padx=20, pady=14)
        bar.pack(fill="x", padx=40, pady=(0, 12))
        action_btn(bar, "Selecionar arquivo .xlsx", self.escolher, accent=True).pack(side="left")
        self.lbl_arq = tk.Label(bar, text="Nenhum arquivo selecionado", font=(FONT, 10),
                                bg=THEME["card"], fg=THEME["muted"])
        self.lbl_arq.pack(side="left", padx=14)

        card = Card(self)
        card.pack(fill="both", expand=True, padx=40, pady=(0, 12))
        self.tree = make_tree(card, self.COLS)
        self.lbl_prev = tk.Label(self, text="", font=(FONT, 10),
                                 bg=THEME["bg"], fg=THEME["muted"])
        self.lbl_prev.pack(anchor="w", padx=40, pady=(0, 8))

        acoes = tk.Frame(self, bg=THEME["bg"])
        acoes.pack(fill="x", padx=40, pady=(0, 0))
        self.btn_importar = action_btn(acoes, "Importar para o banco", self.importar)
        self.btn_importar.pack(side="left", padx=(0, 10))
        action_btn(acoes, "Limpar", self.limpar).pack(side="left")

    def escolher(self):
        path = filedialog.askopenfilename(
            title="Selecione a planilha Excel",
            filetypes=[("Planilha Excel", "*.xlsx")])
        if not path:
            return
        self.arquivo = path
        self.lbl_arq.config(text=os.path.basename(path))
        pv = excel_importer.read_preview(path)
        self.preview = pv
        self.tree.delete(*self.tree.get_children())
        if pv["erros"]:
            messagebox.showerror("Arquivo invalido", "\n".join(pv["erros"]))
            self.btn_importar.config(state="disabled")
            return
        for reg in pv["linhas"]:
            self.tree.insert("", "end", values=(
                reg.get("_linha", ""), reg.get("nome", ""),
                reg.get("cnpj", ""), reg.get("codigo_repis", "")))
        self.lbl_prev.config(
            text=f"Colunas: {', '.join(pv['colunas'])} | "
                 f"Prévia: {len(pv['linhas'])} linha(s) (amostra). "
                 f"Total cadastrado no banco: {database.count_empresas()}")
        self.btn_importar.config(state="normal")

    def importar(self):
        if not self.arquivo:
            return
        rel = excel_importer.importar(self.arquivo)
        msg = (f"Processadas: {rel['total']} linha(s)\n\n"
               f"Novas empresas: {rel['novos']}\n"
               f"Atualizadas: {rel['atualizados']}\n"
               f"Ignoradas: {len(rel['ignorados'])}")
        if rel["ignorados"]:
            msg += "\n\nDetalhes:\n" + "\n".join(rel["ignorados"][:15])
            if len(rel["ignorados"]) > 15:
                msg += f"\n... e mais {len(rel['ignorados']) - 15}."
        messagebox.showinfo("Resultado da importacao", msg)
        self.app.set_status(f"Importacao: {rel['novos']} novas, "
                            f"{rel['atualizados']} atualizadas")
        self.app.refresh_lists()

    def limpar(self):
        self.arquivo = None
        self.preview = None
        self.tree.delete(*self.tree.get_children())
        self.lbl_arq.config(text="Nenhum arquivo selecionado")
        self.lbl_prev.config(text="")
        self.btn_importar.config(state="disabled")


# ============================================================ PREENCHIMENTO EM LOTE
class LoteView(tk.Frame):
    COLS = (("nome", "Empresa", 340), ("cnpj", "CNPJ", 160),
            ("codigo", "Codigo REPIS", 140))

    def __init__(self, master, app):
        super().__init__(master, bg=THEME["bg"])
        self.app = app
        Header(self, "Preenchimento em Lote",
               "Empresas cadastradas no banco de dados").pack(
            fill="x", padx=40, pady=(32, 24))
        card = Card(self)
        card.pack(fill="both", expand=True, padx=40, pady=(0, 12))
        self.tree = make_tree(card, self.COLS)
        self.lbl_contagem = tk.Label(self, text="", font=(FONT, 10),
                                     bg=THEME["bg"], fg=THEME["muted"])
        self.lbl_contagem.pack(anchor="w", padx=40, pady=(0, 8))
        acoes = tk.Frame(self, bg=THEME["bg"])
        acoes.pack(fill="x", padx=40, pady=(0, 0))
        action_btn(acoes, "Preencher PDF (selecionada)", self.preencher, accent=True).pack(side="left", padx=(0, 10))
        action_btn(acoes, "Atualizar lista", self.refresh).pack(side="left", padx=(0, 10))
        action_btn(acoes, "Abrir pasta de saida", lambda: open_path(OUTPUT_DIR)).pack(side="left")
        self.empresas = []

    def refresh(self):
        self.empresas = database.list_empresas()
        self.tree.delete(*self.tree.get_children())
        for r in self.empresas:
            self.tree.insert("", "end", values=(
                r.get("nome", ""), r.get("cnpj", ""), r.get("codigo_repis", "")))
        self.lbl_contagem.config(text=f"{len(self.empresas)} empresa(s) cadastrada(s)")

    def preencher(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Lote", "Selecione uma empresa na lista.")
            return
        i = self.tree.index(sel[0])
        if i >= len(self.empresas):
            return
        reg = self.empresas[i]
        path = gerar_e_registrar(reg.get("nome", ""), reg.get("cnpj", ""), self.app)
        if path and messagebox.askyesno("Sucesso",
                f"PDF gerado!\n{reg.get('nome', '')}\nCodigo REPIS: {self.app.ultimo_codigo}\n\n"
                f"Abrir pasta de destino?"):
            open_path(OUTPUT_DIR)


# ============================================================ HISTORICO
class HistoricoView(tk.Frame):
    COLS = (("data", "Data/Hora", 140), ("empresa", "Empresa", 300),
            ("cnpj", "CNPJ", 150), ("codigo", "Codigo REPIS", 130))

    def __init__(self, master, app):
        super().__init__(master, bg=THEME["bg"])
        self.app = app
        Header(self, "Historico",
               "Certificados ja emitidos neste computador").pack(
            fill="x", padx=40, pady=(32, 24))
        card = Card(self)
        card.pack(fill="both", expand=True, padx=40, pady=(0, 12))
        self.tree = make_tree(card, self.COLS)
        self.tree.bind("<Double-1>", lambda _e: self.abrir())
        acoes = tk.Frame(self, bg=THEME["bg"])
        acoes.pack(fill="x", padx=40, pady=(0, 0))
        action_btn(acoes, "Atualizar", self.refresh).pack(side="left", padx=(0, 10))
        action_btn(acoes, "Abrir PDF selecionado", self.abrir).pack(side="left", padx=(0, 10))
        action_btn(acoes, "Abrir pasta de saida", lambda: open_path(OUTPUT_DIR)).pack(side="left")

    def refresh(self):
        self.tree.delete(*self.tree.get_children())
        for r in history.load_history():
            self.tree.insert("", "end", values=(
                r.get("data", ""), r.get("empresa", ""),
                r.get("cnpj", ""), r.get("codigo", "")))

    def abrir(self):
        sel = self.tree.selection()
        if not sel:
            return
        regs = history.load_history()
        i = self.tree.index(sel[0])
        if i < len(regs):
            open_path(regs[i].get("arquivo", OUTPUT_DIR))


# ============================================================ SOBRE
class SobreView(tk.Frame):
    def __init__(self, master, app):
        super().__init__(master, bg=THEME["bg"])
        Header(self, "Sobre", "").pack(fill="x", padx=40, pady=(32, 24))
        card = Card(self, padx=32, pady=28)
        card.pack(fill="x", padx=40)
        tk.Label(card, font=(FONT, 11), bg=THEME["card"], fg=THEME["text"],
                 justify="left", text=(
            "Gerador de Certificado de Adesao ao REPIS\n"
            "Sindnorte PE - Sindicato Empresarial do Sistema Comercio\n\n"
            "Preenchimento automatico do certificado PDF, geracao do Codigo\n"
            "REPIS, validacao de CNPJ, cadastro em banco SQLite, pesquisa,\n"
            "importacao via Excel (.xlsx), preenchimento em lote e historico.\n\n"
            "Versao 3.0 - Tema escuro com cores da marca.")).pack(anchor="w")


# ============================================================ APP
class RepisApp(tk.Tk):
    SIDEBAR_W = 260

    def __init__(self):
        super().__init__()
        self.title("Gerador de REPIS - Sindnorte PE")
        self.geometry(f"{WINDOW_WIDTH}x{WINDOW_HEIGHT}")
        self.resizable(False, False)
        self.configure(bg=THEME["bg"])
        self.ultimo_codigo = ""
        self._center()
        self._sidebar()
        self.container = tk.Frame(self, bg=THEME["bg"])
        self.container.place(x=self.SIDEBAR_W, y=0,
                             width=WINDOW_WIDTH - self.SIDEBAR_W,
                             height=WINDOW_HEIGHT)
        self.views = {
            "novo": NovoCertificadoView(self.container, self),
            "pesquisa": PesquisaView(self.container, self),
            "importar": ImportarView(self.container, self),
            "lote": LoteView(self.container, self),
            "historico": HistoricoView(self.container, self),
            "sobre": SobreView(self.container, self),
        }
        for v in self.views.values():
            v.place(relwidth=1, relheight=1)
            v.lower()
        self.views["importar"].btn_importar.config(state="disabled")
        self.show("novo")

    def _center(self):
        self.update_idletasks()
        x = (self.winfo_screenwidth() - WINDOW_WIDTH) // 2
        y = (self.winfo_screenheight() - WINDOW_HEIGHT) // 2
        self.geometry(f"{WINDOW_WIDTH}x{WINDOW_HEIGHT}+{max(x, 0)}+{max(y, 0)}")

    def _sidebar(self):
        sb = tk.Frame(self, bg=THEME["sidebar"], width=self.SIDEBAR_W)
        sb.place(x=0, y=0, width=self.SIDEBAR_W, height=WINDOW_HEIGHT)
        sb.pack_propagate(False)
        lf = tk.Frame(sb, bg=THEME["sidebar"])
        lf.pack(fill="x", pady=(24, 18), padx=20)
        try:
            self.logo_img = tk.PhotoImage(file=str(LOGO_PNG))
            tk.Label(lf, image=self.logo_img, bg=THEME["sidebar"]).pack(anchor="w")
        except Exception:
            tk.Label(lf, text="Sindnorte PE", font=(FONT, 16, "bold"),
                     bg=THEME["sidebar"], fg=THEME["gold"]).pack(anchor="w")
        tk.Frame(sb, bg=THEME["border"], height=1).pack(fill="x", padx=16)
        nav = tk.Frame(sb, bg=THEME["sidebar"])
        nav.pack(fill="x", pady=(14, 0))
        self.nav_buttons = {}
        for key, label in (("novo", "Novo Certificado"),
                           ("pesquisa", "Pesquisa"),
                           ("importar", "Importar Excel"),
                           ("lote", "Preenchimento em Lote"),
                           ("historico", "Historico"),
                           ("sobre", "Sobre")):
            b = NavButton(nav, text="   " + label,
                          command=lambda k=key: self.show(k))
            b.pack(fill="x", padx=12, pady=2)
            self.nav_buttons[key] = b
        ft = tk.Frame(sb, bg=THEME["sidebar"])
        ft.pack(side="bottom", fill="x", pady=16, padx=20)
        tk.Label(ft, text="Status", font=(FONT, 8, "bold"),
                 bg=THEME["sidebar"], fg=THEME["muted"]).pack(anchor="w")
        self.lbl_status = tk.Label(ft, text="Pronto", font=(FONT, 9),
                                   bg=THEME["sidebar"], fg=THEME["gold"],
                                   wraplength=self.SIDEBAR_W - 40, justify="left")
        self.lbl_status.pack(anchor="w", pady=(2, 0))

    def show(self, key):
        self.views[key].lift()
        for k, b in self.nav_buttons.items():
            b.set_active(k == key)
        if key in ("pesquisa", "lote"):
            self.views[key].refresh()
        if key == "historico":
            self.views["historico"].refresh()

    def refresh_lists(self):
        if hasattr(self.views["lote"], "refresh"):
            self.views["lote"].refresh()
        if hasattr(self.views["pesquisa"], "refresh"):
            self.views["pesquisa"].refresh()

    def refresh_history(self):
        self.views["historico"].refresh()

    def set_status(self, text):
        self.lbl_status.config(text=text)


def main():
    RepisApp().mainloop()


if __name__ == "__main__":
    main()
