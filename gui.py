import threading
import traceback
import json
import sys
import os
import subprocess
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# Importa funções do pipeline
try:
    from main import (
        parse_lis_table,
        save_df_to_excel_only,
        calcular_estatisticas_do_df,
        escrever_estatisticas_excel,
        criar_grafico_a_partir_do_excel,
        criar_grafico_comparativo,
    )
except Exception:
    # fallback: erro será exibido quando tentar abrir GUI via main
    raise

PREFS_FILE = Path.home() / ".lis_analysis_gui.json"

def _scan_lis(folder: Path):
    """Retorna arquivos .lis/.LIS ordenados por modificação (desc)."""
    folder = Path(folder)
    files = list(folder.glob('*.lis')) + list(folder.glob('*.LIS'))
    try:
        files.sort(key=lambda f: f.stat().st_mtime, reverse=True)
    except Exception:
        files.sort()
    return files


def _fmt_size(nbytes: int) -> str:
    for unit in ['B','KB','MB','GB','TB']:
        if nbytes < 1024:
            return f"{nbytes:.0f} {unit}"
        nbytes /= 1024
    return f"{nbytes:.0f} PB"


def _open_in_file_manager(path: Path):
    try:
        if sys.platform.startswith('linux'):
            subprocess.Popen(['xdg-open', str(path)])
        elif sys.platform == 'darwin':
            subprocess.Popen(['open', str(path)])
        elif os.name == 'nt':
            os.startfile(str(path))  # type: ignore[attr-defined]
        else:
            messagebox.showinfo('Abrir pasta', f'Abra manualmente: {path}')
    except Exception as e:
        messagebox.showerror('Erro ao abrir', str(e))


class _Tooltip:
    def __init__(self, widget, text: str, delay_ms: int = 500):
        self.widget = widget
        self.text = text
        self.delay = delay_ms
        self._after = None
        self.tip = None
        widget.bind('<Enter>', self._schedule)
        widget.bind('<Leave>', self._hide)

    def _schedule(self, _):
        self._after = self.widget.after(self.delay, self._show)

    def _show(self):
        if self.tip or not self.text:
            return
        x, y, cx, cy = self.widget.bbox('insert') if hasattr(self.widget, 'bbox') else (0,0,0,0)
        x = x + self.widget.winfo_rootx() + 20
        y = y + self.widget.winfo_rooty() + 20
        self.tip = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")
        lbl = ttk.Label(tw, text=self.text, relief='solid', padding=6)
        lbl.pack()

    def _hide(self, _):
        if self._after:
            self.widget.after_cancel(self._after)
            self._after = None
        if self.tip:
            self.tip.destroy()
            self.tip = None


class LisAnalysisApp:
    def __init__(self, root: tk.Tk, folder: Path, outdir: Path, start_index: int = 1):
        self.root = root
        self.root.title('LIS Analysis — Interface')
        try:
            self.root.iconbitmap(default='')  # noop se não houver ícone
        except Exception:
            pass

        # estilo/tema
        self.style = ttk.Style()
        # usa 'clam' por padrão para aparência moderna e consistente
        try:
            self.style.theme_use('clam')
        except Exception:
            pass
        # aplica paleta azul clara e branco
        self._apply_theme_colors()

        # estados e variáveis
        self.folder_var = tk.StringVar(value=str(folder))
        self.outdir_var = tk.StringVar(value=str(outdir))
        self.start_idx_var = tk.IntVar(value=start_index)
        self.filter_var = tk.StringVar()
        self.status_var = tk.StringVar(value='Pronto.')
        self.progress_var = tk.IntVar(value=0)
        self.total_var = tk.IntVar(value=0)
        self.cancel_event = threading.Event()
        self._files_cache = []  # lista de Path
        self._sort_desc = False
        self._sort_col = 'nome'

        self._load_prefs()
        self._build_menu()
        self._build_ui()
        self._bind_shortcuts()
        self.refresh_list()

    # preferências
    def _load_prefs(self):
        try:
            if PREFS_FILE.exists():
                data = json.load(PREFS_FILE.open('r', encoding='utf-8'))
                self.folder_var.set(data.get('folder', self.folder_var.get()))
                self.outdir_var.set(data.get('outdir', self.outdir_var.get()))
                theme = data.get('theme')
                if theme and theme in self.style.theme_names():
                    self.style.theme_use(theme)
        except Exception:
            pass

    def _save_prefs(self):
        try:
            data = {
                'folder': self.folder_var.get(),
                'outdir': self.outdir_var.get(),
                'theme': self.style.theme_use(),
            }
            PREFS_FILE.write_text(json.dumps(data, indent=2), encoding='utf-8')
        except Exception:
            pass

    # construção da UI
    def _build_menu(self):
        menubar = tk.Menu(self.root)

        filem = tk.Menu(menubar, tearoff=0)
        filem.add_command(label='Abrir pasta de entrada…', command=self._choose_folder, accelerator='Ctrl+O')
        filem.add_command(label='Escolher pasta de saída…', command=self._choose_outdir, accelerator='Ctrl+S')
        filem.add_separator()
        filem.add_command(label='Abrir pasta de saída', command=self._open_outdir)
        filem.add_separator()
        filem.add_command(label='Sair', command=self.root.quit, accelerator='Ctrl+Q')
        menubar.add_cascade(label='Arquivo', menu=filem)

        viewm = tk.Menu(menubar, tearoff=0)
        theme_menu = tk.Menu(viewm, tearoff=0)
        for th in self.style.theme_names():
            theme_menu.add_radiobutton(label=th, command=lambda t=th: self._set_theme(t), value=th)
        viewm.add_cascade(label='Tema', menu=theme_menu)
        menubar.add_cascade(label='Exibir', menu=viewm)

        helpm = tk.Menu(menubar, tearoff=0)
        helpm.add_command(label='Sobre', command=lambda: messagebox.showinfo('Sobre', 'LIS Analysis GUI\nMelhorias de UI/UX com Tkinter'))
        menubar.add_cascade(label='Ajuda', menu=helpm)

        self.root.config(menu=menubar)

    def _build_ui(self):
        container = ttk.Frame(self.root, padding=10)
        container.pack(fill='both', expand=True)

        # Linha 1: Pastas e índice
        row1 = ttk.LabelFrame(container, text='Configurações', padding=(10,8), style='Card.TLabelframe')
        row1.pack(fill='x')

        ttk.Label(row1, text='Pasta (.lis):').grid(row=0, column=0, sticky='w')
        self.ent_folder = ttk.Entry(row1, textvariable=self.folder_var)
        self.ent_folder.grid(row=0, column=1, sticky='we', padx=6)
        btn_folder = ttk.Button(row1, text='Escolher…', command=self._choose_folder)
        btn_folder.grid(row=0, column=2, sticky='w')
        _Tooltip(btn_folder, 'Selecionar pasta com arquivos .lis')

        ttk.Label(row1, text='Saída:').grid(row=1, column=0, sticky='w', pady=(6,0))
        self.ent_out = ttk.Entry(row1, textvariable=self.outdir_var)
        self.ent_out.grid(row=1, column=1, sticky='we', padx=6, pady=(6,0))
        btn_out = ttk.Button(row1, text='Escolher…', command=self._choose_outdir)
        btn_out.grid(row=1, column=2, sticky='w', pady=(6,0))
        _Tooltip(btn_out, 'Selecionar pasta onde os resultados serão salvos')

        ttk.Label(row1, text='Índice inicial:').grid(row=2, column=0, sticky='w', pady=(6,0))
        spn = ttk.Spinbox(row1, from_=1, to=9999, textvariable=self.start_idx_var, width=8)
        spn.grid(row=2, column=1, sticky='w', pady=(6,0))
        _Tooltip(spn, 'Número para iniciar a contagem dos arquivos de saída')

        row1.columnconfigure(1, weight=1)

        # Linha 2: Filtro
        row2 = ttk.Frame(container, padding=(0,8,0,0))
        row2.pack(fill='x')
        ttk.Label(row2, text='Filtro:').pack(side='left')
        ent_filter = ttk.Entry(row2, textvariable=self.filter_var, width=30)
        ent_filter.pack(side='left', padx=6, fill='x', expand=True)
        _Tooltip(ent_filter, 'Filtra por parte do nome do arquivo (.lis)')
        ttk.Button(row2, text='Aplicar', command=self.refresh_list).pack(side='left')

        # Linha 3: Lista (Treeview)
        row3 = ttk.LabelFrame(container, text='Arquivos .lis encontrados', padding=(6,6), style='Card.TLabelframe')
        row3.pack(fill='both', expand=True, pady=(8,0))

        columns = ('nome', 'tamanho', 'modificado')
        self.tv = ttk.Treeview(row3, columns=columns, show='headings', selectmode='extended')
        self.tv.heading('nome', text='Nome', command=lambda: self._sort_by('nome'))
        self.tv.heading('tamanho', text='Tamanho', command=lambda: self._sort_by('tamanho'))
        self.tv.heading('modificado', text='Modificado', command=lambda: self._sort_by('modificado'))
        self.tv.column('nome', anchor='w', width=420, stretch=True)
        self.tv.column('tamanho', anchor='center', width=120)
        self.tv.column('modificado', anchor='center', width=180)
        vsb = ttk.Scrollbar(row3, orient='vertical', command=self.tv.yview)
        hsb = ttk.Scrollbar(row3, orient='horizontal', command=self.tv.xview)
        self.tv.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tv.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, columnspan=2, sticky='ew')
        row3.rowconfigure(0, weight=1)
        row3.columnconfigure(0, weight=1)

        # Linha 4: Botões de ação
        row4 = ttk.Frame(container, padding=(0,8,0,0))
        row4.pack(fill='x')
        self.btn_refresh = ttk.Button(row4, text='Atualizar (F5)', command=self.refresh_list)
        self.btn_refresh.pack(side='left')
        self.btn_select_all = ttk.Button(row4, text='Selecionar tudo (Ctrl+A)', command=self._select_all)
        self.btn_select_all.pack(side='left', padx=6)
        self.btn_clear = ttk.Button(row4, text='Limpar seleção', command=self._clear_sel)
        self.btn_clear.pack(side='left')
        self.btn_open_out = ttk.Button(row4, text='Abrir saída', command=self._open_outdir)
        self.btn_open_out.pack(side='left', padx=(6,0))
        self.btn_process = ttk.Button(row4, text='Processar selecionados (Ctrl+P)', command=self.process_selected)
        self.btn_process.pack(side='right')
        self.btn_cancel = ttk.Button(row4, text='Cancelar', command=self._cancel, state='disabled')
        self.btn_cancel.pack(side='right', padx=(0,6))

        # Linha 5: Progresso + status
        row5 = ttk.Frame(container)
        row5.pack(fill='x', pady=(6,0))
        self.pb = ttk.Progressbar(row5, variable=self.progress_var, maximum=100, style='Blue.Horizontal.TProgressbar')
        self.pb.pack(fill='x')
        self.status = ttk.Label(container, textvariable=self.status_var, relief='sunken', anchor='w')
        self.status.pack(fill='x', pady=(6,0))
        # cor de fundo da janela
        try:
            self.root.configure(bg='#ffffff')
        except Exception:
            pass

    def _bind_shortcuts(self):
        self.root.bind('<Control-o>', lambda e: self._choose_folder())
        self.root.bind('<Control-s>', lambda e: self._choose_outdir())
        self.root.bind('<Control-q>', lambda e: self.root.quit())
        self.root.bind('<Control-a>', lambda e: self._select_all())
        self.root.bind('<F5>', lambda e: self.refresh_list())
        self.root.bind('<Control-p>', lambda e: self.process_selected())

    # ações
    def _choose_folder(self):
        sel = filedialog.askdirectory(initialdir=self.folder_var.get() or '.')
        if sel:
            self.folder_var.set(sel)
            self._save_prefs()
            self.refresh_list()

    def _choose_outdir(self):
        sel = filedialog.askdirectory(initialdir=self.outdir_var.get() or '.')
        if sel:
            self.outdir_var.set(sel)
            self._save_prefs()

    def _open_outdir(self):
        outp = Path(self.outdir_var.get()).expanduser()
        outp.mkdir(parents=True, exist_ok=True)
        _open_in_file_manager(outp)

    def _select_all(self):
        for iid in self.tv.get_children(''):
            self.tv.selection_add(iid)

    def _clear_sel(self):
        self.tv.selection_remove(self.tv.selection())

    def _set_controls_state(self, state: str):
        widgets = [
            self.btn_refresh, self.btn_select_all, self.btn_clear, self.btn_process,
            self.ent_folder, self.ent_out
        ]
        for w in widgets:
            try:
                w.configure(state=state)
            except Exception:
                pass
        # botão cancelar no inverso
        try:
            self.btn_cancel.configure(state='normal' if state == 'disabled' else 'disabled')
        except Exception:
            pass

    def _filtered_files(self):
        q = (self.filter_var.get() or '').strip().lower()
        if not q:
            return list(self._files_cache)
        return [p for p in self._files_cache if q in p.name.lower()]

    def _sort_by(self, col: str):
        if self._sort_col == col:
            self._sort_desc = not self._sort_desc
        else:
            self._sort_col = col
            self._sort_desc = False
        self._populate_tree()

    def refresh_list(self):
        folder = Path(self.folder_var.get()).expanduser()
        try:
            self._files_cache = _scan_lis(folder)
        except Exception:
            self._files_cache = []
        self._populate_tree()
        self.status_var.set(f"{len(self._files_cache)} arquivo(s) encontrado(s) em {folder}.")

    def _populate_tree(self):
        for iid in self.tv.get_children(''):
            self.tv.delete(iid)
        files = self._filtered_files()

        # ordenação
        def key_func(p: Path):
            try:
                if self._sort_col == 'tamanho':
                    return p.stat().st_size
                if self._sort_col == 'modificado':
                    return p.stat().st_mtime
                return p.name.lower()
            except Exception:
                return 0

        files.sort(key=key_func, reverse=self._sort_desc)

        # configura tags de alternância de linha
        try:
            self.tv.tag_configure('odd', background='#f7fbff')
            self.tv.tag_configure('even', background='#ffffff')
        except Exception:
            pass

        for idx, p in enumerate(files):
            try:
                st = p.stat()
                size = _fmt_size(st.st_size)
                mod = datetime.fromtimestamp(st.st_mtime).strftime('%Y-%m-%d %H:%M:%S')
            except Exception:
                size, mod = '-', '-'
            tag = 'odd' if idx % 2 else 'even'
            self.tv.insert('', 'end', iid=str(p), values=(p.name, size, mod), tags=(tag,))

        self.total_var.set(len(files))

    def _cancel(self):
        if self.cancel_event.is_set():
            return
        self.cancel_event.set()
        self.status_var.set('Cancelando…')

    def process_selected(self):
        sels = self.tv.selection()
        if not sels:
            messagebox.showwarning('Aviso', 'Nenhum arquivo selecionado.')
            return
        paths = [Path(iid) for iid in sels]
        outp = Path(self.outdir_var.get()).expanduser()
        try:
            start = int(self.start_idx_var.get()) if self.start_idx_var.get() else 1
        except Exception:
            start = 1

        def worker():
            try:
                self._set_controls_state('disabled')
                self.status_var.set('Processando…')
                self.cancel_event.clear()
                outp.mkdir(parents=True, exist_ok=True)
                excel_paths = []
                idx = start
                total = len(paths)
                for i, lp in enumerate(paths, start=1):
                    if self.cancel_event.is_set():
                        self.status_var.set('Cancelado pelo usuário.')
                        break
                    self.status_var.set(f'Processando: {lp.name} ({i}/{total})')
                    # parse
                    try:
                        df, stats_lines, summary_from_lis = parse_lis_table(lp)
                    except Exception:
                        messagebox.showerror('Erro ao ler', traceback.format_exc())
                        continue
                    if df is None:
                        continue
                    excel_path = outp / f"Resultados_Simulacao_{idx}.xlsx"
                    try:
                        save_df_to_excel_only(df, excel_path)
                    except Exception:
                        messagebox.showerror('Erro ao salvar Excel', traceback.format_exc())
                        continue
                    # stats
                    try:
                        computed_stats = calcular_estatisticas_do_df(df)
                    except Exception:
                        computed_stats = {}
                    try:
                        escrever_estatisticas_excel(excel_path, computed_stats, summary_from_lis=summary_from_lis)
                    except Exception:
                        pass
                    # plots
                    try:
                        criar_grafico_a_partir_do_excel(excel_path, outp, sim_index=idx, salvar_png=True, mostrar=False)
                    except Exception:
                        pass
                    excel_paths.append(excel_path)
                    idx += 1
                    # progresso
                    pct = int(i * 100 / max(1, total))
                    self.progress_var.set(pct)
                if not self.cancel_event.is_set() and len(excel_paths) > 1:
                    self.status_var.set('Gerando gráfico comparativo…')
                    try:
                        criar_grafico_comparativo(excel_paths, outp, mostrar=False)
                    except Exception:
                        pass
                if self.cancel_event.is_set():
                    messagebox.showinfo('Cancelado', 'O processamento foi cancelado.')
                else:
                    self.status_var.set('Concluído.')
                    messagebox.showinfo('Concluído', f'Processo finalizado. Verifique: {outp}')
            except Exception:
                messagebox.showerror('Erro', traceback.format_exc())
            finally:
                self._set_controls_state('normal')
                self.progress_var.set(0)
                self.cancel_event.clear()
                self._save_prefs()

        threading.Thread(target=worker, daemon=True).start()

    # tema e estilos
    def _set_theme(self, theme: str):
        try:
            self.style.theme_use(theme)
        except Exception:
            return
        self._apply_theme_colors()

    def _apply_theme_colors(self):
        """Aplica uma paleta clara com azul e branco aos widgets ttk."""
        BG = '#ffffff'
        TEXT = '#1a1a1a'
        ACCENT_BG = '#eaf3ff'
        HEADER_BG = '#e6f2ff'
        HEADER_FG = '#0f3d7a'
        SEL_BG = '#cfe8ff'
        BORDER = '#c7defc'
        PROGRESS = '#64b5f6'

        # janela principal
        try:
            self.root.configure(bg=BG)
        except Exception:
            pass

        # frames e LabelFrames
        self.style.configure('TFrame', background=BG)
        self.style.configure('Card.TLabelframe', background=BG, bordercolor=BORDER)
        self.style.configure('Card.TLabelframe.Label', background=BG, foreground=HEADER_FG)
        self.style.configure('TLabelframe', background=BG)
        self.style.configure('TLabelframe.Label', background=BG, foreground=HEADER_FG)

        # labels e entradas
        self.style.configure('TLabel', background=BG, foreground=TEXT)
        self.style.configure('TEntry', fieldbackground=BG, foreground=TEXT)
        try:
            self.style.map('TEntry', fieldbackground=[('disabled', '#f0f0f0')])
        except Exception:
            pass

        # botões
        self.style.configure('TButton', background=ACCENT_BG)
        try:
            self.style.map('TButton', background=[('active', '#d6eaff')])
        except Exception:
            pass

        # treeview
        self.style.configure('Treeview', background=BG, fieldbackground=BG, foreground=TEXT, rowheight=22)
        self.style.map('Treeview', background=[('selected', SEL_BG)], foreground=[('selected', TEXT)])
        self.style.configure('Treeview.Heading', background=HEADER_BG, foreground=HEADER_FG)

        # scrollbars (parcial, nem todos temas suportam)
        try:
            self.style.configure('Vertical.TScrollbar', background=BG)
            self.style.configure('Horizontal.TScrollbar', background=BG)
        except Exception:
            pass

        # progressbar
        self.style.configure('Blue.Horizontal.TProgressbar', troughcolor=ACCENT_BG, background=PROGRESS)


def launch_gui(folder: Path, outdir: Path, start_index: int = 1):
    """Ponto de entrada público mantendo assinatura original."""
    folder = Path(folder)
    outdir = Path(outdir)
    root = tk.Tk()
    app = LisAnalysisApp(root, folder, outdir, start_index)
    root.protocol('WM_DELETE_WINDOW', lambda: (app._save_prefs(), root.destroy()))
    root.mainloop()
