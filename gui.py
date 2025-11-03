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
        parse_lis_time_series,
        save_time_series_to_excel,
        criar_grafico_series_temporais,
    )
    from acp_parser import (
        AcpParser,
        AtpRunner,
        modify_acp_rpi,
        run_acp_simulation
    )
    from control_detector import (
        ControlDetector,
        FileControlInfo,
        analyze_workspace_files
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


def _scan_acp(folder: Path):
    """Retorna arquivos .acp/.ACP ordenados por modificação (desc)."""
    folder = Path(folder)
    files = list(folder.glob('*.acp')) + list(folder.glob('*.ACP'))
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
        self.filetype_var = tk.StringVar(value='.lis')  # '.lis' | '.acp' | 'ambos'
        self.status_var = tk.StringVar(value='Pronto.')
        self.progress_var = tk.IntVar(value=0)
        self.total_var = tk.IntVar(value=0)
        self.cancel_event = threading.Event()
        self._files_cache = []  # lista de Path
        self._sort_desc = False
        self._sort_col = 'nome'
        
        # Checkboxes de opções (8 no total)
        self.show_plots_var = tk.BooleanVar(value=False)
        self.open_output_var = tk.BooleanVar(value=True)
        self.only_comparative_var = tk.BooleanVar(value=False)
        self.save_logs_var = tk.BooleanVar(value=True)
        self.overwrite_var = tk.BooleanVar(value=True)
        self.hide_errors_var = tk.BooleanVar(value=False)
        self.parallel_process_var = tk.BooleanVar(value=False)
        self.auto_organize_var = tk.BooleanVar(value=True)
        
        # Variáveis para seleção de variáveis do .lis
        self.available_variables = []  # Lista de variáveis detectadas
        self.variable_checkboxes = {}  # Dict: {var_name: BooleanVar}
        self.variables_frame = None  # Frame que contém os checkboxes de variáveis
        
        # Variáveis para controle inteligente de parâmetros
        self.detected_controls = []  # Lista de FileControlInfo
        self.control_widgets = {}  # Dict: {param_name: widget}
        self.control_frame = None  # Frame para controles dinâmicos

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
                # Carregar preferências dos checkboxes
                self.show_plots_var.set(data.get('show_plots', False))
                self.open_output_var.set(data.get('open_output', True))
                self.only_comparative_var.set(data.get('only_comparative', False))
                self.save_logs_var.set(data.get('save_logs', True))
                self.overwrite_var.set(data.get('overwrite', True))
                self.hide_errors_var.set(data.get('hide_errors', False))
                self.parallel_process_var.set(data.get('parallel_process', False))
                self.auto_organize_var.set(data.get('auto_organize', True))
        except Exception:
            pass

    def _save_prefs(self):
        try:
            data = {
                'folder': self.folder_var.get(),
                'outdir': self.outdir_var.get(),
                'theme': self.style.theme_use(),
                'show_plots': self.show_plots_var.get(),
                'open_output': self.open_output_var.get(),
                'only_comparative': self.only_comparative_var.get(),
                'save_logs': self.save_logs_var.get(),
                'overwrite': self.overwrite_var.get(),
                'hide_errors': self.hide_errors_var.get(),
                'parallel_process': self.parallel_process_var.get(),
                'auto_organize': self.auto_organize_var.get(),
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
        # Frame principal com Canvas e Scrollbar
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill='both', expand=True)
        
        # Canvas para conter todo o conteúdo
        canvas = tk.Canvas(main_frame, highlightthickness=0)
        canvas.pack(side='left', fill='both', expand=True)
        
        # Scrollbar vertical
        scrollbar = ttk.Scrollbar(main_frame, orient='vertical', command=canvas.yview)
        scrollbar.pack(side='right', fill='y')
        
        # Configurar canvas com scrollbar
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Frame interno que conterá todos os widgets
        container = ttk.Frame(canvas, padding=10)
        canvas_window = canvas.create_window((0, 0), window=container, anchor='nw')
        
        # Atualizar scroll region quando o tamanho mudar
        def _on_frame_configure(event=None):
            canvas.configure(scrollregion=canvas.bbox('all'))
        
        def _on_canvas_configure(event):
            # Ajustar largura do frame interno para preencher o canvas
            canvas.itemconfig(canvas_window, width=event.width)
        
        container.bind('<Configure>', _on_frame_configure)
        canvas.bind('<Configure>', _on_canvas_configure)
        
        # Suporte para scroll com mouse wheel
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), 'units')
        
        def _bind_mousewheel(event):
            canvas.bind_all('<MouseWheel>', _on_mousewheel)
        
        def _unbind_mousewheel(event):
            canvas.unbind_all('<MouseWheel>')
        
        canvas.bind('<Enter>', _bind_mousewheel)
        canvas.bind('<Leave>', _unbind_mousewheel)

        # Linha 1: Pastas e índice
        row1 = ttk.LabelFrame(container, text='⚙️ Configurações', padding=(10,8), style='Card.TLabelframe')
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

        # Linha 1.5: Opções (Checkboxes - EXPANDIDO)
        row1_5 = ttk.LabelFrame(container, text='⚙️ Opções de Processamento', padding=(10,8), style='Card.TLabelframe')
        row1_5.pack(fill='x', pady=(8,0))

        # Frame para organizar checkboxes em 4 colunas
        chk_col1 = ttk.Frame(row1_5)
        chk_col1.pack(side='left', fill='both', expand=True, padx=2)
        
        chk_col2 = ttk.Frame(row1_5)
        chk_col2.pack(side='left', fill='both', expand=True, padx=2)
        
        chk_col3 = ttk.Frame(row1_5)
        chk_col3.pack(side='left', fill='both', expand=True, padx=2)
        
        chk_col4 = ttk.Frame(row1_5)
        chk_col4.pack(side='left', fill='both', expand=True, padx=2)

        # Coluna 1
        chk1 = ttk.Checkbutton(chk_col1, text='📊 Mostrar gráficos', variable=self.show_plots_var)
        chk1.pack(anchor='w', pady=2)
        _Tooltip(chk1, 'Abre gráficos automaticamente')

        chk5 = ttk.Checkbutton(chk_col1, text='🔇 Ocultar erros', variable=self.hide_errors_var)
        chk5.pack(anchor='w', pady=2)
        _Tooltip(chk5, 'Não exibe caixas de erro')

        # Coluna 2
        chk2 = ttk.Checkbutton(chk_col2, text='📂 Abrir pasta', variable=self.open_output_var)
        chk2.pack(anchor='w', pady=2)
        _Tooltip(chk2, 'Abre pasta ao concluir')

        chk6 = ttk.Checkbutton(chk_col2, text='⚙️ Processar paralelo', variable=self.parallel_process_var)
        chk6.pack(anchor='w', pady=2)
        _Tooltip(chk6, 'Processa múltiplos em paralelo')

        # Coluna 3
        chk3 = ttk.Checkbutton(chk_col3, text='⚡ Só comparativo', variable=self.only_comparative_var)
        chk3.pack(anchor='w', pady=2)
        _Tooltip(chk3, 'Apenas gráfico comparativo (~50% rápido)')

        chk7 = ttk.Checkbutton(chk_col3, text='📁 Auto-organizar', variable=self.auto_organize_var)
        chk7.pack(anchor='w', pady=2)
        _Tooltip(chk7, 'Organiza resultados em pastas')

        # Coluna 4
        chk4 = ttk.Checkbutton(chk_col4, text='📝 Salvar logs', variable=self.save_logs_var)
        chk4.pack(anchor='w', pady=2)
        _Tooltip(chk4, 'Cria log de processamento')

        chk8 = ttk.Checkbutton(chk_col4, text='♻️ Sobrescrever', variable=self.overwrite_var)
        chk8.pack(anchor='w', pady=2)
        _Tooltip(chk8, 'Substitui arquivos existentes')

        # Linha 1.7: Seleção de Variáveis do .lis (NOVA SEÇÃO DINÂMICA)
        row1_7 = ttk.LabelFrame(container, text='📊 Variáveis do Arquivo .lis', padding=(10,8), style='Card.TLabelframe')
        row1_7.pack(fill='x', pady=(8,0))
        
        # Frame interno para os checkboxes de variáveis
        self.variables_frame = ttk.Frame(row1_7)
        self.variables_frame.pack(fill='both', expand=True)
        
        # Mensagem inicial (será substituída quando variáveis forem detectadas)
        self.variables_label = ttk.Label(
            self.variables_frame, 
            text='💡 Selecione um arquivo .lis para detectar variáveis disponíveis',
            foreground='gray'
        )
        self.variables_label.pack(pady=10)
        
        # Botão para detectar variáveis
        btn_detect = ttk.Button(row1_7, text='🔍 Detectar Variáveis', command=self._detect_variables)
        btn_detect.pack(pady=(5,0))
        _Tooltip(btn_detect, 'Analisa o primeiro arquivo selecionado para detectar variáveis')

        # Linha 1.8: Controle Inteligente de Parâmetros (NOVA SEÇÃO DINÂMICA)
        row1_8 = ttk.LabelFrame(container, text='🎯 Controle Inteligente de Parâmetros', padding=(10,8), style='Card.TLabelframe')
        row1_8.pack(fill='x', pady=(8,0))
        
        # Frame interno para controles dinâmicos
        self.control_frame = ttk.Frame(row1_8)
        self.control_frame.pack(fill='both', expand=True)
        
        # Mensagem inicial
        self.control_label = ttk.Label(
            self.control_frame,
            text='💡 Selecione arquivos .lis/.acp para detectar parâmetros (RPI, RF, etc)',
            foreground='gray'
        )
        self.control_label.pack(pady=10)
        
        # Botões de ação
        control_buttons = ttk.Frame(row1_8)
        control_buttons.pack(fill='x', pady=(5,0))
        
        self.btn_detect_controls = ttk.Button(control_buttons, text='🔍 Detectar Parâmetros', command=self._detect_control_parameters)
        self.btn_detect_controls.pack(side='left', padx=2)
        _Tooltip(self.btn_detect_controls, 'Analisa arquivos selecionados e identifica RPI, RF e outros parâmetros')
        
        self.btn_show_summary = ttk.Button(control_buttons, text='📊 Resumo', command=self._show_control_summary)
        self.btn_show_summary.pack(side='left', padx=2)
        _Tooltip(self.btn_show_summary, 'Mostra resumo dos parâmetros detectados')
        
        self.btn_run_atp = ttk.Button(control_buttons, text='🚀 Rodar ATP', command=self._run_atp_simulation)
        self.btn_run_atp.pack(side='left', padx=2)
        _Tooltip(self.btn_run_atp, 'Executa simulação ATP do arquivo .acp selecionado e salva o .lis resultante')
        
        # Linha para executável ATP (fixo)
        atp_exe_frame = ttk.Frame(row1_8)
        atp_exe_frame.pack(fill='x', pady=(8,0))
        
        ttk.Label(atp_exe_frame, text='Executável ATP:').pack(side='left')
        self.atp_exe_var = tk.StringVar()
        self.ent_atp_exe = ttk.Entry(atp_exe_frame, textvariable=self.atp_exe_var, width=35)
        self.ent_atp_exe.pack(side='left', padx=6, fill='x', expand=True)
        btn_atp_exe = ttk.Button(atp_exe_frame, text='Escolher…', command=self._choose_atp_executable)
        btn_atp_exe.pack(side='left')
        _Tooltip(btn_atp_exe, 'Caminho para tpbig ou atpmingw (opcional)')

        # Linha 2: Filtro
        row2 = ttk.Frame(container, padding=(0,4,0,0))
        row2.pack(fill='x')
        ttk.Label(row2, text='🔍 Filtro:').pack(side='left')
        ent_filter = ttk.Entry(row2, textvariable=self.filter_var, width=30)
        ent_filter.pack(side='left', padx=6, fill='x', expand=True)
        _Tooltip(ent_filter, 'Filtra por parte do nome do arquivo')
        ttk.Button(row2, text='Aplicar', command=self.refresh_list).pack(side='left', padx=(0,6))
        ttk.Label(row2, text='Tipo:').pack(side='left')
        cmb_type = ttk.Combobox(row2, textvariable=self.filetype_var, width=8, state='readonly', values=('.lis', '.acp', 'ambos'))
        cmb_type.pack(side='left', padx=6)
        cmb_type.bind('<<ComboboxSelected>>', lambda e: self.refresh_list())
        _Tooltip(cmb_type, 'Escolha o tipo de arquivo a listar (.lis, .acp ou ambos)')

        # Linha 4: Botões de ação (MOVIDO PARA CIMA)
        row4 = ttk.Frame(container, padding=(0,8,0,0))
        row4.pack(fill='x')
        self.btn_refresh = ttk.Button(row4, text='🔄 Atualizar (F5)', command=self.refresh_list)
        self.btn_refresh.pack(side='left', padx=1)
        self.btn_select_all = ttk.Button(row4, text='✓ Selecionar tudo (Ctrl+A)', command=self._select_all)
        self.btn_select_all.pack(side='left', padx=1)
        self.btn_clear = ttk.Button(row4, text='✗ Limpar seleção', command=self._clear_sel)
        self.btn_clear.pack(side='left', padx=1)
        self.btn_clean = ttk.Button(row4, text='🗑️ Limpar Resultados', command=self._clean_results)
        self.btn_clean.pack(side='left', padx=1)
        _Tooltip(self.btn_clean, 'Remove todos os arquivos da pasta de saída')
        self.btn_open_out = ttk.Button(row4, text='📁 Abrir saída', command=self._open_outdir)
        self.btn_open_out.pack(side='left', padx=1)
        self.btn_process = ttk.Button(row4, text='▶ Processar (Ctrl+P)', command=self.process_selected)
        self.btn_process.pack(side='right', padx=1)
        self.btn_cancel = ttk.Button(row4, text='⊗ Cancelar', command=self._cancel, state='disabled')
        self.btn_cancel.pack(side='right', padx=1)

        # Linha 3: Lista (Treeview) - REDUZIDA
        row3 = ttk.LabelFrame(container, text='📋 Arquivos encontrados', padding=(6,6), style='Card.TLabelframe')
        row3.pack(fill='both', expand=True, pady=(8,0)) # expand=True para preencher o espaço

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

        # Linha 5: Progresso + status
        row5 = ttk.Frame(container)
        row5.pack(fill='x', pady=(2,0))
        self.pb = ttk.Progressbar(row5, variable=self.progress_var, maximum=100, style='Blue.Horizontal.TProgressbar')
        self.pb.pack(fill='x')
        self.status = ttk.Label(container, textvariable=self.status_var, relief='sunken', anchor='w')
        self.status.pack(fill='x', pady=(2,0))
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
            self.ent_folder, self.ent_out, self.btn_clean, self.btn_open_out
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

    def _clean_results(self):
        """Remove todos os arquivos da pasta de saída."""
        outp = Path(self.outdir_var.get()).expanduser()
        if not outp.exists():
            messagebox.showwarning('Aviso', 'Pasta de saída não existe.')
            return
        
        files_to_delete = list(outp.glob('*.xlsx')) + list(outp.glob('*.png')) + list(outp.glob('*.txt'))
        if not files_to_delete:
            messagebox.showinfo('Aviso', 'Nenhum arquivo para limpar.')
            return
        
        result = messagebox.askyesno('Confirmação', 
            f'Remover {len(files_to_delete)} arquivo(s) da pasta de saída?\n\nIsso é irreversível!')
        if result:
            deleted = 0
            for f in files_to_delete:
                try:
                    f.unlink()
                    deleted += 1
                except Exception:
                    pass
            messagebox.showinfo('Concluído', f'{deleted} arquivo(s) removido(s).')
            self.status_var.set(f'{deleted} arquivo(s) removido(s) de {outp}.')

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
            ftype = (self.filetype_var.get() or '.lis').strip().lower()
            if ftype == '.acp':
                self._files_cache = _scan_acp(folder)
            elif ftype == 'ambos':
                self._files_cache = _scan_lis(folder) + _scan_acp(folder)
                try:
                    self._files_cache.sort(key=lambda f: f.stat().st_mtime, reverse=True)
                except Exception:
                    self._files_cache.sort()
            else:
                self._files_cache = _scan_lis(folder)
        except Exception:
            self._files_cache = []
        self._populate_tree()
        self.status_var.set(f"{len(self._files_cache)} arquivo(s) encontrado(s) em {folder} (tipo: {self.filetype_var.get()}).")

    def _detect_variables(self):
        """Detecta variáveis do primeiro arquivo .lis selecionado e cria checkboxes."""
        from main import parse_lis_output_variables
        
        # Pegar arquivo selecionado ou o primeiro da lista
        sels = self.tv.selection()
        if sels:
            lis_path = Path(sels[0])
        elif self._files_cache:
            lis_path = self._files_cache[0]
        else:
            messagebox.showwarning('Aviso', 'Nenhum arquivo .lis encontrado.\n\nSelecione uma pasta com arquivos .lis primeiro.')
            return
        
        try:
            # Detectar variáveis
            self.status_var.set(f'Detectando variáveis de {lis_path.name}...')
            self.root.update_idletasks()
            
            variables = parse_lis_output_variables(lis_path)
            
            if not variables:
                messagebox.showinfo('Info', 'Nenhuma variável detectada no arquivo.\n\nVerifique se o arquivo .lis contém a seção "Column headings".')
                self.status_var.set('Pronto.')
                return
            
            # Limpar checkboxes anteriores
            for widget in self.variables_frame.winfo_children():
                widget.destroy()
            
            self.available_variables = variables
            self.variable_checkboxes.clear()
            
            # Criar título
            title_label = ttk.Label(
                self.variables_frame,
                text=f'✅ {len(variables)} variável(is) detectada(s) — Selecione as que deseja analisar:',
                font=('TkDefaultFont', 9, 'bold'),
                foreground='#2F75B5'
            )
            title_label.pack(anchor='w', pady=(0, 5))
            
            # Frame para os checkboxes (layout em colunas)
            chk_container = ttk.Frame(self.variables_frame)
            chk_container.pack(fill='both', expand=True)
            
            # Determinar número de colunas (4 ou menos se houver poucas variáveis)
            num_cols = min(4, max(2, len(variables)))
            cols_frames = []
            for i in range(num_cols):
                col_frame = ttk.Frame(chk_container)
                col_frame.pack(side='left', fill='both', expand=True, padx=2)
                cols_frames.append(col_frame)
            
            # Distribuir checkboxes entre as colunas
            for idx, var in enumerate(variables):
                col_idx = idx % num_cols
                var_checkbox = tk.BooleanVar(value=True)  # Todas selecionadas por padrão
                self.variable_checkboxes[var] = var_checkbox
                
                chk = ttk.Checkbutton(
                    cols_frames[col_idx],
                    text=f'📌 {var}',
                    variable=var_checkbox
                )
                chk.pack(anchor='w', pady=1)
                _Tooltip(chk, f'Incluir variável "{var}" na análise')
            
            # Botões de controle
            btn_frame = ttk.Frame(self.variables_frame)
            btn_frame.pack(fill='x', pady=(8, 0))
            
            ttk.Button(
                btn_frame,
                text='✓ Selecionar Todas',
                command=lambda: self._toggle_all_variables(True)
            ).pack(side='left', padx=2)
            
            ttk.Button(
                btn_frame,
                text='✗ Desmarcar Todas',
                command=lambda: self._toggle_all_variables(False)
            ).pack(side='left', padx=2)
            
            self.status_var.set(f'✅ {len(variables)} variável(is) detectada(s) em {lis_path.name}')
            messagebox.showinfo(
                'Variáveis Detectadas',
                f'✅ {len(variables)} variável(is) encontrada(s):\n\n' + 
                ', '.join(variables) +
                '\n\nTodas foram selecionadas por padrão.\nDesmarque as que não deseja analisar.'
            )
            
        except Exception as e:
            messagebox.showerror('Erro', f'Falha ao detectar variáveis:\n\n{str(e)}')
            self.status_var.set('Erro ao detectar variáveis.')
            import traceback
            traceback.print_exc()

    def _toggle_all_variables(self, state: bool):
        """Marca ou desmarca todas as variáveis."""
        for var_bool in self.variable_checkboxes.values():
            var_bool.set(state)
        action = 'selecionadas' if state else 'desmarcadas'
        self.status_var.set(f'{len(self.variable_checkboxes)} variável(is) {action}.')

    # ==================== MÉTODOS DE CONTROLE INTELIGENTE ====================
    
    def _detect_control_parameters(self):
        """Detecta parâmetros de controle (RPI, RF, etc) nos arquivos selecionados"""
        # Pegar arquivos selecionados
        sels = self.tv.selection()
        if not sels:
            messagebox.showwarning('Aviso', 'Selecione ao menos um arquivo .lis!')
            return
        
        selected_files = [Path(item) for item in sels]
        
        try:
            self.status_var.set('Detectando parâmetros de controle...')
            self.root.update_idletasks()
            
            # Detectar parâmetros
            self.detected_controls = ControlDetector.detect_from_files(selected_files)
            
            # Filtrar apenas arquivos COM controle
            with_control = [info for info in self.detected_controls if info.has_control]
            without_control = [info for info in self.detected_controls if not info.has_control]
            
            if not with_control:
                msg = f'❌ Nenhum parâmetro de controle detectado!\n\n'
                msg += f'Arquivos analisados: {len(selected_files)}\n'
                if without_control:
                    msg += f'\n📄 Arquivos "Sem Controle":\n'
                    for info in without_control:
                        msg += f'  • {info.original_path.name}\n'
                messagebox.showinfo('Info', msg)
                return
            
            # Limpar frame anterior
            for widget in self.control_frame.winfo_children():
                widget.destroy()
            
            self.control_widgets.clear()
            
            # Criar título
            title_text = f'✅ {len(with_control)} arquivo(s) com controle detectado(s)'
            if without_control:
                title_text += f' | {len(without_control)} sem controle'
            
            title_label = ttk.Label(
                self.control_frame,
                text=title_text,
                font=('TkDefaultFont', 9, 'bold'),
                foreground='#2F75B5'
            )
            title_label.pack(anchor='w', pady=(0, 5))
            
            # Coletar todos os parâmetros únicos detectados
            all_params = {}  # {param_name: [values]}
            for info in with_control:
                for param in info.parameters:
                    if param.name not in all_params:
                        all_params[param.name] = set()
                    all_params[param.name].add(param.value)
            
            # Criar controles dinâmicos
            controls_container = ttk.Frame(self.control_frame)
            controls_container.pack(fill='both', expand=True, pady=(5, 0))
            
            row = 0
            for param_name in sorted(all_params.keys()):
                values = sorted(all_params[param_name])
                desc = ControlDetector.get_parameter_description(param_name)
                unit = ControlDetector.UNITS.get(param_name, '')
                
                # Label
                label_text = f'{param_name} ({desc}):'
                ttk.Label(controls_container, text=label_text).grid(
                    row=row, column=0, sticky='w', pady=4
                )
                
                # Valores detectados
                values_text = ', '.join([f'{v:.0f}{unit}' for v in values])
                ttk.Label(
                    controls_container,
                    text=f'📌 Detectado: {values_text}',
                    foreground='#666'
                ).grid(row=row, column=1, sticky='w', padx=(10, 0), pady=4)
                
                row += 1
                
                # Campo de entrada para novo valor
                ttk.Label(controls_container, text=f'   Novo valor:').grid(
                    row=row, column=0, sticky='w', pady=2
                )
                
                new_value_var = tk.DoubleVar(value=values[0] if values else 0)
                self.control_widgets[param_name] = new_value_var
                
                # Spinbox ou Combobox com sugestões
                suggestions = ControlDetector.suggest_values(param_name, values[0] if values else 0)
                
                entry_frame = ttk.Frame(controls_container)
                entry_frame.grid(row=row, column=1, sticky='w', padx=(10, 0), pady=2)
                
                spinbox = ttk.Spinbox(
                    entry_frame,
                    from_=1, to=10000,
                    textvariable=new_value_var,
                    width=12
                )
                spinbox.pack(side='left')
                
                # Botões de sugestões rápidas
                for sug_val in suggestions[:5]:
                    btn = ttk.Button(
                        entry_frame,
                        text=f'{int(sug_val)}',
                        width=4,
                        command=lambda v=sug_val, var=new_value_var: var.set(v)
                    )
                    btn.pack(side='left', padx=1)
                    _Tooltip(btn, f'Definir para {sug_val}{unit}')
                
                row += 1
            
            controls_container.columnconfigure(1, weight=1)
            
            # Botões de ação
            action_frame = ttk.Frame(self.control_frame)
            action_frame.pack(fill='x', pady=(10, 0))
            
            ttk.Button(
                action_frame,
                text='🔧 Aplicar Modificações',
                command=self._apply_control_modifications
            ).pack(side='left', padx=2)
            
            ttk.Button(
                action_frame,
                text='📊 Ver Resumo',
                command=self._show_control_summary
            ).pack(side='left', padx=2)
            
            self.status_var.set(f'✅ {len(with_control)} arquivo(s) com controle detectado(s)')
            
        except Exception as e:
            messagebox.showerror('Erro', f'Falha ao detectar parâmetros:\n\n{str(e)}')
            self.status_var.set('Erro ao detectar parâmetros')
            traceback.print_exc()
    
    def _show_control_summary(self):
        """Mostra resumo detalhado dos parâmetros detectados"""
        if not self.detected_controls:
            messagebox.showinfo('Info', 'Nenhum parâmetro detectado ainda.\n\nClique em "Detectar Parâmetros" primeiro.')
            return
        
        with_control = [info for info in self.detected_controls if info.has_control]
        without_control = [info for info in self.detected_controls if not info.has_control]
        
        msg = "📊 RESUMO DOS PARÂMETROS DETECTADOS\n"
        msg += "=" * 60 + "\n\n"
        
        if with_control:
            msg += f"✅ Arquivos COM controle: {len(with_control)}\n\n"
            
            for info in with_control:
                msg += f"📄 {info.original_path.name}\n"
                msg += f"   Tipo: {info.file_type}\n"
                
                if info.parameters:
                    msg += f"   Parâmetros:\n"
                    for param in info.parameters:
                        desc = ControlDetector.get_parameter_description(param.name)
                        msg += f"      • {param.name} ({desc}): {param.value} {param.unit}\n"
                msg += "\n"
        
        if without_control:
            msg += f"❌ Arquivos SEM controle: {len(without_control)}\n"
            for info in without_control:
                msg += f"   • {info.original_path.name}\n"
        
        messagebox.showinfo('Resumo dos Parâmetros', msg)
    
    def _apply_control_modifications(self):
        """Aplica modificações nos arquivos .acp baseado nos parâmetros editados"""
        if not self.detected_controls:
            messagebox.showwarning('Aviso', 'Nenhum parâmetro detectado!')
            return
        
        if not self.control_widgets:
            messagebox.showwarning('Aviso', 'Nenhuma modificação definida!')
            return
        
        # Coletar novos valores
        new_params = {}
        for param_name, var in self.control_widgets.items():
            new_params[param_name] = var.get()
        
        # Confirmar ação
        params_str = ', '.join([f'{k}={v:.0f}' for k, v in new_params.items()])
        
        confirm = messagebox.askyesno(
            'Confirmar Modificações',
            f'Aplicar as seguintes modificações?\n\n'
            f'{params_str}\n\n'
            f'Serão criados novos arquivos .acp modificados.'
        )
        
        if not confirm:
            return
        
        self.status_var.set('Aplicando modificações...')
        
        try:
            modified_count = 0
            
            with_control = [info for info in self.detected_controls if info.has_control]
            
            for info in with_control:
                # Tentar encontrar arquivo .acp correspondente
                acp_path = info.original_path.with_suffix('.acp')
                
                if not acp_path.exists():
                    # Tentar variações de maiúsculas
                    acp_path = info.original_path.with_suffix('.ACP')
                
                if not acp_path.exists():
                    print(f"⚠️ Arquivo .acp não encontrado para {info.original_path.name}")
                    continue
                
                # Gerar novo nome de arquivo
                new_filename = ControlDetector.generate_new_filename(info, new_params)
                output_path = acp_path.parent / new_filename
                
                # Modificar arquivo .acp
                parser = AcpParser(acp_path)
                parser.extract_atp_from_acp()
                
                modified = False
                for param_name, new_value in new_params.items():
                    if param_name == 'RPI':
                        if parser.modify_rpi_value(new_value):
                            modified = True
                
                if modified:
                    if parser.save_modified_acp(output_path):
                        modified_count += 1
                        print(f"✅ Modificado: {output_path.name}")
            
            if modified_count > 0:
                messagebox.showinfo(
                    'Sucesso',
                    f'✅ {modified_count} arquivo(s) .acp modificado(s)!\n\n'
                    f'Novos arquivos criados com parâmetros atualizados.'
                )
                self.status_var.set(f'✅ {modified_count} arquivo(s) modificado(s)')
            else:
                messagebox.showwarning(
                    'Aviso',
                    'Nenhum arquivo foi modificado.\n\n'
                    'Verifique se os arquivos .acp estão na mesma pasta dos .lis'
                )
                self.status_var.set('Nenhum arquivo modificado')
        
        except Exception as e:
            messagebox.showerror('Erro', f'Falha ao aplicar modificações:\n\n{str(e)}')
            self.status_var.set('Erro ao aplicar modificações')
            traceback.print_exc()

    # ==================== MÉTODOS DE CONTROLE ATP (MANTIDOS PARA COMPATIBILIDADE) ====================
    
    def _choose_atp_executable(self):
        """Escolhe executável do ATP (tpbig, atpmingw)"""
        filepath = filedialog.askopenfilename(
            title='Selecionar executável ATP',
            initialdir='/usr/local/bin',
            filetypes=[
                ('Executáveis', 'tpbig;atpmingw;*.exe'),
                ('Todos os arquivos', '*.*')
            ]
        )
        
        if filepath:
            self.atp_exe_var.set(filepath)
            self.status_var.set(f'Executável ATP: {Path(filepath).name}')
    
    def _choose_acp_file(self):
        """Escolhe arquivo .acp para modificar/simular"""
        folder = Path(self.folder_var.get())
        initial_dir = folder if folder.is_dir() else Path.home()
        
        filepath = filedialog.askopenfilename(
            title='Selecionar arquivo .acp',
            initialdir=initial_dir,
            filetypes=[
                ('Arquivos ATPDraw', '*.acp'),
                ('Todos os arquivos', '*.*')
            ]
        )
        
        if filepath:
            self.acp_file_var.set(filepath)
            self.status_var.set(f'Arquivo .acp selecionado: {Path(filepath).name}')
    
    def _choose_atp_executable(self):
        """Escolhe executável do ATP (tpbig, atpmingw)"""
        filepath = filedialog.askopenfilename(
            title='Selecionar executável ATP',
            initialdir='/usr/local/bin',
            filetypes=[
                ('Executáveis', 'tpbig;atpmingw;*.exe'),
                ('Todos os arquivos', '*.*')
            ]
        )
        
        if filepath:
            self.atp_exe_var.set(filepath)
            self.status_var.set(f'Executável ATP: {Path(filepath).name}')
    
    def _analyze_acp(self):
        """Analisa arquivo .acp e mostra resumo"""
        acp_path = self.acp_file_var.get()
        
        if not acp_path:
            messagebox.showwarning('Aviso', 'Selecione um arquivo .acp primeiro!')
            return
        
        acp_path = Path(acp_path)
        
        if not acp_path.exists():
            messagebox.showerror('Erro', f'Arquivo não encontrado:\n{acp_path}')
            return
        
        self.status_var.set('Analisando arquivo .acp...')
        
        try:
            parser = AcpParser(acp_path)
            parser.extract_atp_from_acp()
            
            if not parser.atp_text:
                messagebox.showerror('Erro', 'Não foi possível extrair conteúdo ATP do arquivo .acp')
                return
            
            params = parser.find_control_parameters()
            
            # Montar mensagem de resumo
            msg = f"📋 Resumo do Arquivo: {acp_path.name}\n"
            msg += "=" * 60 + "\n\n"
            
            if params['dt'] and params['tmax']:
                msg += f"⚙️ Configuração de Tempo:\n"
                msg += f"   • dT   = {params['dt']:.6E} s\n"
                msg += f"   • Tmax = {params['tmax']:.6f} s\n\n"
            
            if params['rpi_values']:
                msg += f"🔌 Resistências de Pré-Inserção (RPI): {len(params['rpi_values'])}\n"
                for rpi in params['rpi_values'][:10]:  # Mostrar até 10
                    msg += f"   • Linha {rpi['line']+1}: {rpi['value']:.2f} Ω\n"
                if len(params['rpi_values']) > 10:
                    msg += f"   ... e mais {len(params['rpi_values']) - 10}\n"
                msg += "\n"
            else:
                msg += "⚠️ Nenhum RPI detectado no arquivo\n\n"
            
            if params['switch_times']:
                msg += f"🔀 Tempos de Chaveamento: {len(params['switch_times'])}\n"
                for sw in params['switch_times'][:5]:
                    msg += f"   • Linha {sw['line']+1}: {sw['time']:.6f} s\n"
                if len(params['switch_times']) > 5:
                    msg += f"   ... e mais {len(params['switch_times']) - 5}\n"
            
            messagebox.showinfo('Análise do Arquivo .acp', msg)
            self.status_var.set(f'Análise concluída: {acp_path.name}')
            
        except Exception as e:
            messagebox.showerror('Erro', f'Falha ao analisar .acp:\n\n{str(e)}')
            self.status_var.set('Erro ao analisar .acp')
            import traceback
            traceback.print_exc()
    
    def _modify_acp_rpi(self):
        """Modifica valor de RPI no arquivo .acp"""
        acp_path = self.acp_file_var.get()
        
        if not acp_path:
            messagebox.showwarning('Aviso', 'Selecione um arquivo .acp primeiro!')
            return
        
        acp_path = Path(acp_path)
        
        if not acp_path.exists():
            messagebox.showerror('Erro', f'Arquivo não encontrado:\n{acp_path}')
            return
        
        new_rpi = self.rpi_value_var.get()
        
        if new_rpi <= 0:
            messagebox.showerror('Erro', 'Valor de RPI deve ser maior que zero!')
            return
        
        # Confirmar ação
        confirm = messagebox.askyesno(
            'Confirmar Modificação',
            f'Modificar RPI para {new_rpi:.2f} Ω?\n\n'
            f'Arquivo original: {acp_path.name}\n'
            f'Novo arquivo: {acp_path.stem}_RPI{int(new_rpi)}.acp'
        )
        
        if not confirm:
            return
        
        self.status_var.set(f'Modificando RPI para {new_rpi:.2f} Ω...')
        
        try:
            output_path = modify_acp_rpi(acp_path, new_rpi)
            
            if output_path:
                messagebox.showinfo(
                    'Sucesso',
                    f'✅ Arquivo modificado criado:\n\n{output_path.name}\n\n'
                    f'RPI = {new_rpi:.2f} Ω'
                )
                
                # Atualizar campo para novo arquivo
                self.acp_file_var.set(str(output_path))
                self.status_var.set(f'RPI modificado: {output_path.name}')
            else:
                messagebox.showerror('Erro', 'Falha ao modificar arquivo .acp')
                self.status_var.set('Erro ao modificar .acp')
        
        except Exception as e:
            messagebox.showerror('Erro', f'Falha ao modificar RPI:\n\n{str(e)}')
            self.status_var.set('Erro ao modificar RPI')
            import traceback
            traceback.print_exc()
    
    def _run_atp_simulation(self):
        """Executa simulação ATP a partir de arquivos .acp selecionados e gera arquivo .lis"""
        # Pegar arquivos .acp selecionados
        sels = self.tv.selection()
        if not sels:
            messagebox.showwarning('Aviso', 'Selecione ao menos um arquivo .acp na lista!')
            return
        
        # Filtrar apenas arquivos .acp
        acp_files = [Path(item) for item in sels if Path(item).suffix.lower() == '.acp']
        
        if not acp_files:
            messagebox.showwarning('Aviso', 'Nenhum arquivo .acp selecionado!\n\nSelecione arquivos .acp para executar a simulação.')
            return
        
        output_dir = Path(self.outdir_var.get())
        atp_exe = self.atp_exe_var.get() or None
        
        # Confirmar ação
        confirm = messagebox.askyesno(
            'Confirmar Simulação ATP',
            f'Executar simulação ATP para {len(acp_files)} arquivo(s)?\n\n' +
            '\n'.join([f'  • {f.name}' for f in acp_files[:5]]) +
            (f'\n  ... e mais {len(acp_files)-5}' if len(acp_files) > 5 else '') +
            f'\n\nResultados serão salvos em:\n{output_dir}'
        )
        
        if not confirm:
            return
        
        self.status_var.set(f'Executando simulação ATP para {len(acp_files)} arquivo(s)...')
        self._set_controls_state('disabled')
        
        def run_thread():
            try:
                runner = AtpRunner(atp_exe)
                
                if not runner.atpdraw_path:
                    error_msg = (
                        '❌ Executável do ATP não encontrado!\n\n'
                        '💡 Soluções:\n\n'
                        '1. Instale o ATP nativo para Linux (tpbig)\n'
                        '   sudo apt install atp\n\n'
                        '2. Use Wine + ATPDraw:\n'
                        '   sudo apt install wine wine64\n'
                        '   chmod +x /home/pedro/ATPDraw/Atpdraw.exe\n\n'
                        '3. Configure manualmente:\n'
                        '   - Campo "Executável ATP" → Escolher executável\n'
                        '   - Para Wine: wine /caminho/para/Atpdraw.exe\n\n'
                        '📖 Veja INSTRUCOES_ATP.md para mais detalhes'
                    )
                    self.root.after(0, lambda: messagebox.showerror('Erro - ATP Não Configurado', error_msg))
                    self.status_var.set('Erro: ATP não encontrado')
                    self._set_controls_state('normal')
                    return
                
                success_count = 0
                failed_files = []
                
                for idx, acp_path in enumerate(acp_files, 1):
                    self.status_var.set(f'[{idx}/{len(acp_files)}] Simulando {acp_path.name}...')
                    self.root.update_idletasks()
                    
                    try:
                        lis_path = runner.run_simulation(acp_path, output_dir)
                        
                        if lis_path:
                            success_count += 1
                            print(f"✅ [{idx}/{len(acp_files)}] Sucesso: {lis_path.name}")
                        else:
                            failed_files.append(acp_path.name)
                            print(f"❌ [{idx}/{len(acp_files)}] Falha: {acp_path.name}")
                    
                    except Exception as e:
                        failed_files.append(acp_path.name)
                        print(f"❌ [{idx}/{len(acp_files)}] Erro em {acp_path.name}: {e}")
                
                # Atualizar lista de arquivos
                self.root.after(0, self.refresh_list)
                
                # Mostrar resultado
                if success_count == len(acp_files):
                    self.root.after(0, lambda: messagebox.showinfo(
                        'Simulação Concluída',
                        f'✅ Todas as {success_count} simulação(ões) concluída(s) com sucesso!\n\n'
                        f'Arquivos .lis salvos em:\n{output_dir}'
                    ))
                    self.status_var.set(f'✅ {success_count} simulação(ões) concluída(s)')
                elif success_count > 0:
                    self.root.after(0, lambda: messagebox.showwarning(
                        'Simulação Parcialmente Concluída',
                        f'✅ {success_count} de {len(acp_files)} simulação(ões) concluída(s)\n\n'
                        f'❌ Falhas em:\n' + '\n'.join([f'  • {f}' for f in failed_files[:10]]) +
                        (f'\n  ... e mais {len(failed_files)-10}' if len(failed_files) > 10 else '')
                    ))
                    self.status_var.set(f'⚠️ {success_count}/{len(acp_files)} simulações concluídas')
                else:
                    self.root.after(0, lambda: messagebox.showerror(
                        'Erro',
                        'Todas as simulações falharam!\n\n'
                        'Verifique a saída do console para mais detalhes.'
                    ))
                    self.status_var.set('❌ Todas as simulações falharam')
            
            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror(
                    'Erro',
                    f'Falha na simulação:\n\n{str(e)}'
                ))
                self.status_var.set('Erro na simulação')
                import traceback
                traceback.print_exc()
            
            finally:
                self._set_controls_state('normal')
        
        threading.Thread(target=run_thread, daemon=True).start()
    
    def _run_full_cycle(self):
        """Executa ciclo completo: Modificar RPI → Simular → Analisar"""
        acp_path = self.acp_file_var.get()
        
        if not acp_path:
            messagebox.showwarning('Aviso', 'Selecione um arquivo .acp primeiro!')
            return
        
        acp_path = Path(acp_path)
        
        if not acp_path.exists():
            messagebox.showerror('Erro', f'Arquivo não encontrado:\n{acp_path}')
            return
        
        new_rpi = self.rpi_value_var.get()
        
        if new_rpi <= 0:
            messagebox.showerror('Erro', 'Valor de RPI deve ser maior que zero!')
            return
        
        # Confirmar ciclo completo
        confirm = messagebox.askyesno(
            'Confirmar Ciclo Completo',
            f'Executar ciclo completo?\n\n'
            f'1. Modificar RPI para {new_rpi:.2f} Ω\n'
            f'2. Executar simulação ATP\n'
            f'3. Analisar resultado .lis\n\n'
            f'Isso pode levar alguns minutos.'
        )
        
        if not confirm:
            return
        
        self.status_var.set('Iniciando ciclo completo...')
        self.btn_full_cycle.config(state='disabled')
        
        def full_cycle_thread():
            try:
                output_dir = Path(self.outdir_var.get())
                atp_exe = self.atp_exe_var.get() or None
                
                # Etapa 1: Modificar RPI
                self.status_var.set(f'[1/3] Modificando RPI para {new_rpi:.2f} Ω...')
                modified_acp = modify_acp_rpi(acp_path, new_rpi)
                
                if not modified_acp:
                    raise Exception('Falha ao modificar arquivo .acp')
                
                print(f"✅ Arquivo modificado: {modified_acp}")
                
                # Etapa 2: Simular
                self.status_var.set('[2/3] Executando simulação ATP...')
                runner = AtpRunner(atp_exe)
                
                if not runner.atpdraw_path:
                    raise Exception('Executável do ATP não encontrado')
                
                lis_path = runner.run_simulation(modified_acp, output_dir)
                
                if not lis_path:
                    raise Exception('Simulação falhou ou .lis não foi gerado')
                
                print(f"✅ Simulação concluída: {lis_path}")
                
                # Etapa 3: Adicionar à lista para análise
                self.status_var.set('[3/3] Atualizando lista...')
                self.root.after(0, self.refresh_list)
                
                # Sucesso!
                self.root.after(0, lambda: messagebox.showinfo(
                    'Ciclo Completo Concluído',
                    f'✅ Ciclo completo executado com sucesso!\n\n'
                    f'1. RPI modificado: {new_rpi:.2f} Ω\n'
                    f'2. Simulação concluída\n'
                    f'3. Arquivo .lis: {lis_path.name}\n\n'
                    f'Agora você pode selecionar o arquivo .lis na lista\n'
                    f'e clicar em "Processar" para gerar gráficos e estatísticas.'
                ))
                
                self.status_var.set(f'Ciclo completo concluído: {lis_path.name}')
            
            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror(
                    'Erro no Ciclo Completo',
                    f'Falha durante o ciclo:\n\n{str(e)}'
                ))
                self.status_var.set('Erro no ciclo completo')
                import traceback
                traceback.print_exc()
            
            finally:
                self.btn_full_cycle.config(state='normal')
        
        threading.Thread(target=full_cycle_thread, daemon=True).start()

    # ==================== FIM DOS MÉTODOS ATP ====================

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
        # Apenas .lis são processados pelo pipeline atual
        lis_paths = [p for p in paths if p.suffix.lower() == '.lis']
        non_lis = [p for p in paths if p.suffix.lower() != '.lis']
        if non_lis:
            messagebox.showinfo('Aviso', f"{len(non_lis)} arquivo(s) não .lis foram ignorados.")
        if not lis_paths:
            messagebox.showwarning('Aviso', 'Nenhum arquivo .lis selecionado para processar.')
            return
        outp = Path(self.outdir_var.get()).expanduser()
        try:
            start = int(self.start_idx_var.get()) if self.start_idx_var.get() else 1
        except Exception:
            start = 1

        # Capturar valores dos checkboxes
        show_plots = self.show_plots_var.get()
        open_output = self.open_output_var.get()
        only_comparative = self.only_comparative_var.get()
        save_logs = self.save_logs_var.get()
        overwrite = self.overwrite_var.get()
        
        # 🆕 CAPTURAR VARIÁVEIS SELECIONADAS
        selected_variables = None
        if self.variable_checkboxes:
            selected_variables = [var for var, boolvar in self.variable_checkboxes.items() if boolvar.get()]
            if not selected_variables:
                messagebox.showwarning('Aviso', 
                    'Nenhuma variável selecionada!\n\n' +
                    'Por favor, selecione pelo menos uma variável para analisar ou\n' +
                    'deixe a seção de variáveis vazia para processar modo tradicional (estatísticas de picos).')
                return
            print(f"📊 Variáveis selecionadas para análise: {', '.join(selected_variables)}")
        
        log_lines = []

        def worker():
            try:
                self._set_controls_state('disabled')
                self.status_var.set('Processando…')
                self.cancel_event.clear()
                outp.mkdir(parents=True, exist_ok=True)
                excel_paths = []
                idx = start
                total = len(lis_paths)
                
                # Log inicial
                if save_logs:
                    log_lines.append(f"=== Processamento iniciado em {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} ===")
                    log_lines.append(f"Pasta de entrada: {self.folder_var.get()}")
                    log_lines.append(f"Pasta de saída: {outp}")
                    log_lines.append(f"Arquivos selecionados: {total}")
                    log_lines.append(f"Índice inicial: {start}")
                    log_lines.append("Opções:")
                    log_lines.append(f"  - Mostrar gráficos: {show_plots}")
                    log_lines.append(f"  - Abrir pasta ao concluir: {open_output}")
                    log_lines.append(f"  - Apenas comparativo: {only_comparative}")
                    log_lines.append(f"  - Salvar logs: {save_logs}")
                    log_lines.append(f"  - Sobrescrever: {overwrite}")
                    if selected_variables:
                        log_lines.append(f"  - Variáveis selecionadas: {', '.join(selected_variables)}")
                    log_lines.append("")
                
                for i, lp in enumerate(lis_paths, start=1):
                    if self.cancel_event.is_set():
                        self.status_var.set('Cancelado pelo usuário.')
                        if save_logs:
                            log_lines.append(f"[CANCELADO] Processamento interrompido em {i}/{total}")
                        break
                    
                    self.status_var.set(f'Processando: {lp.name} ({i}/{total})')
                    
                    if save_logs:
                        log_lines.append(f"[{i}/{total}] {lp.name}")
                    
                    # Verificar se arquivo existe (para lógica de sobrescrita)
                    excel_path = outp / f"Resultados_Simulacao_{idx}.xlsx"
                    if excel_path.exists() and not overwrite:
                        if save_logs:
                            log_lines.append(f"  [PULADO] Arquivo já existe (sobrescrita desativada)")
                        idx += 1
                        continue
                    
                    # 🆕 PROCESSAMENTO BASEADO EM VARIÁVEIS SELECIONADAS
                    if selected_variables:
                        # MODO 1: Análise de séries temporais (novas variáveis)
                        try:
                            df_time_series = parse_lis_time_series(lp, selected_variables)
                            if save_logs:
                                log_lines.append(f"  [OK] Parsing de séries temporais concluído")
                            
                            if df_time_series is not None and not df_time_series.empty:
                                # Salvar séries temporais no Excel
                                save_time_series_to_excel(df_time_series, excel_path, sheet_name='Dados_Temporais')
                                if save_logs:
                                    log_lines.append(f"  [OK] Séries temporais salvas no Excel")
                                
                                # Criar gráfico de séries temporais
                                if not only_comparative:
                                    png_path = outp / f"series_temporais_{idx}.png"
                                    criar_grafico_series_temporais(
                                        df_time_series, 
                                        png_path, 
                                        lis_name=lp.stem,
                                        salvar_png=True, 
                                        mostrar=show_plots
                                    )
                                    if save_logs:
                                        log_lines.append(f"  [OK] Gráfico de séries temporais gerado")
                            else:
                                if save_logs:
                                    log_lines.append(f"  [AVISO] DataFrame de séries temporais vazio")
                        except Exception as e:
                            if save_logs:
                                log_lines.append(f"  [ERRO] Falha no processamento de séries temporais: {str(e)}")
                            import traceback
                            traceback.print_exc()
                    else:
                        # MODO 2: Análise tradicional de estatísticas de picos (modo original)
                        try:
                            df, stats_lines, summary_from_lis = parse_lis_table(lp)
                            if save_logs:
                                log_lines.append(f"  [OK] Parsing tradicional concluído")
                        except Exception as e:
                            if save_logs:
                                log_lines.append(f"  [ERRO] Falha ao fazer parsing: {str(e)}")
                            continue
                        
                        if df is None:
                            if save_logs:
                                log_lines.append(f"  [ERRO] DataFrame vazio")
                            continue
                        
                        # Salvar Excel
                        try:
                            save_df_to_excel_only(df, excel_path)
                            if save_logs:
                                log_lines.append(f"  [OK] Excel salvo")
                        except Exception as e:
                            if save_logs:
                                log_lines.append(f"  [ERRO] Falha ao salvar Excel: {str(e)}")
                            continue
                        
                        # Calcular estatísticas
                        try:
                            computed_stats = calcular_estatisticas_do_df(df)
                            if save_logs:
                                log_lines.append(f"  [OK] Estatísticas calculadas")
                        except Exception as e:
                            computed_stats = {}
                            if save_logs:
                                log_lines.append(f"  [AVISO] Falha ao calcular estatísticas: {str(e)}")
                        
                        # Escrever estatísticas no Excel
                        try:
                            escrever_estatisticas_excel(excel_path, computed_stats, summary_from_lis=summary_from_lis)
                            if save_logs:
                                log_lines.append(f"  [OK] Estatísticas salvas")
                        except Exception as e:
                            if save_logs:
                                log_lines.append(f"  [AVISO] Falha ao escrever estatísticas: {str(e)}")
                        
                        # Gerar gráficos (opcional)
                        if not only_comparative:
                            try:
                                criar_grafico_a_partir_do_excel(excel_path, outp, sim_index=idx, salvar_png=True, mostrar=False)
                                if save_logs:
                                    log_lines.append(f"  [OK] Gráfico individual gerado")
                            except Exception as e:
                                if save_logs:
                                    log_lines.append(f"  [AVISO] Falha ao gerar gráfico: {str(e)}")
                        else:
                            if save_logs:
                                log_lines.append(f"  [PULADO] Gráfico individual (modo comparativo ativado)")
                    
                    excel_paths.append(excel_path)
                    idx += 1
                    # progresso
                    pct = int(i * 100 / max(1, total))
                    self.progress_var.set(pct)
                
                # Gráfico comparativo (apenas para modo tradicional)
                if not self.cancel_event.is_set() and len(excel_paths) > 1 and not selected_variables:
                    self.status_var.set('Gerando gráfico comparativo…')
                    if save_logs:
                        log_lines.append(f"\n[COMPARATIVO] {len(excel_paths)} arquivos")
                    try:
                        criar_grafico_comparativo(excel_paths, outp, mostrar=show_plots)
                        if save_logs:
                            log_lines.append(f"  [OK] Gráfico comparativo gerado")
                    except Exception as e:
                        if save_logs:
                            log_lines.append(f"  [AVISO] Falha ao gerar comparativo: {str(e)}")
                
                # Salvar log
                if save_logs:
                    try:
                        log_path = outp / f"log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
                        log_lines.append(f"\n=== Concluído em {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} ===")
                        log_lines.append(f"Total processado: {len(excel_paths)}/{total}")
                        log_path.write_text('\n'.join(log_lines), encoding='utf-8')
                    except Exception as e:
                        log_lines.append(f"[AVISO] Erro ao salvar log: {str(e)}")
                
                # Finalização
                if self.cancel_event.is_set():
                    messagebox.showinfo('Cancelado', 'Processamento cancelado pelo usuário.')
                else:
                    self.status_var.set('Concluído!')
                    msg = f'Processamento concluído!\n\n'
                    msg += f'Arquivos processados: {len(excel_paths)}/{total}\n'
                    msg += f'Pasta: {outp}'
                    messagebox.showinfo('Concluído', msg)
                    
                    # Abrir pasta se solicitado
                    if open_output:
                        _open_in_file_manager(outp)
                        
            except Exception:
                err_msg = traceback.format_exc()
                if save_logs:
                    log_lines.append(f"\n[ERRO CRÍTICO]\n{err_msg}")
                messagebox.showerror('Erro', err_msg)
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
        self.style.configure('Card.TLabelframe.Label', background=BG, foreground=HEADER_FG, font=('TkDefaultFont', 10, 'bold'))
        self.style.configure('TLabelframe', background=BG)
        self.style.configure('TLabelframe.Label', background=BG, foreground=HEADER_FG, font=('TkDefaultFont', 10, 'bold'))

        # labels e entradas
        self.style.configure('TLabel', background=BG, foreground=TEXT)
        self.style.configure('TEntry', fieldbackground=BG, foreground=TEXT, borderwidth=1)
        try:
            self.style.map('TEntry', fieldbackground=[('disabled', '#f0f0f0')])
        except Exception:
            pass

        # botões (melhorado)
        self.style.configure('TButton', background=ACCENT_BG, foreground=TEXT, borderwidth=1, relief='raised', padding=8)
        try:
            self.style.map('TButton', 
                          background=[('active', '#d6eaff'), ('disabled', '#f0f0f0')],
                          foreground=[('disabled', '#999999')])
        except Exception:
            pass

        # checkbuttons (novo)
        self.style.configure('TCheckbutton', background=BG, foreground=TEXT)
        try:
            self.style.map('TCheckbutton', background=[('active', ACCENT_BG)])
        except Exception:
            pass

        # treeview
        self.style.configure('Treeview', background=BG, fieldbackground=BG, foreground=TEXT, rowheight=24, borderwidth=1)
        self.style.map('Treeview', background=[('selected', SEL_BG)], foreground=[('selected', TEXT)])
        self.style.configure('Treeview.Heading', background=HEADER_BG, foreground=HEADER_FG, borderwidth=1)
        try:
            self.style.map('Treeview.Heading', background=[('active', '#d6eaff')])
        except Exception:
            pass

        # scrollbars
        try:
            self.style.configure('Vertical.TScrollbar', background=BG, troughcolor=ACCENT_BG)
            self.style.configure('Horizontal.TScrollbar', background=BG, troughcolor=ACCENT_BG)
        except Exception:
            pass

        # progressbar (melhorado)
        self.style.configure('Blue.Horizontal.TProgressbar', troughcolor=ACCENT_BG, background=PROGRESS, borderwidth=0, relief='flat')

        # Spinbox (novo)
        try:
            self.style.configure('TSpinbox', fieldbackground=BG, foreground=TEXT, borderwidth=1)
        except Exception:
            pass


def launch_gui(folder: Path, outdir: Path, start_index: int = 1):
    """Ponto de entrada público mantendo assinatura original."""
    folder = Path(folder)
    outdir = Path(outdir)
    root = tk.Tk()
    app = LisAnalysisApp(root, folder, outdir, start_index)
    root.protocol('WM_DELETE_WINDOW', lambda: (app._save_prefs(), root.destroy()))
    root.mainloop()
