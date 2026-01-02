import threading
import traceback
import json
import sys
import os
import subprocess
from datetime import datetime
from pathlib import Path

# IMPORTANTE: Configurar matplotlib ANTES de importar customtkinter
import matplotlib
matplotlib.use('TkAgg')  # Backend com GUI para evitar conflito com customtkinter

import customtkinter as ctk
from tkinter import filedialog, messagebox
import tkinter as tk

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
    raise

# Configurações globais do CustomTkinter
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

PREFS_FILE = Path.home() / ".lis_analysis_gui.json"


def _scan_lis(folder: Path):
    """Retorna arquivos .lis/.LIS ordenados por modificação (desc)."""
    folder = Path(folder)
    files = list(folder.glob('*.lis')) + list(folder.glob('*.LIS'))
    seen = set()
    unique_files = []
    for f in files:
        key = str(f).lower()
        if key in seen:
            continue
        seen.add(key)
        unique_files.append(f)
    try:
        unique_files.sort(key=lambda f: f.stat().st_mtime, reverse=True)
    except Exception:
        unique_files.sort()
    return unique_files


def _scan_acp(folder: Path):
    """Retorna arquivos .acp/.ACP ordenados por modificação (desc)."""
    folder = Path(folder)
    files = list(folder.glob('*.acp')) + list(folder.glob('*.ACP'))
    seen = set()
    unique_files = []
    for f in files:
        key = str(f).lower()
        if key in seen:
            continue
        seen.add(key)
        unique_files.append(f)
    try:
        unique_files.sort(key=lambda f: f.stat().st_mtime, reverse=True)
    except Exception:
        unique_files.sort()
    return unique_files


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
            os.startfile(str(path))
        else:
            messagebox.showinfo('Abrir pasta', f'Abra manualmente: {path}')
    except Exception as e:
        messagebox.showerror('Erro ao abrir', str(e))


class ModernLisAnalysisApp(ctk.CTk):
    def __init__(self, folder: Path, outdir: Path, start_index: int = 1):
        super().__init__()
        
        # Configurações da janela
        self.title("LIS Analysis")
        self.geometry("1200x800")
        
        # Variáveis de estado
        self.folder_var = tk.StringVar(value=str(folder))
        self.outdir_var = tk.StringVar(value=str(outdir))
        self.start_idx_var = tk.IntVar(value=start_index)
        self.filter_var = tk.StringVar()
        self.filetype_var = tk.StringVar(value='.lis')
        self.status_var = tk.StringVar(value='Pronto')
        self.progress_var = tk.DoubleVar(value=0)
        self.total_var = tk.IntVar(value=0)
        self.cancel_event = threading.Event()
        self._files_cache = []
        self._sort_desc = False
        self._sort_col = 'nome'
        
        # Opções (checkboxes)
        self.show_plots_var = tk.BooleanVar(value=False)
        self.open_output_var = tk.BooleanVar(value=True)
        self.only_comparative_var = tk.BooleanVar(value=False)
        self.save_logs_var = tk.BooleanVar(value=True)
        self.overwrite_var = tk.BooleanVar(value=True)
        self.hide_errors_var = tk.BooleanVar(value=False)
        self.parallel_process_var = tk.BooleanVar(value=False)
        self.auto_organize_var = tk.BooleanVar(value=True)
        
        # Variáveis ATP
        self.atp_executable_var = tk.StringVar(value='')
        self.acp_file_var = tk.StringVar(value='')
        
        # Carregar preferências
        self._load_prefs()
        
        # Construir interface
        self._build_ui()
        
        # Protocolo de fechamento
        self.protocol("WM_DELETE_WINDOW", self._on_closing)
        
        # Refresh inicial
        self.after(100, self.refresh_list)
    
    def _load_prefs(self):
        """Carrega preferências salvas"""
        try:
            if PREFS_FILE.exists():
                data = json.load(PREFS_FILE.open('r', encoding='utf-8'))
                self.folder_var.set(data.get('folder', self.folder_var.get()))
                self.outdir_var.set(data.get('outdir', self.outdir_var.get()))
                self.show_plots_var.set(data.get('show_plots', False))
                self.open_output_var.set(data.get('open_output', True))
                self.only_comparative_var.set(data.get('only_comparative', False))
                self.save_logs_var.set(data.get('save_logs', True))
                self.overwrite_var.set(data.get('overwrite', True))
                self.hide_errors_var.set(data.get('hide_errors', False))
                self.parallel_process_var.set(data.get('parallel_process', False))
                self.auto_organize_var.set(data.get('auto_organize', True))
                self.atp_executable_var.set(data.get('atp_executable', ''))
                
                # Carregar tema
                appearance = data.get('appearance_mode', 'System')
                ctk.set_appearance_mode(appearance)
        except Exception:
            pass
    
    def _save_prefs(self):
        """Salva preferências"""
        try:
            data = {
                'folder': self.folder_var.get(),
                'outdir': self.outdir_var.get(),
                'show_plots': self.show_plots_var.get(),
                'open_output': self.open_output_var.get(),
                'only_comparative': self.only_comparative_var.get(),
                'save_logs': self.save_logs_var.get(),
                'overwrite': self.overwrite_var.get(),
                'hide_errors': self.hide_errors_var.get(),
                'parallel_process': self.parallel_process_var.get(),
                'auto_organize': self.auto_organize_var.get(),
                'atp_executable': self.atp_executable_var.get(),
                'appearance_mode': ctk.get_appearance_mode(),
            }
            PREFS_FILE.write_text(json.dumps(data, indent=2), encoding='utf-8')
        except Exception:
            pass
    
    def _on_closing(self):
        """Handler para fechamento da janela"""
        self._save_prefs()
        self.destroy()
    
    def _build_ui(self):
        """Constrói a interface principal com abas"""
        # Barra superior com controles globais
        self._build_header()
        
        # TabView principal
        self.tabview = ctk.CTkTabview(self, width=1150, height=650)
        self.tabview.pack(pady=20, padx=20, fill="both", expand=True)
        
        # Criar abas (SEM EMOJIS para evitar segfault em alguns sistemas)
        self.tabview.add("Configuracoes")
        self.tabview.add("Analise .lis")
        self.tabview.add("Simulacao ATP")
        self.tabview.add("Logs")
        
        # Popular cada aba
        self._build_config_tab()
        self._build_analysis_tab()
        self._build_simulation_tab()
        self._build_logs_tab()
        
        # Barra de status inferior
        self._build_status_bar()
    
    def _build_header(self):
        """Constrói o cabeçalho com tema e título"""
        header = ctk.CTkFrame(self, height=60, fg_color="transparent")
        header.pack(fill="x", padx=20, pady=(20, 0))
        
        # Título
        title_label = ctk.CTkLabel(
            header, 
            text="LIS Analysis",
            font=ctk.CTkFont(size=24, weight="bold")
        )
        title_label.pack(side="left", padx=10)
        
        # Toggle de tema (Dark/Light)
        theme_frame = ctk.CTkFrame(header, fg_color="transparent")
        theme_frame.pack(side="right", padx=10)
        
        ctk.CTkLabel(theme_frame, text="Tema:").pack(side="left", padx=5)
        
        theme_menu = ctk.CTkOptionMenu(
            theme_frame,
            values=["System", "Light", "Dark"],
            command=self._change_appearance_mode,
            width=120
        )
        theme_menu.set(ctk.get_appearance_mode())
        theme_menu.pack(side="left")
    
    def _change_appearance_mode(self, new_mode: str):
        """Altera o tema (claro/escuro)"""
        ctk.set_appearance_mode(new_mode)
        self._save_prefs()
    
    def _build_config_tab(self):
        """Aba de Configurações"""
        tab = self.tabview.tab("Configuracoes")
        
        # Frame scrollable
        scroll_frame = ctk.CTkScrollableFrame(tab, width=1100, height=550)
        scroll_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Card: Diretórios
        dir_card = ctk.CTkFrame(scroll_frame, corner_radius=10)
        dir_card.pack(fill="x", pady=(0, 15))
        
        ctk.CTkLabel(
            dir_card, 
            text="Diretórios",
            font=ctk.CTkFont(size=16, weight="bold")
        ).pack(anchor="w", padx=15, pady=(15, 10))
        
        # Pasta de entrada (.lis)
        ctk.CTkLabel(dir_card, text="Pasta de entrada (.lis):").pack(anchor="w", padx=15, pady=(5, 0))
        folder_frame = ctk.CTkFrame(dir_card, fg_color="transparent")
        folder_frame.pack(fill="x", padx=15, pady=(5, 10))
        
        self.folder_entry = ctk.CTkEntry(folder_frame, textvariable=self.folder_var, width=900)
        self.folder_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
        
        ctk.CTkButton(
            folder_frame, 
            text="Escolher", 
            command=self._choose_folder,
            width=120
        ).pack(side="left")
        
        # Pasta de saída
        ctk.CTkLabel(dir_card, text="Pasta de saída (Excel/gráficos):").pack(anchor="w", padx=15, pady=(5, 0))
        outdir_frame = ctk.CTkFrame(dir_card, fg_color="transparent")
        outdir_frame.pack(fill="x", padx=15, pady=(5, 15))
        
        self.outdir_entry = ctk.CTkEntry(outdir_frame, textvariable=self.outdir_var, width=900)
        self.outdir_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
        
        ctk.CTkButton(
            outdir_frame, 
            text="Escolher", 
            command=self._choose_outdir,
            width=120
        ).pack(side="left")
        
        # Card: Opções de Processamento
        options_card = ctk.CTkFrame(scroll_frame, corner_radius=10)
        options_card.pack(fill="x", pady=(0, 15))
        
        ctk.CTkLabel(
            options_card, 
            text="Opções de Processamento",
            font=ctk.CTkFont(size=16, weight="bold")
        ).pack(anchor="w", padx=15, pady=(15, 10))
        
        # Grid de checkboxes
        checks_frame = ctk.CTkFrame(options_card, fg_color="transparent")
        checks_frame.pack(fill="x", padx=15, pady=(0, 15))
        
        # Coluna 1
        col1 = ctk.CTkFrame(checks_frame, fg_color="transparent")
        col1.pack(side="left", fill="both", expand=True, padx=(0, 10))
        
        ctk.CTkCheckBox(col1, text="Mostrar gráficos", variable=self.show_plots_var).pack(anchor="w", pady=5)
        ctk.CTkCheckBox(col1, text="Abrir pasta de saída", variable=self.open_output_var).pack(anchor="w", pady=5)
        ctk.CTkCheckBox(col1, text="Só gráfico comparativo", variable=self.only_comparative_var).pack(anchor="w", pady=5)
        ctk.CTkCheckBox(col1, text="Salvar logs", variable=self.save_logs_var).pack(anchor="w", pady=5)
        
        # Coluna 2
        col2 = ctk.CTkFrame(checks_frame, fg_color="transparent")
        col2.pack(side="left", fill="both", expand=True)
        
        ctk.CTkCheckBox(col2, text="Sobrescrever arquivos", variable=self.overwrite_var).pack(anchor="w", pady=5)
        ctk.CTkCheckBox(col2, text="Ocultar erros individuais", variable=self.hide_errors_var).pack(anchor="w", pady=5)
        ctk.CTkCheckBox(col2, text="Processar em paralelo", variable=self.parallel_process_var).pack(anchor="w", pady=5)
        ctk.CTkCheckBox(col2, text="Auto-organizar arquivos", variable=self.auto_organize_var).pack(anchor="w", pady=5)
        
        # Card: Índice Inicial
        index_card = ctk.CTkFrame(scroll_frame, corner_radius=10)
        index_card.pack(fill="x", pady=(0, 15))
        
        ctk.CTkLabel(
            index_card, 
            text="Configurações Avançadas",
            font=ctk.CTkFont(size=16, weight="bold")
        ).pack(anchor="w", padx=15, pady=(15, 10))
        
        index_frame = ctk.CTkFrame(index_card, fg_color="transparent")
        index_frame.pack(anchor="w", padx=15, pady=(0, 15))
        
        ctk.CTkLabel(index_frame, text="Índice inicial:").pack(side="left", padx=(0, 10))
        ctk.CTkEntry(index_frame, textvariable=self.start_idx_var, width=80).pack(side="left")
    
    def _build_analysis_tab(self):
        """Aba de Análise de Arquivos .lis"""
        tab = self.tabview.tab("Analise .lis")
        
        # Frame superior: filtros e botões
        top_frame = ctk.CTkFrame(tab, fg_color="transparent")
        top_frame.pack(fill="x", padx=10, pady=10)
        
        # Filtro
        filter_frame = ctk.CTkFrame(top_frame, fg_color="transparent")
        filter_frame.pack(side="left", fill="x", expand=True)
        
        ctk.CTkLabel(filter_frame, text="Filtro:").pack(side="left", padx=(0, 10))
        filter_entry = ctk.CTkEntry(filter_frame, textvariable=self.filter_var, width=300)
        filter_entry.pack(side="left", fill="x", expand=True)
        filter_entry.bind('<KeyRelease>', lambda e: self.refresh_list())
        
        ctk.CTkButton(
            filter_frame, 
            text="Limpar", 
            command=self._clear_filter,
            width=80,
            fg_color="gray"
        ).pack(side="left", padx=10)
        
        # Botões de ação
        action_frame = ctk.CTkFrame(top_frame, fg_color="transparent")
        action_frame.pack(side="right")
        
        ctk.CTkButton(
            action_frame, 
            text="Atualizar", 
            command=self.refresh_list,
            width=120
        ).pack(side="left", padx=5)
        
        ctk.CTkButton(
            action_frame, 
            text="Processar Selecionados", 
            command=self._process_selected,
            width=180,
            fg_color="#4CAF50",
            hover_color="#45a049"
        ).pack(side="left", padx=5)
        
        # Frame da lista de arquivos (scrollable)
        self.files_scroll_frame = ctk.CTkScrollableFrame(tab, width=1100, height=520)
        self.files_scroll_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        
        # Dicionário para rastrear checkboxes dos arquivos
        self.file_checkboxes = {}
        self.file_selection_vars = {}
    
    def _build_simulation_tab(self):
        """Aba de Simulação ATP"""
        tab = self.tabview.tab("Simulacao ATP")
        
        # Frame scrollable
        scroll_frame = ctk.CTkScrollableFrame(tab, width=1100, height=550)
        scroll_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Card: Executável ATP
        exec_card = ctk.CTkFrame(scroll_frame, corner_radius=10)
        exec_card.pack(fill="x", pady=(0, 15))
        
        ctk.CTkLabel(
            exec_card, 
            text="Executável ATP",
            font=ctk.CTkFont(size=16, weight="bold")
        ).pack(anchor="w", padx=15, pady=(15, 10))
        
        ctk.CTkLabel(exec_card, text="Caminho para tpbig.exe ou runATP.bat:").pack(anchor="w", padx=15, pady=(5, 0))
        exec_frame = ctk.CTkFrame(exec_card, fg_color="transparent")
        exec_frame.pack(fill="x", padx=15, pady=(5, 15))
        
        self.atp_exec_entry = ctk.CTkEntry(exec_frame, textvariable=self.atp_executable_var, width=900)
        self.atp_exec_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
        
        ctk.CTkButton(
            exec_frame, 
            text="Escolher", 
            command=self._choose_atp_executable,
            width=120
        ).pack(side="left")
        
        # Card: Arquivo .acp
        acp_card = ctk.CTkFrame(scroll_frame, corner_radius=10)
        acp_card.pack(fill="x", pady=(0, 15))
        
        ctk.CTkLabel(
            acp_card, 
            text="Arquivo .acp",
            font=ctk.CTkFont(size=16, weight="bold")
        ).pack(anchor="w", padx=15, pady=(15, 10))
        
        ctk.CTkLabel(acp_card, text="Arquivo .acp para simular:").pack(anchor="w", padx=15, pady=(5, 0))
        acp_frame = ctk.CTkFrame(acp_card, fg_color="transparent")
        acp_frame.pack(fill="x", padx=15, pady=(5, 15))
        
        self.acp_file_entry = ctk.CTkEntry(acp_frame, textvariable=self.acp_file_var, width=900)
        self.acp_file_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
        
        ctk.CTkButton(
            acp_frame, 
            text="Escolher", 
            command=self._choose_acp_file,
            width=120
        ).pack(side="left")
        
        # Card: Ações
        action_card = ctk.CTkFrame(scroll_frame, corner_radius=10)
        action_card.pack(fill="x", pady=(0, 15))
        
        ctk.CTkLabel(
            action_card, 
            text="Executar Simulação",
            font=ctk.CTkFont(size=16, weight="bold")
        ).pack(anchor="w", padx=15, pady=(15, 10))
        
        buttons_frame = ctk.CTkFrame(action_card, fg_color="transparent")
        buttons_frame.pack(padx=15, pady=(0, 15))
        
        ctk.CTkButton(
            buttons_frame, 
            text="Executar ATP", 
            command=self._run_atp_simulation,
            width=200,
            height=40,
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color="#2196F3",
            hover_color="#1976D2"
        ).pack(side="left", padx=5)
        
        ctk.CTkButton(
            buttons_frame, 
            text="Detectar Parâmetros", 
            command=self._detect_parameters,
            width=200,
            height=40,
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color="#FF9800",
            hover_color="#F57C00"
        ).pack(side="left", padx=5)
        
        # Área de resultados
        self.simulation_results = ctk.CTkTextbox(scroll_frame, width=1100, height=200)
        self.simulation_results.pack(fill="both", expand=True, pady=(0, 10))
        self.simulation_results.insert("1.0", "Resultados da simulação aparecerão aqui...\n")
        self.simulation_results.configure(state="disabled")
    
    def _build_logs_tab(self):
        """Aba de Logs"""
        tab = self.tabview.tab("Logs")
        
        # Toolbar
        toolbar = ctk.CTkFrame(tab, fg_color="transparent")
        toolbar.pack(fill="x", padx=10, pady=10)
        
        ctk.CTkButton(
            toolbar, 
            text="Limpar Logs", 
            command=self._clear_logs,
            width=120,
            fg_color="#f44336",
            hover_color="#d32f2f"
        ).pack(side="left")
        
        ctk.CTkButton(
            toolbar, 
            text="Salvar Logs", 
            command=self._save_logs_to_file,
            width=120
        ).pack(side="left", padx=10)
        
        # Área de texto para logs
        self.log_textbox = ctk.CTkTextbox(tab, width=1100, height=550)
        self.log_textbox.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        
        # Log inicial
        self.log("Interface inicializada com sucesso")
        self.log(f"Pasta de entrada: {self.folder_var.get()}")
        self.log(f"Pasta de saída: {self.outdir_var.get()}")
    
    def _build_status_bar(self):
        """Barra de status inferior"""
        status_frame = ctk.CTkFrame(self, height=50)
        status_frame.pack(fill="x", side="bottom", padx=20, pady=(0, 20))
        
        # Status text
        self.status_label = ctk.CTkLabel(
            status_frame, 
            textvariable=self.status_var,
            font=ctk.CTkFont(size=12)
        )
        self.status_label.pack(side="left", padx=15)
        
        # Progress bar
        self.progress_bar = ctk.CTkProgressBar(status_frame, width=400)
        self.progress_bar.pack(side="left", padx=15)
        self.progress_bar.set(0)
        
        # Botão de cancelar (oculto inicialmente)
        self.cancel_btn = ctk.CTkButton(
            status_frame, 
            text="Cancelar", 
            command=self._cancel_processing,
            width=100,
            fg_color="#f44336",
            hover_color="#d32f2f"
        )
        # Não empacotar ainda, só quando houver processamento
    
    # ========== MÉTODOS DE AÇÃO ==========
    
    def _choose_folder(self):
        """Escolher pasta de entrada"""
        folder = filedialog.askdirectory(title="Escolher pasta de entrada (.lis)")
        if folder:
            self.folder_var.set(folder)
            self.refresh_list()
            self._save_prefs()
    
    def _choose_outdir(self):
        """Escolher pasta de saída"""
        folder = filedialog.askdirectory(title="Escolher pasta de saída")
        if folder:
            self.outdir_var.set(folder)
            self._save_prefs()
    
    def _choose_atp_executable(self):
        """Escolher executável ATP"""
        file = filedialog.askopenfilename(
            title="Escolher executável ATP",
            filetypes=[("Executáveis", "*.exe *.bat *.cmd"), ("Todos", "*.*")]
        )
        if file:
            self.atp_executable_var.set(file)
            self._save_prefs()
    
    def _choose_acp_file(self):
        """Escolher arquivo .acp"""
        file = filedialog.askopenfilename(
            title="Escolher arquivo .acp",
            filetypes=[("Arquivos ACP", "*.acp *.ACP"), ("Todos", "*.*")]
        )
        if file:
            self.acp_file_var.set(file)
    
    def _clear_filter(self):
        """Limpar filtro"""
        self.filter_var.set('')
        self.refresh_list()
    
    def _cancel_processing(self):
        """Cancelar processamento em andamento"""
        self.cancel_event.set()
        self.log("Cancelamento solicitado...")
    
    def _clear_logs(self):
        """Limpar área de logs"""
        self.log_textbox.configure(state="normal")
        self.log_textbox.delete("1.0", "end")
        self.log_textbox.configure(state="disabled")
        self.log("Logs limpos")
    
    def _save_logs_to_file(self):
        """Salvar logs em arquivo"""
        file = filedialog.asksaveasfilename(
            title="Salvar logs",
            defaultextension=".txt",
            filetypes=[("Arquivo de texto", "*.txt"), ("Todos", "*.*")]
        )
        if file:
            try:
                content = self.log_textbox.get("1.0", "end")
                Path(file).write_text(content, encoding='utf-8')
                messagebox.showinfo("Sucesso", f"Logs salvos em:\n{file}")
            except Exception as e:
                messagebox.showerror("Erro", f"Falha ao salvar logs:\n{e}")
    
    def log(self, message: str):
        """Adiciona mensagem ao log"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        full_msg = f"[{timestamp}] {message}\n"
        self.log_textbox.configure(state="normal")
        self.log_textbox.insert("end", full_msg)
        self.log_textbox.see("end")
        self.log_textbox.configure(state="disabled")
    
    def refresh_list(self):
        """Atualiza lista de arquivos .lis"""
        folder = Path(self.folder_var.get())
        if not folder.exists():
            self.log(f"Pasta não encontrada: {folder}")
            return
        
        # Escanear arquivos
        files = _scan_lis(folder)
        filter_text = self.filter_var.get().lower()
        if filter_text:
            files = [f for f in files if filter_text in f.name.lower()]
        
        self._files_cache = files
        self.log(f"{len(files)} arquivo(s) .lis encontrado(s)")
        
        # Limpar lista anterior
        for widget in self.files_scroll_frame.winfo_children():
            widget.destroy()
        self.file_checkboxes.clear()
        self.file_selection_vars.clear()
        
        # Popular lista com cards de arquivos
        if not files:
            no_files_label = ctk.CTkLabel(
                self.files_scroll_frame,
                text="Nenhum arquivo .lis encontrado nesta pasta",
                font=ctk.CTkFont(size=14),
                text_color="gray"
            )
            no_files_label.pack(pady=50)
            return
        
        for file_path in files:
            self._create_file_card(file_path)
    
    def _create_file_card(self, file_path: Path):
        """Cria um card para representar um arquivo .lis"""
        # Card principal
        card = ctk.CTkFrame(self.files_scroll_frame, corner_radius=8, fg_color=("gray85", "gray20"))
        card.pack(fill="x", padx=5, pady=3)
        
        # Frame interno com checkbox e info
        content_frame = ctk.CTkFrame(card, fg_color="transparent")
        content_frame.pack(fill="x", padx=10, pady=8)
        
        # Checkbox para seleção
        var = tk.BooleanVar(value=False)
        self.file_selection_vars[str(file_path)] = var
        
        checkbox = ctk.CTkCheckBox(
            content_frame,
            text="",
            variable=var,
            width=20
        )
        checkbox.pack(side="left", padx=(0, 10))
        
        # Informações do arquivo
        info_frame = ctk.CTkFrame(content_frame, fg_color="transparent")
        info_frame.pack(side="left", fill="x", expand=True)
        
        # Nome do arquivo
        name_label = ctk.CTkLabel(
            info_frame,
            text=f"{file_path.name}",
            font=ctk.CTkFont(size=13, weight="bold"),
            anchor="w"
        )
        name_label.pack(anchor="w", fill="x")
        
        # Detalhes (tamanho e data)
        try:
            file_size = _fmt_size(file_path.stat().st_size)
            file_mtime = datetime.fromtimestamp(file_path.stat().st_mtime).strftime("%d/%m/%Y %H:%M")
            details = f"{file_size}  •  {file_mtime}"
        except Exception:
            details = "Tamanho desconhecido"
        
        details_label = ctk.CTkLabel(
            info_frame,
            text=details,
            font=ctk.CTkFont(size=11),
            text_color="gray60",
            anchor="w"
        )
        details_label.pack(anchor="w", fill="x")
        
        # Botão de abrir/visualizar
        btn_frame = ctk.CTkFrame(content_frame, fg_color="transparent")
        btn_frame.pack(side="right", padx=(10, 0))
        
        open_btn = ctk.CTkButton(
            btn_frame,
            text="Abrir",
            width=80,
            height=28,
            command=lambda: self._open_file_in_editor(file_path)
        )
        open_btn.pack()
        
        self.file_checkboxes[str(file_path)] = checkbox
    
    def _open_file_in_editor(self, file_path: Path):
        """Abre arquivo .lis no editor do sistema"""
        try:
            if sys.platform.startswith('linux'):
                subprocess.Popen(['xdg-open', str(file_path)])
            elif sys.platform == 'darwin':
                subprocess.Popen(['open', str(file_path)])
            elif os.name == 'nt':
                os.startfile(str(file_path))
            else:
                messagebox.showinfo('Abrir arquivo', f'Abra manualmente: {file_path}')
        except Exception as e:
            messagebox.showerror('Erro ao abrir', str(e))
    
    def _process_selected(self):
        """Processar arquivos .lis selecionados"""
        # Coletar arquivos selecionados
        selected_files = []
        for file_str, var in self.file_selection_vars.items():
            if var.get():
                selected_files.append(Path(file_str))
        
        if not selected_files:
            messagebox.showwarning("Aviso", "Nenhum arquivo selecionado.\n\nMarque os arquivos que deseja processar.")
            return
        
        self.log(f"Iniciando processamento de {len(selected_files)} arquivo(s)...")
        self.status_var.set(f"Processando {len(selected_files)} arquivo(s)...")
        
        # Mostrar botão de cancelar
        self.cancel_btn.pack(side="right", padx=15)
        
        def worker():
            try:
                outdir = Path(self.outdir_var.get())
                outdir.mkdir(parents=True, exist_ok=True)
                
                total = len(selected_files)
                for idx, lis_path in enumerate(selected_files, start=self.start_idx_var.get()):
                    if self.cancel_event.is_set():
                        self.log("Processamento cancelado pelo usuário")
                        break
                    
                    # Atualizar progresso
                    progress = (idx - self.start_idx_var.get() + 1) / total
                    self.progress_bar.set(progress)
                    self.status_var.set(f"Processando {lis_path.name}... ({idx - self.start_idx_var.get() + 1}/{total})")
                    
                    self.log(f"Processando: {lis_path.name}")
                    
                    # Parse do .lis
                    df, stats_lines, summary = parse_lis_table(lis_path)
                    if df is None:
                        self.log(f"Tabela não encontrada em: {lis_path.name}")
                        continue
                    
                    # Salvar Excel
                    excel_path = outdir / f"Resultados_Simulacao_{idx}.xlsx"
                    save_df_to_excel_only(df, excel_path)
                    
                    # Estatísticas
                    try:
                        computed_stats = calcular_estatisticas_do_df(df)
                        escrever_estatisticas_excel(excel_path, computed_stats, summary_from_lis=summary)
                    except Exception as e:
                        self.log(f"Erro ao calcular estatísticas: {e}")
                    
                    # Gráfico
                    if self.show_plots_var.get() or not self.only_comparative_var.get():
                        criar_grafico_a_partir_do_excel(excel_path, outdir, sim_index=idx, salvar_png=True, mostrar=self.show_plots_var.get())
                    
                    # Séries temporais
                    time_series_df = parse_lis_time_series(lis_path)
                    if time_series_df is not None:
                        save_time_series_to_excel(time_series_df, excel_path)
                        criar_grafico_series_temporais(time_series_df, outdir / f"series_temporais_{idx}.png", lis_name=lis_path.name)
                    
                    self.log(f"Concluído: {lis_path.name}")
                
                # Finalizar
                self.progress_bar.set(1.0)
                self.status_var.set(f"Processamento concluído! {len(selected_files)} arquivo(s)")
                self.log(f"Processamento finalizado com sucesso")
                
                if self.open_output_var.get():
                    _open_in_file_manager(outdir)
                
                messagebox.showinfo("Sucesso", f"Processamento concluído!\n\n{len(selected_files)} arquivo(s) processado(s)\nResultados em: {outdir}")
                
            except Exception as e:
                self.log(f"Erro durante processamento: {e}")
                messagebox.showerror("Erro", f"Erro durante processamento:\n{str(e)}")
            finally:
                self.cancel_btn.pack_forget()
                self.cancel_event.clear()
                self.progress_bar.set(0)
                if not self.cancel_event.is_set():
                    self.status_var.set("Pronto")
        
        threading.Thread(target=worker, daemon=True).start()
    
    def _run_atp_simulation(self):
        """Executar simulação ATP"""
        atp_exe = self.atp_executable_var.get()
        acp_file = self.acp_file_var.get()
        
        if not atp_exe or not Path(atp_exe).exists():
            messagebox.showerror("Erro", "Executável ATP não encontrado")
            return
        
        if not acp_file or not Path(acp_file).exists():
            messagebox.showerror("Erro", "Arquivo .acp não encontrado")
            return
        
        self.log("Iniciando simulação ATP...")
        
        def worker():
            try:
                runner = AtpRunner(atp_exe)
                result = runner.run_simulation(Path(acp_file))
                
                if result:
                    self.log(f"Simulação concluída: {result}")
                    self._update_simulation_results(f"Sucesso!\nArquivo .lis gerado: {result}")
                else:
                    self.log("Simulação falhou")
                    self._update_simulation_results("Falha na simulação. Verifique os logs.")
            except Exception as e:
                self.log(f"Erro: {e}")
                self._update_simulation_results(f"Erro:\n{str(e)}")
        
        threading.Thread(target=worker, daemon=True).start()
    
    def _detect_parameters(self):
        """Detectar parâmetros do arquivo .acp"""
        acp_file = self.acp_file_var.get()
        
        if not acp_file or not Path(acp_file).exists():
            messagebox.showerror("Erro", "Arquivo .acp não encontrado")
            return
        
        self.log("Detectando parâmetros...")
        
        try:
            parser = AcpParser(Path(acp_file))
            parser.extract_atp_from_acp()
            params = parser.find_control_parameters()
            
            result_text = "Parâmetros detectados:\n\n"
            result_text += f"• RPI values: {len(params.get('rpi_values', []))}\n"
            result_text += f"• Switch times: {len(params.get('switch_times', []))}\n"
            result_text += f"• dt: {params.get('dt')}\n"
            result_text += f"• tmax: {params.get('tmax')}\n"
            
            self._update_simulation_results(result_text)
            self.log("Parâmetros detectados com sucesso")
        except Exception as e:
            self.log(f"Erro ao detectar parâmetros: {e}")
            self._update_simulation_results(f"Erro:\n{str(e)}")
    
    def _update_simulation_results(self, text: str):
        """Atualiza área de resultados da simulação"""
        self.simulation_results.configure(state="normal")
        self.simulation_results.delete("1.0", "end")
        self.simulation_results.insert("1.0", text)
        self.simulation_results.configure(state="disabled")
    
    def _get_bg_color(self):
        """Retorna cor de fundo baseada no tema atual"""
        mode = ctk.get_appearance_mode()
        if mode == "Dark":
            return "#2b2b2b"
        else:
            return "#ffffff"


def launch_gui(folder: Path, outdir: Path, start_index: int = 1):
    """Ponto de entrada para a interface gráfica"""
    app = ModernLisAnalysisApp(folder, outdir, start_index)
    app.mainloop()


if __name__ == "__main__":
    # Teste standalone
    import sys
    folder = Path(sys.argv[1]) if len(sys.argv) > 1 else Path.cwd() / "ACP"
    outdir = Path(sys.argv[2]) if len(sys.argv) > 2 else Path.cwd() / "Simulation_Result"
    launch_gui(folder, outdir)
