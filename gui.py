import threading
import traceback
import json
import queue
import sys
import os
import subprocess
import time
from datetime import datetime
from pathlib import Path

# IMPORTANTE: Configurar matplotlib ANTES de importar customtkinter
import matplotlib
matplotlib.use('TkAgg')  # Backend com GUI para evitar conflito com customtkinter
import matplotlib.pyplot as plt
import numpy as np

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
    from solver.atp_runner import run_atp_solver
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
        self.hide_errors_var = tk.BooleanVar(value=False)
        self.parallel_process_var = tk.BooleanVar(value=False)

        # Estado da execucao ATP
        self._atp_running = False
        self._atp_started_at = None
        self._atp_worker_thread = None
        self._atp_result_queue = None
        self.atp_run_status_var = tk.StringVar(value="Status: aguardando execucao")

        # Simulacao ATP (.atp)
        self.atp_file_var = tk.StringVar(value='')
        
        
        # Opções de visualização de gráficos
        self.plot_bars_var = tk.BooleanVar(value=True)
        self.plot_points_var = tk.BooleanVar(value=True)
        self.plot_gaussian_var = tk.BooleanVar(value=True)
        self.plot_cumulative_var = tk.BooleanVar(value=True)
        self.plot_stats_box_var = tk.BooleanVar(value=True)
        
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
                self.hide_errors_var.set(data.get('hide_errors', False))
                self.parallel_process_var.set(data.get('parallel_process', False))
                self.atp_file_var.set(data.get('atp_file', ''))
                
                # Carregar opções de gráfico
                self.plot_bars_var.set(data.get('plot_bars', True))
                self.plot_points_var.set(data.get('plot_points', True))
                self.plot_gaussian_var.set(data.get('plot_gaussian', True))
                self.plot_cumulative_var.set(data.get('plot_cumulative', True))
                self.plot_stats_box_var.set(data.get('plot_stats_box', True))
                
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
                'hide_errors': self.hide_errors_var.get(),
                'parallel_process': self.parallel_process_var.get(),
                'atp_file': self.atp_file_var.get(),
                'plot_bars': self.plot_bars_var.get(),
                'plot_points': self.plot_points_var.get(),
                'plot_gaussian': self.plot_gaussian_var.get(),
                'plot_cumulative': self.plot_cumulative_var.get(),
                'plot_stats_box': self.plot_stats_box_var.get(),
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
        
        ctk.CTkCheckBox(col2, text="Ocultar erros individuais", variable=self.hide_errors_var).pack(anchor="w", pady=5)
        
        # Card: Opções de Visualização de Gráficos
        plot_card = ctk.CTkFrame(scroll_frame, corner_radius=10)
        plot_card.pack(fill="x", pady=(0, 15))
        
        ctk.CTkLabel(
            plot_card,
            text="Opções de Visualização dos Gráficos",
            font=ctk.CTkFont(size=16, weight="bold")
        ).pack(anchor="w", padx=15, pady=(15, 10))
        
        ctk.CTkLabel(
            plot_card,
            text="Selecione os elementos que devem aparecer nos gráficos:",
            font=ctk.CTkFont(size=11)
        ).pack(anchor="w", padx=15, pady=(0, 10))
        
        plot_checks_frame = ctk.CTkFrame(plot_card, fg_color="transparent")
        plot_checks_frame.pack(fill="x", padx=15, pady=(0, 15))
        
        plot_col1 = ctk.CTkFrame(plot_checks_frame, fg_color="transparent")
        plot_col1.pack(side="left", fill="both", expand=True, padx=(0, 10))
        
        ctk.CTkCheckBox(plot_col1, text="Barras (histograma)", variable=self.plot_bars_var).pack(anchor="w", pady=5)
        ctk.CTkCheckBox(plot_col1, text="Pontos (scatter)", variable=self.plot_points_var).pack(anchor="w", pady=5)
        ctk.CTkCheckBox(plot_col1, text="Curva gaussiana", variable=self.plot_gaussian_var).pack(anchor="w", pady=5)
        
        plot_col2 = ctk.CTkFrame(plot_checks_frame, fg_color="transparent")
        plot_col2.pack(side="left", fill="both", expand=True)
        
        ctk.CTkCheckBox(plot_col2, text="Curva acumulada (%)", variable=self.plot_cumulative_var).pack(anchor="w", pady=5)
        ctk.CTkCheckBox(plot_col2, text="Caixa de estatísticas", variable=self.plot_stats_box_var).pack(anchor="w", pady=5)
        
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
            text="Limpar Pasta Saida", 
            command=self._clear_output_folder,
            width=140,
            fg_color="#FF5722",
            hover_color="#E64A19"
        ).pack(side="left", padx=5)
        
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
        """Aba de Simulacao ATP"""
        tab = self.tabview.tab("Simulacao ATP")

        scroll_frame = ctk.CTkScrollableFrame(tab, width=1100, height=550)
        scroll_frame.pack(fill="both", expand=True, padx=10, pady=10)

        atp_card = ctk.CTkFrame(scroll_frame, corner_radius=10)
        atp_card.pack(fill="x", pady=(0, 15))

        ctk.CTkLabel(
            atp_card,
            text="Arquivo .atp",
            font=ctk.CTkFont(size=16, weight="bold")
        ).pack(anchor="w", padx=15, pady=(15, 10))

        ctk.CTkLabel(atp_card, text="Arquivo .atp base para simular:").pack(anchor="w", padx=15, pady=(5, 0))
        atp_frame = ctk.CTkFrame(atp_card, fg_color="transparent")
        atp_frame.pack(fill="x", padx=15, pady=(5, 15))

        self.atp_file_entry = ctk.CTkEntry(atp_frame, textvariable=self.atp_file_var, width=900)
        self.atp_file_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))

        ctk.CTkButton(
            atp_frame,
            text="Escolher",
            command=self._choose_atp_file,
            width=120
        ).pack(side="left")

        info_card = ctk.CTkFrame(scroll_frame, corner_radius=10)
        info_card.pack(fill="x", pady=(0, 15))

        ctk.CTkLabel(
            info_card,
            text="Status da Integracao ATP",
            font=ctk.CTkFont(size=16, weight="bold")
        ).pack(anchor="w", padx=15, pady=(15, 10))

        ctk.CTkLabel(
            info_card,
            text="A simulacao usa runATP.exe no mesmo diretorio do arquivo .atp e aguarda o termino da execucao.",
            justify="left"
        ).pack(anchor="w", padx=15, pady=(0, 15))

        action_card = ctk.CTkFrame(scroll_frame, corner_radius=10)
        action_card.pack(fill="x", pady=(0, 15))

        ctk.CTkLabel(
            action_card,
            text="Executar Simulacao",
            font=ctk.CTkFont(size=16, weight="bold")
        ).pack(anchor="w", padx=15, pady=(15, 10))

        buttons_frame = ctk.CTkFrame(action_card, fg_color="transparent")
        buttons_frame.pack(padx=15, pady=(0, 15))

        self.atp_run_button = ctk.CTkButton(
            buttons_frame,
            text="Run Simulation",
            command=self._run_atp_simulation,
            width=200,
            height=40,
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color="#2196F3",
            hover_color="#1976D2"
        )
        self.atp_run_button.pack(side="left", padx=5)

        self.atp_progress = ctk.CTkProgressBar(action_card, mode="indeterminate")
        self.atp_progress.pack(fill="x", padx=15, pady=(0, 8))
        self.atp_progress.stop()

        ctk.CTkLabel(
            action_card,
            textvariable=self.atp_run_status_var,
            font=ctk.CTkFont(size=12)
        ).pack(anchor="w", padx=15, pady=(0, 12))

        self.simulation_results = ctk.CTkTextbox(scroll_frame, width=1100, height=160)
        self.simulation_results.pack(fill="both", expand=True, pady=(0, 10))
        self.simulation_results.insert("1.0", "Aguardando execucao da simulacao ATP...\n")
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

    def _choose_atp_file(self):
        """Escolher arquivo .atp"""
        file = filedialog.askopenfilename(
            title="Escolher arquivo .atp",
            filetypes=[("Arquivos ATP", "*.atp *.ATP"), ("Todos", "*.*")]
        )
        if file:
            self.atp_file_var.set(file)
            self._save_prefs()
    
    
    def _clear_filter(self):
        """Limpar filtro"""
        self.filter_var.set('')
        self.refresh_list()
    
    def _clear_output_folder(self):
        """Limpar todos os arquivos da pasta de saída"""
        outdir = Path(self.outdir_var.get())
        
        if not outdir.exists():
            messagebox.showinfo("Informação", f"A pasta de saída não existe:\n{outdir}")
            return
        
        # Contar arquivos
        all_files = list(outdir.rglob('*'))
        files_count = sum(1 for f in all_files if f.is_file())
        dirs_count = sum(1 for f in all_files if f.is_dir())
        
        if files_count == 0 and dirs_count == 0:
            messagebox.showinfo("Informação", "A pasta de saída já está vazia.")
            return
        
        # Confirmação
        msg = f"Deseja EXCLUIR todos os arquivos da pasta de saída?\n\n"
        msg += f"Pasta: {outdir}\n"
        msg += f"Arquivos: {files_count}\n"
        msg += f"Subpastas: {dirs_count}\n\n"
        msg += "Esta ação NÃO PODE ser desfeita!"
        
        if not messagebox.askyesno("Confirmar Exclusão", msg, icon='warning'):
            return
        
        # Executar limpeza
        try:
            import shutil
            deleted_files = 0
            deleted_dirs = 0
            
            for item in outdir.iterdir():
                try:
                    if item.is_file():
                        item.unlink()
                        deleted_files += 1
                    elif item.is_dir():
                        shutil.rmtree(item)
                        deleted_dirs += 1
                except Exception as e:
                    self.log(f"Erro ao excluir {item.name}: {e}")
            
            self.log(f"Pasta de saída limpa: {deleted_files} arquivo(s), {deleted_dirs} pasta(s) excluída(s)")
            messagebox.showinfo("Sucesso", f"Pasta limpa com sucesso!\n\n{deleted_files} arquivo(s)\n{deleted_dirs} pasta(s)")
            
        except Exception as e:
            self.log(f"Erro ao limpar pasta: {e}")
            messagebox.showerror("Erro", f"Erro ao limpar pasta:\n{str(e)}")

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

    def _run_atp_simulation(self):
        """Executa o solver ATP em background e atualiza a GUI ao finalizar."""
        if self._atp_running:
            messagebox.showinfo("Simulacao ATP", "Ja existe uma simulacao ATP em andamento.")
            return

        atp_file = self.atp_file_var.get().strip()
        if not atp_file or not Path(atp_file).exists():
            messagebox.showerror("Erro", "Arquivo .atp nao encontrado")
            return

        self._set_atp_feedback_running()
        self.status_var.set("Executando simulacao ATP...")
        self.log(f"Iniciando simulacao ATP para: {atp_file}")
        self._update_simulation_results("Executando simulacao ATP... aguarde a conclusao do solver.\n")

        def worker():
            try:
                lis_path = run_atp_solver(atp_file)
                self._atp_result_queue.put(("success", lis_path))
            except Exception as e:
                self._atp_result_queue.put(("error", str(e)))

        self._atp_result_queue = queue.Queue()
        self._atp_worker_thread = threading.Thread(target=worker, daemon=True)
        self._atp_worker_thread.start()
        self.after(100, self._poll_atp_worker_result)

    def _poll_atp_worker_result(self):
        """Faz polling do resultado do worker ATP no thread principal."""
        if not self._atp_running or self._atp_result_queue is None:
            return

        try:
            result_type, payload = self._atp_result_queue.get_nowait()
        except queue.Empty:
            if self._atp_worker_thread is not None and self._atp_worker_thread.is_alive():
                self.after(100, self._poll_atp_worker_result)
                return

            # Worker terminou sem publicar resultado (falha inesperada)
            result_type, payload = "error", "Falha inesperada: worker ATP finalizou sem resultado."

        elapsed = self._set_atp_feedback_finished(success=(result_type == "success"))
        self.status_var.set("Pronto")

        if result_type == "success":
            lis_path = payload
            self.log(f"Simulacao concluida. LIS gerado em: {lis_path}")
            self._update_simulation_results(
                f"Simulation completed in {elapsed:.1f}s\nLIS file: {lis_path}"
            )
            messagebox.showinfo("Sucesso", f"Simulacao concluida em {elapsed:.1f}s.\n\nLIS file:\n{lis_path}")
        else:
            error_msg = str(payload)
            self.log(f"Erro na simulacao ATP: {error_msg}")
            self._update_simulation_results(
                f"Erro na simulacao ATP apos {elapsed:.1f}s:\n{error_msg}"
            )
            messagebox.showerror("Erro", f"Falha na simulacao ATP apos {elapsed:.1f}s:\n{error_msg}")

    def _set_atp_feedback_running(self):
        """Ativa indicadores visuais de simulacao ATP em andamento."""
        self._atp_running = True
        self._atp_started_at = time.time()
        self.atp_run_button.configure(state="disabled", text="Executando...")
        self.atp_progress.start()
        self.atp_run_status_var.set("Status: executando (0s)")
        self._tick_atp_running_status()

    def _set_atp_feedback_finished(self, success: bool) -> float:
        """Desativa indicadores visuais e retorna tempo decorrido."""
        elapsed = 0.0
        if self._atp_started_at is not None:
            elapsed = time.time() - self._atp_started_at

        self._atp_running = False
        self._atp_started_at = None
        self.atp_progress.stop()
        self.atp_progress.set(0)
        self.atp_run_button.configure(state="normal", text="Run Simulation")

        if success:
            self.atp_run_status_var.set(f"Status: concluido ({elapsed:.1f}s)")
        else:
            self.atp_run_status_var.set(f"Status: erro ({elapsed:.1f}s)")

        return elapsed

    def _tick_atp_running_status(self):
        """Atualiza o status de tempo da execucao ATP a cada segundo."""
        if not self._atp_running or self._atp_started_at is None:
            return

        elapsed = int(time.time() - self._atp_started_at)
        self.atp_run_status_var.set(f"Status: executando ({elapsed}s)")
        self.after(1000, self._tick_atp_running_status)

    def _update_simulation_results(self, text: str):
        """Atualiza area de resultados da simulacao"""
        self.simulation_results.configure(state="normal")
        self.simulation_results.delete("1.0", "end")
        self.simulation_results.insert("1.0", text)
        self.simulation_results.configure(state="disabled")
    
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
                base_outdir = Path(self.outdir_var.get())
                base_outdir.mkdir(parents=True, exist_ok=True)
                
                # Criar subpasta com data e hora
                timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
                outdir = base_outdir / timestamp
                outdir.mkdir(parents=True, exist_ok=True)
                
                self.log(f"Criada pasta de saída: {outdir.name}")
                
                # Coletar opções de visualização
                plot_options = {
                    'show_bars': self.plot_bars_var.get(),
                    'show_points': self.plot_points_var.get(),
                    'show_gaussian': self.plot_gaussian_var.get(),
                    'show_cumulative': self.plot_cumulative_var.get(),
                    'show_stats_box': self.plot_stats_box_var.get(),
                }
                
                excel_paths = []  # Armazenar paths dos Excel para comparativo
                total = len(selected_files)
                
                # NOTA: Processamento paralelo (parallel_process) desabilitado por segurança
                # Para evitar conflitos de matplotlib em threads múltiplas
                
                for idx, lis_path in enumerate(selected_files, start=1):
                    if self.cancel_event.is_set():
                        self.log("Processamento cancelado pelo usuário")
                        break
                    
                    # Atualizar progresso
                    progress = idx / total
                    self.progress_bar.set(progress)
                    self.status_var.set(f"Processando {lis_path.name}... ({idx}/{total})")
                    
                    self.log(f"Processando: {lis_path.name}")
                    
                    # Parse do .lis
                    df, stats_lines, summary = parse_lis_table(lis_path)
                    if df is None:
                        self.log(f"Tabela não encontrada em: {lis_path.name}")
                        continue
                    
                    # Nome base do arquivo (sem extensão)
                    base_name = lis_path.stem
                    
                    # Salvar Excel com nome do arquivo .lis (sempre sobrescreve)
                    excel_path = outdir / f"{base_name}.xlsx"
                    save_df_to_excel_only(df, excel_path)
                    excel_paths.append(excel_path)  # Adicionar à lista para comparativo
                    
                    # Estatísticas
                    try:
                        computed_stats = calcular_estatisticas_do_df(df)
                        escrever_estatisticas_excel(excel_path, computed_stats, summary_from_lis=summary)
                    except Exception as e:
                        error_msg = f"Erro ao calcular estatísticas: {e}"
                        if not self.hide_errors_var.get():
                            self.log(error_msg)
                            messagebox.showwarning("Aviso", error_msg)
                    
                    # Gráfico individual (se não for modo "só comparativo")
                    if not self.only_comparative_var.get():
                        graph_name = f"grafico_{base_name}.png"
                        self._criar_grafico_customizado(excel_path, outdir, graph_name, plot_options, mostrar=self.show_plots_var.get())
                    
                    # Séries temporais
                    try:
                        time_series_df = parse_lis_time_series(lis_path)
                        if time_series_df is not None:
                            save_time_series_to_excel(time_series_df, excel_path)
                            series_name = f"series_temporais_{base_name}.png"
                            criar_grafico_series_temporais(time_series_df, outdir / series_name, lis_name=lis_path.name)
                    except Exception as e:
                        error_msg = f"Erro ao processar séries temporais: {e}"
                        if not self.hide_errors_var.get():
                            self.log(error_msg)
                    
                    self.log(f"Concluído: {lis_path.name}")
                
                # Gerar gráfico comparativo se múltiplos arquivos
                if len(excel_paths) > 1 and not self.cancel_event.is_set():
                    self.log(f"Gerando gráfico comparativo de {len(excel_paths)} arquivos...")
                    self.status_var.set("Gerando gráfico comparativo...")
                    comp_name = f"comparativo_{timestamp}.png"
                    self._criar_grafico_comparativo_customizado(excel_paths, outdir, comp_name, plot_options, mostrar=self.show_plots_var.get())
                    self.log("Gráfico comparativo gerado com sucesso")
                
                # Salvar log se habilitado
                if self.save_logs_var.get():
                    log_content = self.log_textbox.get("1.0", "end")
                    log_file = outdir / f"log_{timestamp}.txt"
                    log_file.write_text(log_content, encoding='utf-8')
                    self.log(f"Log salvo: {log_file.name}")
                
                # Finalizar
                self.progress_bar.set(1.0)
                self.status_var.set(f"Processamento concluído! {len(selected_files)} arquivo(s)")
                self.log(f"Processamento finalizado com sucesso em: {outdir.name}")
                
                if self.open_output_var.get():
                    _open_in_file_manager(outdir)
                
                messagebox.showinfo("Sucesso", f"Processamento concluído!\n\n{len(selected_files)} arquivo(s) processado(s)\nResultados em:\n{outdir}")
                
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
    
    def _criar_grafico_customizado(self, excel_path: Path, outdir: Path, output_name: str, plot_options: dict, mostrar: bool = False):
        """Cria gráfico individual com opções customizadas de visualização"""
        try:
            # Importar aqui para evitar circular import
            from main import obter_xy_e_stats_de_excel
            
            res = obter_xy_e_stats_de_excel(excel_path)
            if res is None:
                self.log(f"Erro ao extrair dados de {excel_path.name}")
                return
            
            x, y, mu, sigma = res
            
            # Criar figura
            fig, ax = plt.subplots(figsize=(11, 7))
            
            # Barras de frequência
            if plot_options.get('show_bars', True):
                unique_x = np.unique(x)
                if unique_x.size > 1:
                    diffs = np.diff(unique_x)
                    diffs_pos = diffs[diffs > 0]
                    bin_width = float(np.median(diffs_pos)) if diffs_pos.size > 0 else (np.max(x) - np.min(x)) / max(1, len(unique_x))
                else:
                    bin_width = 0.1
                bar_width = bin_width * 0.9
                ax.bar(x, y, width=bar_width, alpha=0.35, label='Frequência (bins)', align='center', edgecolor='k', linewidth=0.3)
            
            # Pontos de dados
            if plot_options.get('show_points', True):
                ax.scatter(x, y, color='tab:blue', s=30, zorder=5, label='Pontos (x vs freq)')
            
            # Curva Gaussiana - estendida para mostrar curva completa
            if plot_options.get('show_gaussian', True) and sigma and np.isfinite(sigma) and sigma > 0:
                # Estender para ±3σ do centro (ou além dos dados se for maior)
                x_min_gauss = min(np.min(x), mu - 3*sigma)
                x_max_gauss = max(np.max(x), mu + 3*sigma)
                # Adicionar margem extra de 10%
                margin = (x_max_gauss - x_min_gauss) * 0.1
                x_smooth = np.linspace(x_min_gauss - margin, x_max_gauss + margin, 1000)
                pdf = np.exp(-0.5 * ((x_smooth - mu) / sigma)**2) / (sigma * np.sqrt(2 * np.pi))
                scale_factor = (np.max(y) / np.max(pdf)) if np.max(pdf) > 0 else 1.0
                y_smooth = pdf * scale_factor
                ax.plot(x_smooth, y_smooth, color='tab:orange', linewidth=2.2, label='Ajuste Gaussiano')
            
            ax.set_xlabel('Tensão (pu)')
            ax.set_ylabel('Frequência')
            ax.grid(alpha=0.25)
            ax.legend(loc='upper left')
            
            # Curva acumulada
            ax2 = None
            if plot_options.get('show_cumulative', True):
                ax2 = ax.twinx()
                total_weight = np.sum(y)
                cumsum = np.cumsum(y)
                cum_pct = (cumsum / total_weight) * 100.0
                ax2.plot(x, cum_pct, color='tab:green', marker='o', linestyle='--', label='Acumulado (%)', markersize=4)
                ax2.set_ylabel('Acumulado (%)')
                ax2.set_ylim(0, 100)
            
            # Caixa de estatísticas
            if plot_options.get('show_stats_box', True):
                try:
                    from main import calcular_estatisticas_do_df
                    import pandas as pd
                    df_excel = pd.read_excel(excel_path, sheet_name='Dados')
                    computed_stats = calcular_estatisticas_do_df(df_excel)
                    
                    def _safe_float(val):
                        if val is None:
                            return float('nan')
                        if hasattr(val, '__len__') and not isinstance(val, str):
                            try:
                                if len(val) > 0:
                                    val = val[0] if hasattr(val, '__getitem__') else float(val)
                                else:
                                    return float('nan')
                            except (TypeError, IndexError):
                                pass
                        if hasattr(val, 'item'):
                            try:
                                return val.item()
                            except (ValueError, TypeError):
                                pass
                        try:
                            return float(val)
                        except (TypeError, ValueError):
                            return float('nan')
                    
                    stats_text = (
                        f"μ = {_safe_float(mu):.6g}\\n"
                        f"σ = {_safe_float(sigma):.6g}\\n"
                        f"Mediana = {_safe_float(computed_stats.get('median', float('nan'))):.6g}\\n"
                        f"CV = {_safe_float(computed_stats.get('cv', float('nan'))):.6g}\\n"
                        f"R² = {_safe_float(computed_stats.get('r2', float('nan'))):.5g}"
                    )
                    
                    bbox_props = dict(boxstyle="round,pad=0.6", fc="white", ec="0.4", alpha=0.9)
                    ax.text(0.98, 0.95, stats_text, transform=ax.transAxes, fontsize=9,
                            verticalalignment='top', horizontalalignment='right', bbox=bbox_props)
                except Exception:
                    pass
            
            ax.set_title(f"Análise Detalhada — {excel_path.stem}")
            
            # Combinar legendas
            if ax2:
                lines, labels = ax.get_legend_handles_labels()
                lines2, labels2 = ax2.get_legend_handles_labels()
                ax2.legend(lines + lines2, labels + labels2, loc='lower right')
            
            # Salvar
            outdir = Path(outdir)
            outdir.mkdir(parents=True, exist_ok=True)
            out_png = outdir / output_name
            
            plt.tight_layout()
            plt.savefig(out_png, dpi=220, bbox_inches='tight')
            self.log(f"Gráfico salvo: {out_png.name}")
            
            if mostrar:
                plt.show()
            else:
                plt.close(fig)
                
        except Exception as e:
            self.log(f"Erro ao criar gráfico: {e}")
            import traceback
            traceback.print_exc()
    
    def _criar_grafico_comparativo_customizado(self, excel_paths: list, outdir: Path, output_name: str, plot_options: dict, mostrar: bool = False):
        """Cria gráfico comparativo com opções customizadas"""
        try:
            from main import obter_xy_e_stats_de_excel
            
            series_data = []
            labels = []
            
            for excel_path in excel_paths:
                res = obter_xy_e_stats_de_excel(excel_path)
                if res is None:
                    continue
                x, y, mu, sigma = res
                series_data.append((x, y, mu, sigma))
                labels.append(excel_path.stem)
            
            if not series_data:
                self.log("Sem dados para gráfico comparativo")
                return
            
            # Criar figura
            fig, ax = plt.subplots(figsize=(14, 8))
            colors = plt.cm.tab10(range(len(series_data)))
            
            # Preparar eixo secundário se curva acumulada estiver habilitada
            ax2 = None
            if plot_options.get('show_cumulative', True):
                ax2 = ax.twinx()
                ax2.set_ylabel('Acumulado (%)', fontsize=12, fontweight='bold')
                ax2.set_ylim(0, 100)
            
            for idx, ((x, y, mu, sigma), label) in enumerate(zip(series_data, labels)):
                color = colors[idx]
                
                # Barras (histograma) - com transparência para não sobrepor muito
                if plot_options.get('show_bars', True):
                    unique_x = np.unique(x)
                    if unique_x.size > 1:
                        diffs = np.diff(unique_x)
                        diffs_pos = diffs[diffs > 0]
                        bin_width = float(np.median(diffs_pos)) if diffs_pos.size > 0 else (np.max(x) - np.min(x)) / max(1, len(unique_x))
                    else:
                        bin_width = 0.1
                    bar_width = bin_width * 0.7  # Mais estreito no comparativo
                    # Offset para não sobrepor barras de diferentes séries
                    offset = (idx - len(series_data)/2) * bar_width * 0.3
                    ax.bar(x + offset, y, width=bar_width*0.8, alpha=0.25, color=color, 
                           label=f"{label} (barras)", align='center', edgecolor=color, linewidth=0.5)
                
                # Pontos
                if plot_options.get('show_points', True):
                    ax.scatter(x, y, s=25, alpha=0.6, color=color, label=f"{label} (pontos)", marker='o', edgecolors='white', linewidths=0.5)
                
                # Curva gaussiana - estendida para mostrar curva completa
                if plot_options.get('show_gaussian', True) and sigma and np.isfinite(sigma) and sigma > 0:
                    # Estender para ±3σ do centro (ou além dos dados se for maior)
                    x_min_gauss = min(np.min(x), mu - 3*sigma)
                    x_max_gauss = max(np.max(x), mu + 3*sigma)
                    # Adicionar margem extra de 10%
                    margin = (x_max_gauss - x_min_gauss) * 0.1
                    x_smooth = np.linspace(x_min_gauss - margin, x_max_gauss + margin, 1000)
                    pdf = np.exp(-0.5 * ((x_smooth - mu) / sigma)**2) / (sigma * np.sqrt(2 * np.pi))
                    scale_factor = (np.max(y) / np.max(pdf)) if np.max(pdf) > 0 else 1.0
                    y_smooth = pdf * scale_factor
                    ax.plot(x_smooth, y_smooth, linewidth=2.5, color=color, label=f"{label} (gaussiana)", linestyle='-', alpha=0.9)
                
                # Curva acumulada
                if plot_options.get('show_cumulative', True) and ax2:
                    total_weight = np.sum(y)
                    cumsum = np.cumsum(y)
                    cum_pct = (cumsum / total_weight) * 100.0
                    ax2.plot(x, cum_pct, color=color, marker='d', linestyle=':', 
                            label=f"{label} (acum.)", markersize=3, linewidth=1.5, alpha=0.7)
            
            ax.set_xlabel('Tensão (pu)', fontsize=12, fontweight='bold')
            ax.set_ylabel('Frequência', fontsize=12, fontweight='bold')
            ax.grid(alpha=0.3, linestyle='--')
            ax.set_title('Gráfico Comparativo - Distribuição e Ajuste Gaussiano', fontsize=14, fontweight='bold', pad=15)
            
            # Combinar legendas dos dois eixos
            if ax2:
                lines1, labels1 = ax.get_legend_handles_labels()
                lines2, labels2 = ax2.get_legend_handles_labels()
                ax.legend(lines1 + lines2, labels1 + labels2, ncol=2, fontsize=8, loc='best', framealpha=0.95)
            else:
                ax.legend(ncol=2, fontsize=9, loc='best', framealpha=0.9)
            
            # Caixa de estatísticas resumida
            if plot_options.get('show_stats_box', True):
                stats_text = f"Comparativo de {len(series_data)} arquivos:\n"
                for idx, ((x, y, mu, sigma), label) in enumerate(zip(series_data, labels)):
                    def _safe_float(val):
                        if val is None or (isinstance(val, float) and np.isnan(val)):
                            return float('nan')
                        if hasattr(val, 'item'):
                            try:
                                return val.item()
                            except:
                                pass
                        return float(val)
                    
                    mu_val = _safe_float(mu)
                    sigma_val = _safe_float(sigma)
                    stats_text += f"\n{label}:  μ={mu_val:.4f}  σ={sigma_val:.4f}"
                
                bbox_props = dict(boxstyle="round,pad=0.7", fc="white", ec="0.4", alpha=0.92)
                ax.text(0.02, 0.98, stats_text, transform=ax.transAxes, fontsize=8,
                       verticalalignment='top', horizontalalignment='left', 
                       bbox=bbox_props, family='monospace')
            
            # Salvar
            outdir = Path(outdir)
            outdir.mkdir(parents=True, exist_ok=True)
            out_png = outdir / output_name
            
            plt.tight_layout()
            plt.savefig(out_png, dpi=220, bbox_inches='tight')
            self.log(f"Gráfico comparativo salvo: {out_png.name}")
            
            if mostrar:
                plt.show()
            else:
                plt.close(fig)
                
        except Exception as e:
            self.log(f"Erro ao criar gráfico comparativo: {e}")
            import traceback
            traceback.print_exc()
    
    # TODO: implementar simulacao ATP com entrada .atp
    
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
