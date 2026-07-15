import threading
import traceback
import json
import re
import sys
import os
import subprocess
import time
import math
from datetime import datetime
from pathlib import Path

import customtkinter as ctk
from tkinter import filedialog, messagebox
import tkinter as tk

from .batch_runner import SweepParameterRef, generate_sweep_values, run_parameter_sweep
from .solver.atp_runner import ATPExecutionCancelled, run_atp_solver, get_missing_insert_dependencies
from .atp_parser import parse_atp_file_cached, get_editable_parameters, update_parameter
from .atp_writer import write_atp_file
from .failure_risk import FailureRiskConfig, FailureRiskResult, calculate_failure_risk
from .control_detector import (
    ControlDetector,
    FileControlInfo,
    analyze_workspace_files
)

# Configurações globais do CustomTkinter
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

PREFS_FILE = Path.home() / ".lis_analysis_gui.json"

parse_lis_table = None
parse_lis_once = None
save_df_to_excel_only = None
calcular_estatisticas_do_df = None
escrever_estatisticas_excel = None
criar_grafico_comparativo = None
parse_lis_time_series = None
save_time_series_to_excel = None
criar_grafico_series_temporais = None
obter_xy_e_stats_de_excel = None

_pipeline_imported = False
_pipeline_import_lock = threading.Lock()

plt = None
np = None
_plot_import_lock = threading.Lock()


def _ensure_pipeline_imports():
    global _pipeline_imported
    global parse_lis_table
    global parse_lis_once
    global save_df_to_excel_only
    global calcular_estatisticas_do_df
    global escrever_estatisticas_excel
    global criar_grafico_comparativo
    global parse_lis_time_series
    global save_time_series_to_excel
    global criar_grafico_series_temporais
    global obter_xy_e_stats_de_excel

    if _pipeline_imported:
        return

    with _pipeline_import_lock:
        if _pipeline_imported:
            return

        from .main import (
            parse_lis_table as _parse_lis_table,
            parse_lis_once as _parse_lis_once,
            save_df_to_excel_only as _save_df_to_excel_only,
            calcular_estatisticas_do_df as _calcular_estatisticas_do_df,
            escrever_estatisticas_excel as _escrever_estatisticas_excel,
            criar_grafico_comparativo as _criar_grafico_comparativo,
            parse_lis_time_series as _parse_lis_time_series,
            save_time_series_to_excel as _save_time_series_to_excel,
            criar_grafico_series_temporais as _criar_grafico_series_temporais,
            obter_xy_e_stats_de_excel as _obter_xy_e_stats_de_excel,
        )

        parse_lis_table = _parse_lis_table
        parse_lis_once = _parse_lis_once
        save_df_to_excel_only = _save_df_to_excel_only
        calcular_estatisticas_do_df = _calcular_estatisticas_do_df
        escrever_estatisticas_excel = _escrever_estatisticas_excel
        criar_grafico_comparativo = _criar_grafico_comparativo
        parse_lis_time_series = _parse_lis_time_series
        save_time_series_to_excel = _save_time_series_to_excel
        criar_grafico_series_temporais = _criar_grafico_series_temporais
        obter_xy_e_stats_de_excel = _obter_xy_e_stats_de_excel

        _pipeline_imported = True


def _ensure_plot_imports():
    global plt
    global np

    if plt is not None and np is not None:
        return

    with _plot_import_lock:
        if np is None:
            import numpy as _np
            np = _np

        if plt is None:
            import matplotlib

            if 'matplotlib.pyplot' not in sys.modules:
                try:
                    matplotlib.use('TkAgg')
                except Exception:
                    pass

            import matplotlib.pyplot as _plt
            plt = _plt


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


def _read_text_lines_preserve_newlines(path: Path) -> list[str]:
    """Lê arquivo texto preservando terminadores de linha originais."""
    with path.open("r", encoding="latin-1", errors="replace", newline="") as f:
        return f.read().splitlines(keepends=True)


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
        self._log_max_lines = 3000
        self._log_trim_every = 20
        self._log_entries_since_trim = 0

        # Risco de falha - valores de referencia do artigo SBSE 2025.
        self.enable_risk_var = tk.BooleanVar(value=False)
        self.risk_base_voltage_var = tk.StringVar(value="429")
        self.risk_height_var = tk.StringVar(value="17.98")
        self.risk_distance_var = tk.StringVar(value="3.50")
        self.risk_width_var = tk.StringVar(value="2.0")
        self.risk_subconductors_var = tk.StringVar(value="4")
        self.risk_insulation_distance_var = tk.StringVar(value="3.50")
        self.risk_parallel_gaps_var = tk.StringVar(value="804")
        
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
        self._atp_timeout_sec = 600
        self._atp_progress_value = 0.0
        self.atp_run_status_var = tk.StringVar(value="Status: aguardando execucao")
        self._atp_cancel_event = threading.Event()
        self._atp_active_mode = None
        self._atp_runtime_status_text = "Aguardando execucao"
        self._atp_sweep_completed_runs = set()
        self._atp_sweep_parallel_active = False

        # Simulacao ATP (.atp)
        self.atp_file_var = tk.StringVar(value='')
        self.atp_param_status_var = tk.StringVar(value="Nenhum parametro carregado")
        self.atp_sweep_parameter_var = tk.StringVar(value="")
        self.atp_sweep_start_var = tk.StringVar(value="")
        self.atp_sweep_stop_var = tk.StringVar(value="")
        self.atp_sweep_step_var = tk.StringVar(value="")
        self.atp_sweep_parse_lis_var = tk.BooleanVar(value=True)
        self.atp_sweep_stop_on_error_var = tk.BooleanVar(value=False)
        self.atp_batch_progress_var = tk.StringVar(value="0 / 0")
        self.atp_batch_elapsed_var = tk.StringVar(value="0s")
        self.atp_batch_status_var = tk.StringVar(value="Aguardando")
        self._atp_elements_cache = []
        self._atp_original_lines_cache = []
        self._atp_source_signature = None
        self._atp_lines_cache_signature = None
        self._atp_lines_cache = []
        self._atp_param_rows = []
        self.atp_param_filter_var = tk.StringVar(value="")
        self.parameter_overrides = {}
        self._atp_section_collapsed = {"branch": False, "switch": False, "source": False}
        self._atp_section_widgets = {}
        self._atp_section_order = ["branch", "switch", "source"]
        self._atp_filter_no_results_label = None
        self.atp_params_scroll_frame = None
        self._atp_expand_toggle_btn = None
        self._atp_params_expanded = False
        self._atp_params_collapsed_height = 260
        self._atp_params_expanded_height = 760
        self._sim_atp_card = None
        self._sim_params_card = None
        self._sim_action_card = None
        self._simulation_scroll_frame = None
        self._atp_sweep_parameter_options = {}
        self.atp_sweep_parameter_menu = None
        self.atp_sweep_run_button = None
        self.atp_cancel_button = None
        self._atp_scroll_isolation_bound = False
        self._atp_scroll_targets = ()
        self._atp_scroll_step_units = 10
        
        
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
                self.enable_risk_var.set(data.get('enable_failure_risk', False))
                self.risk_base_voltage_var.set(data.get('risk_base_voltage_kv', '429'))
                self.risk_height_var.set(data.get('risk_height_m', '17.98'))
                self.risk_distance_var.set(data.get('risk_distance_m', '3.50'))
                self.risk_width_var.set(data.get('risk_tower_width_m', '2.0'))
                self.risk_subconductors_var.set(data.get('risk_subconductors', '4'))
                self.risk_insulation_distance_var.set(data.get('risk_insulation_distance_m', '3.50'))
                self.risk_parallel_gaps_var.set(data.get('risk_parallel_gaps', '804'))
                
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
                'enable_failure_risk': self.enable_risk_var.get(),
                'risk_base_voltage_kv': self.risk_base_voltage_var.get(),
                'risk_height_m': self.risk_height_var.get(),
                'risk_distance_m': self.risk_distance_var.get(),
                'risk_tower_width_m': self.risk_width_var.get(),
                'risk_subconductors': self.risk_subconductors_var.get(),
                'risk_insulation_distance_m': self.risk_insulation_distance_var.get(),
                'risk_parallel_gaps': self.risk_parallel_gaps_var.get(),
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
        
        # Card: calculo do risco de falha conforme a subsecao 3.2 do artigo.
        risk_card = ctk.CTkFrame(scroll_frame, corner_radius=10)
        risk_card.pack(fill="x", pady=(0, 15))

        ctk.CTkLabel(
            risk_card,
            text="Risco de Falha - Hileman (1999)",
            font=ctk.CTkFont(size=16, weight="bold"),
        ).pack(anchor="w", padx=15, pady=(15, 8))
        ctk.CTkLabel(
            risk_card,
            text=(
                "Equações (1)-(6) da subseção 3.2. A média e o desvio-padrão "
                "do LIS são convertidos de p.u. para kV."
            ),
            font=ctk.CTkFont(size=11),
            justify="left",
        ).pack(anchor="w", padx=15, pady=(0, 8))
        ctk.CTkCheckBox(
            risk_card,
            text="Calcular e exibir o risco de falha nos gráficos",
            variable=self.enable_risk_var,
            command=self._toggle_risk_inputs,
        ).pack(anchor="w", padx=15, pady=(0, 10))

        self.risk_inputs_frame = ctk.CTkFrame(risk_card, fg_color="transparent")
        self.risk_inputs_frame.pack(fill="x", padx=15, pady=(0, 15))
        risk_grid = ctk.CTkFrame(self.risk_inputs_frame, fg_color="transparent")
        risk_grid.pack(fill="x")

        risk_fields = [
            ("Tensão-base [kV]", self.risk_base_voltage_var),
            ("Altura do condutor H [m]", self.risk_height_var),
            ("Distância condutor-estrutura D [m]", self.risk_distance_var),
            ("Largura da torre S [m]", self.risk_width_var),
            ("Subcondutores/fase N", self.risk_subconductors_var),
            ("Menor distância de isolamento d [m]", self.risk_insulation_distance_var),
            ("Gaps em paralelo n (201 x 4)", self.risk_parallel_gaps_var),
        ]
        self._risk_entries = []
        for idx, (label, variable) in enumerate(risk_fields):
            row = idx // 2
            pair_col = (idx % 2) * 2
            ctk.CTkLabel(risk_grid, text=f"{label}:").grid(
                row=row, column=pair_col, sticky="w", padx=(0, 8), pady=4
            )
            entry = ctk.CTkEntry(risk_grid, textvariable=variable, width=130)
            entry.grid(row=row, column=pair_col + 1, sticky="w", padx=(0, 24), pady=4)
            self._risk_entries.append(entry)

        self._toggle_risk_inputs()

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
    
    def _toggle_risk_inputs(self):
        frame = getattr(self, "risk_inputs_frame", None)
        if frame is None:
            return

        if self.enable_risk_var.get():
            if not frame.winfo_manager():
                frame.pack(fill="x", padx=15, pady=(0, 15))
        else:
            frame.pack_forget()

    @staticmethod
    def _parse_risk_number(raw_value: str, field_name: str) -> float:
        normalized = str(raw_value).strip().replace(",", ".")
        if not normalized:
            raise ValueError(f"{field_name} nao informado")
        try:
            return float(normalized)
        except ValueError as exc:
            raise ValueError(f"{field_name} invalido: {raw_value}") from exc

    def _collect_risk_config(self) -> FailureRiskConfig | None:
        if not self.enable_risk_var.get():
            return None

        config = FailureRiskConfig(
            base_voltage_kv=self._parse_risk_number(self.risk_base_voltage_var.get(), "Tensao-base"),
            conductor_height_m=self._parse_risk_number(self.risk_height_var.get(), "H"),
            conductor_structure_distance_m=self._parse_risk_number(self.risk_distance_var.get(), "D"),
            tower_width_m=self._parse_risk_number(self.risk_width_var.get(), "S"),
            subconductors_per_phase=self._parse_risk_number(self.risk_subconductors_var.get(), "N"),
            insulation_distance_m=self._parse_risk_number(
                self.risk_insulation_distance_var.get(), "d"
            ),
            parallel_gaps=self._parse_risk_number(self.risk_parallel_gaps_var.get(), "n"),
        )
        calculate_failure_risk(1.0, 0.1, config)
        return config

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
        self._simulation_scroll_frame = scroll_frame

        atp_card = ctk.CTkFrame(scroll_frame, corner_radius=10)
        atp_card.pack(fill="x", pady=(0, 15))
        self._sim_atp_card = atp_card

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

        params_card = ctk.CTkFrame(scroll_frame, corner_radius=10)
        params_card.pack(fill="x", pady=(0, 15))
        self._sim_params_card = params_card

        params_header = ctk.CTkFrame(params_card, fg_color="transparent")
        params_header.pack(fill="x", padx=15, pady=(15, 8))

        ctk.CTkLabel(
            params_header,
            text="Parametros editaveis do .atp",
            font=ctk.CTkFont(size=16, weight="bold")
        ).pack(side="left")

        params_actions = ctk.CTkFrame(params_card, fg_color="transparent")
        params_actions.pack(fill="x", padx=15, pady=(0, 6))

        ctk.CTkButton(
            params_actions,
            text="Detectar parametros",
            command=self._load_atp_parameters,
            width=180
        ).pack(side="left")

        ctk.CTkButton(
            params_actions,
            text="Baixar TXT",
            command=self._export_atp_parameters_txt,
            width=150
        ).pack(side="left", padx=(8, 0))

        ctk.CTkButton(
            params_actions,
            text="Resetar alteracoes",
            command=self._reset_atp_parameter_changes,
            width=170,
            fg_color="#757575",
            hover_color="#616161"
        ).pack(side="left", padx=(8, 0))

        self._atp_expand_toggle_btn = ctk.CTkButton(
            params_actions,
            text="Expandir",
            width=120,
            fg_color="#455A64",
            hover_color="#37474F",
            command=self._toggle_atp_params_expand,
        )
        self._atp_expand_toggle_btn.pack(side="left", padx=(8, 0))

        ctk.CTkLabel(
            params_actions,
            textvariable=self.atp_param_status_var,
            font=ctk.CTkFont(size=12, weight="bold")
        ).pack(side="left", padx=10)

        filter_row = ctk.CTkFrame(params_card, fg_color="transparent")
        filter_row.pack(fill="x", padx=15, pady=(0, 8))

        ctk.CTkLabel(filter_row, text="Buscar nó/parâmetro:").pack(side="left", padx=(0, 8))

        filter_entry = ctk.CTkEntry(
            filter_row,
            width=320,
            textvariable=self.atp_param_filter_var,
            placeholder_text="Ex.: X0001A, resistência, branch...",
        )
        filter_entry.pack(side="left", fill="x", expand=True)

        ctk.CTkButton(
            filter_row,
            text="Limpar",
            width=90,
            command=lambda: self.atp_param_filter_var.set(""),
        ).pack(side="left", padx=(8, 0))

        self.atp_param_filter_var.trace_add("write", lambda *_args: self._apply_atp_parameter_filter())

        self.atp_params_scroll_frame = ctk.CTkScrollableFrame(params_card, width=1060, height=260)
        self.atp_params_scroll_frame.pack(fill="x", padx=15, pady=(0, 15))

        self._setup_atp_params_scroll_isolation()

        sweep_card = ctk.CTkFrame(scroll_frame, corner_radius=10)
        sweep_card.pack(fill="x", pady=(0, 15))

        ctk.CTkLabel(
            sweep_card,
            text="Execucao em lote (parameter sweep)",
            font=ctk.CTkFont(size=16, weight="bold")
        ).pack(anchor="w", padx=15, pady=(15, 10))

        ctk.CTkLabel(
            sweep_card,
            text="Escolha um parametro editavel e defina inicio, fim e passo para executar varias simulacoes automaticamente.",
            justify="left",
        ).pack(anchor="w", padx=15, pady=(0, 10))

        sweep_grid = ctk.CTkFrame(sweep_card, fg_color="transparent")
        sweep_grid.pack(fill="x", padx=15, pady=(0, 10))
        sweep_grid.grid_columnconfigure(1, weight=1)
        sweep_grid.grid_columnconfigure(3, weight=1)

        ctk.CTkLabel(sweep_grid, text="Parametro do sweep:").grid(row=0, column=0, sticky="w", padx=(0, 10), pady=6)
        self.atp_sweep_parameter_menu = ctk.CTkOptionMenu(
            sweep_grid,
            values=["Nenhum parametro carregado"],
            variable=self.atp_sweep_parameter_var,
            command=self._on_atp_sweep_parameter_selected,
            width=520,
        )
        self.atp_sweep_parameter_menu.grid(row=0, column=1, columnspan=3, sticky="ew", pady=6)

        ctk.CTkLabel(sweep_grid, text="Inicio:").grid(row=1, column=0, sticky="w", padx=(0, 10), pady=6)
        ctk.CTkEntry(sweep_grid, textvariable=self.atp_sweep_start_var, width=160).grid(row=1, column=1, sticky="ew", pady=6)

        ctk.CTkLabel(sweep_grid, text="Fim:").grid(row=1, column=2, sticky="w", padx=(18, 10), pady=6)
        ctk.CTkEntry(sweep_grid, textvariable=self.atp_sweep_stop_var, width=160).grid(row=1, column=3, sticky="ew", pady=6)

        ctk.CTkLabel(sweep_grid, text="Passo:").grid(row=2, column=0, sticky="w", padx=(0, 10), pady=6)
        ctk.CTkEntry(sweep_grid, textvariable=self.atp_sweep_step_var, width=160).grid(row=2, column=1, sticky="ew", pady=6)

        sweep_options = ctk.CTkFrame(sweep_card, fg_color="transparent")
        sweep_options.pack(fill="x", padx=15, pady=(0, 15))

        ctk.CTkCheckBox(
            sweep_options,
            text="Pos-processar .lis de cada execucao",
            variable=self.atp_sweep_parse_lis_var,
        ).pack(side="left")

        ctk.CTkCheckBox(
            sweep_options,
            text="Parar no primeiro erro",
            variable=self.atp_sweep_stop_on_error_var,
        ).pack(side="left", padx=(18, 0))

        action_card = ctk.CTkFrame(scroll_frame, corner_radius=10)
        action_card.pack(fill="x", pady=(0, 15))
        self._sim_action_card = action_card

        ctk.CTkLabel(
            action_card,
            text="Executar Simulacao",
            font=ctk.CTkFont(size=16, weight="bold")
        ).pack(anchor="w", padx=15, pady=(15, 10))

        buttons_frame = ctk.CTkFrame(action_card, fg_color="transparent")
        buttons_frame.pack(padx=15, pady=(0, 15))

        self.atp_run_button = ctk.CTkButton(
            buttons_frame,
            text="Executar Simulação",
            command=self._run_atp_simulation,
            width=200,
            height=40,
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color="#2196F3",
            hover_color="#1976D2"
        )
        self.atp_run_button.pack(side="left", padx=5)

        self.atp_sweep_run_button = ctk.CTkButton(
            buttons_frame,
            text="Executar em Lote",
            command=self._run_atp_parameter_sweep,
            width=220,
            height=40,
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color="#2E7D32",
            hover_color="#1B5E20"
        )
        self.atp_sweep_run_button.pack(side="left", padx=5)

        self.atp_cancel_button = ctk.CTkButton(
            buttons_frame,
            text="Cancelar",
            command=self._cancel_atp_execution,
            width=140,
            height=40,
            state="disabled",
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color="#C62828",
            hover_color="#8E0000"
        )
        self.atp_cancel_button.pack(side="left", padx=5)

        metrics_frame = ctk.CTkFrame(action_card, fg_color="transparent")
        metrics_frame.pack(fill="x", padx=15, pady=(0, 8))

        ctk.CTkLabel(metrics_frame, text="Progresso:", font=ctk.CTkFont(size=12, weight="bold")).pack(side="left")
        ctk.CTkLabel(metrics_frame, textvariable=self.atp_batch_progress_var).pack(side="left", padx=(6, 18))
        ctk.CTkLabel(metrics_frame, text="Tempo:", font=ctk.CTkFont(size=12, weight="bold")).pack(side="left")
        ctk.CTkLabel(metrics_frame, textvariable=self.atp_batch_elapsed_var).pack(side="left", padx=(6, 18))
        ctk.CTkLabel(metrics_frame, text="Status atual:", font=ctk.CTkFont(size=12, weight="bold")).pack(side="left")
        ctk.CTkLabel(metrics_frame, textvariable=self.atp_batch_status_var).pack(side="left", padx=(6, 0))

        self.atp_progress = ctk.CTkProgressBar(action_card)
        self.atp_progress.pack(fill="x", padx=15, pady=(0, 8))
        self.atp_progress.set(0)

        ctk.CTkLabel(
            action_card,
            textvariable=self.atp_run_status_var,
            font=ctk.CTkFont(size=12)
        ).pack(anchor="w", padx=15, pady=(0, 12))

        self.simulation_results = ctk.CTkTextbox(scroll_frame, width=1100, height=160)
        self.simulation_results.pack(fill="both", expand=True, pady=(0, 10))
        self.simulation_results.insert("1.0", "Aguardando execucao da simulacao ATP...\n")
        self.simulation_results.configure(state="disabled")
        self._apply_atp_params_expand_state()
    
    @staticmethod
    def _failure_risk_record(label: str, result: FailureRiskResult) -> dict:
        return {
            "label": str(label),
            "risk": float(result.risk),
            "corrected_cfo_kv": float(result.corrected_cfo_kv),
            "z_score": float(result.z_score),
            "mean_overvoltage_kv": float(result.mean_overvoltage_kv),
            "switching_std_kv": float(result.switching_std_kv),
        }

    @staticmethod
    def _write_failure_risk_report(output_dir: Path, context: str, records: list[dict]) -> Path | None:
        """Grava os resultados de risco em um arquivo proprio na pasta de saida."""
        if not records:
            return None

        output_dir = Path(output_dir)
        output_dir.mkdir(parents=True, exist_ok=True)
        report_path = output_dir / "risco_de_falha.txt"
        lines = [
            "RELATORIO DE RISCO DE FALHA",
            f"Contexto: {context}",
            f"Gerado em: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
            f"Resultados: {len(records)}",
            "",
        ]

        for index, record in enumerate(records, start=1):
            risk = float(record["risk"])
            lines.extend(
                [
                    f"[{index}] {record['label']}",
                    f"Probabilidade de falha: {risk * 100.0:.8g}%",
                    f"R: {risk:.10e}",
                    f"CFO corrigido: {float(record['corrected_cfo_kv']):.4f} kV",
                    f"Indice normalizado Z: {float(record['z_score']):.6f}",
                    f"Sobretensao media: {float(record['mean_overvoltage_kv']):.4f} kV",
                    f"Desvio-padrao da sobretensao: {float(record['switching_std_kv']):.4f} kV",
                    "",
                ]
            )

        report_path.write_text("\n".join(lines), encoding="utf-8")
        return report_path

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
        # Rodape removido da interface. Estado/acoes permanecem via variaveis e logs.
        self.status_label = None
        self.progress_bar = None
        self.cancel_btn = None

    def _set_main_progress(self, value: float):
        """Atualiza barra principal somente se o widget existir."""
        if self.progress_bar is None:
            return
        try:
            self.progress_bar.set(float(value))
        except Exception:
            pass

    def _show_main_cancel(self):
        """Exibe botao de cancelar somente se o widget existir."""
        if self.cancel_btn is None:
            return
        try:
            self.cancel_btn.pack(side="right", padx=15)
        except Exception:
            pass

    def _hide_main_cancel(self):
        """Oculta botao de cancelar somente se o widget existir."""
        if self.cancel_btn is None:
            return
        try:
            self.cancel_btn.pack_forget()
        except Exception:
            pass
    
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
            self._clear_atp_parameter_editor()
            self._load_atp_parameters(show_dialog_errors=False)
            self._save_prefs()

    def _get_atp_file_signature(self, atp_path: Path) -> tuple[str, int, int]:
        """Assinatura do arquivo ATP para invalidar cache quando houver alteração."""
        resolved = str(atp_path.resolve())
        st = atp_path.stat()
        return (resolved, int(st.st_mtime_ns), int(st.st_size))

    def _get_cached_atp_original_lines(self, atp_path: Path, signature: tuple[str, int, int]) -> list[str]:
        """Retorna linhas originais do ATP com cache simples por assinatura."""
        if self._atp_lines_cache_signature == signature and self._atp_lines_cache:
            return list(self._atp_lines_cache)

        lines = _read_text_lines_preserve_newlines(atp_path)
        self._atp_lines_cache_signature = signature
        self._atp_lines_cache = list(lines)
        return lines

    def _clear_atp_parameter_editor(self):
        """Limpa a lista visual e caches de parametros ATP detectados."""
        for row in self._atp_param_rows:
            job = row.get("_repeat_job")
            if job is None:
                continue
            try:
                self.after_cancel(job)
            except Exception:
                pass

        self._atp_elements_cache = []
        self._atp_original_lines_cache = []
        self._atp_source_signature = None
        self._atp_param_rows = []
        self.parameter_overrides = {}
        self._atp_section_widgets = {}
        self._atp_filter_no_results_label = None
        self.atp_param_status_var.set("Nenhum parametro carregado")
        self._refresh_atp_sweep_parameter_menu()

        if self.atp_params_scroll_frame is None:
            return

        for widget in self.atp_params_scroll_frame.winfo_children():
            widget.destroy()

    def _toggle_atp_params_expand(self):
        """Alterna modo de foco do container de parâmetros ATP."""
        self._atp_params_expanded = not bool(self._atp_params_expanded)
        self._apply_atp_params_expand_state()

    def _apply_atp_params_expand_state(self):
        """Aplica visualmente o estado expandido/reduzido do container de parâmetros ATP."""
        if self.atp_params_scroll_frame is not None:
            if self._atp_params_expanded:
                target_height = max(self._atp_params_collapsed_height, int(self.winfo_height() * 0.72))
            else:
                target_height = self._atp_params_collapsed_height
            try:
                self.atp_params_scroll_frame.configure(height=target_height)
            except Exception:
                pass

        if self._atp_expand_toggle_btn is not None:
            self._atp_expand_toggle_btn.configure(
                text="Reduzir" if self._atp_params_expanded else "Expandir"
            )

    def _setup_atp_params_scroll_isolation(self):
        """Isola o scroll da lista de parametros ATP para não mover o scroll da aba."""
        if self.atp_params_scroll_frame is not None:
            targets = [self.atp_params_scroll_frame]

            parent_canvas = getattr(self.atp_params_scroll_frame, "_parent_canvas", None)
            if parent_canvas is not None:
                targets.append(parent_canvas)

            scrollbar = getattr(self.atp_params_scroll_frame, "_scrollbar", None)
            if scrollbar is not None:
                targets.append(scrollbar)

            self._atp_scroll_targets = tuple(targets)

        if self._atp_scroll_isolation_bound:
            return

        self.bind("<MouseWheel>", self._on_root_mousewheel_for_atp_params, add="+")
        self.bind("<Button-4>", self._on_root_mousewheel_for_atp_params, add="+")
        self.bind("<Button-5>", self._on_root_mousewheel_for_atp_params, add="+")
        self._atp_scroll_isolation_bound = True

    def _widget_belongs_to_atp_params(self, widget) -> bool:
        if widget is None:
            return False

        targets = self._atp_scroll_targets
        if not targets:
            return False

        current = widget
        while current is not None:
            for target in targets:
                if current == target:
                    return True
            current = getattr(current, "master", None)

        return False

    def _on_root_mousewheel_for_atp_params(self, event):
        """Redireciona roda do mouse para o container ATP e bloqueia propagacao para o pai."""
        if self.atp_params_scroll_frame is None:
            return None

        if self.tabview.get() != "Simulacao ATP":
            return None

        widget_under_pointer = getattr(event, "widget", None)
        if widget_under_pointer is None:
            widget_under_pointer = self.winfo_containing(event.x_root, event.y_root)

        if not self._widget_belongs_to_atp_params(widget_under_pointer):
            return None

        canvas = getattr(self.atp_params_scroll_frame, "_parent_canvas", None)
        if canvas is None:
            return "break"

        event_num = getattr(event, "num", None)
        if event_num == 4:
            units = -self._atp_scroll_step_units
        elif event_num == 5:
            units = self._atp_scroll_step_units
        else:
            delta = int(getattr(event, "delta", 0))
            if delta == 0:
                return "break"

            if sys.platform == "darwin":
                units = -1 if delta > 0 else 1
            else:
                units = -int(delta / 120)
                if units == 0:
                    units = -1 if delta > 0 else 1

            units *= self._atp_scroll_step_units

        canvas.yview_scroll(units, "units")
        return "break"

    def _apply_atp_parameter_filter(self):
        if self.atp_params_scroll_frame is None:
            return

        query = self.atp_param_filter_var.get().strip().lower()
        total_visible_cards = 0

        for section_key in self._atp_section_order:
            section_data = self._atp_section_widgets.get(section_key)
            if not section_data:
                continue

            visible_cards = 0
            for card_data in section_data.get("cards", []):
                widget = card_data["widget"]
                should_show = (not query) or (query in card_data.get("search_text", ""))

                if should_show:
                    visible_cards += 1
                    if not card_data.get("visible", True):
                        widget.pack(fill="x", padx=10, pady=6)
                        card_data["visible"] = True
                else:
                    if card_data.get("visible", True):
                        widget.pack_forget()
                        card_data["visible"] = False

            section_frame = section_data.get("section")
            if section_frame is None:
                continue

            if visible_cards > 0:
                total_visible_cards += visible_cards
                if not section_data.get("section_visible", True):
                    section_frame.pack(fill="x", padx=4, pady=(6, 8))
                    section_data["section_visible"] = True
                self._apply_atp_section_visibility(section_key)
            else:
                if section_data.get("section_visible", True):
                    section_frame.pack_forget()
                    section_data["section_visible"] = False

        if query and total_visible_cards == 0:
            if self._atp_filter_no_results_label is None:
                self._atp_filter_no_results_label = ctk.CTkLabel(
                    self.atp_params_scroll_frame,
                    text="Nenhum parâmetro corresponde ao filtro informado.",
                    justify="left",
                )
            self._atp_filter_no_results_label.pack(anchor="w", padx=8, pady=8)
        else:
            if self._atp_filter_no_results_label is not None:
                self._atp_filter_no_results_label.pack_forget()

    def _apply_atp_section_visibility(self, section_key: str):
        section_data = self._atp_section_widgets.get(section_key)
        if not section_data:
            return

        body = section_data.get("body")
        toggle = section_data.get("toggle")
        if body is None or toggle is None:
            return

        collapsed = bool(self._atp_section_collapsed.get(section_key, False))
        if collapsed:
            body.pack_forget()
            toggle.configure(text="▸")
        else:
            body.pack(fill="x", padx=10, pady=(0, 8))
            toggle.configure(text="▾")

    def _toggle_atp_section(self, section_key: str):
        current = bool(self._atp_section_collapsed.get(section_key, False))
        self._atp_section_collapsed[section_key] = not current
        self._apply_atp_section_visibility(section_key)

    def _friendly_atp_param_label(self, param_name: str) -> str:
        labels = {
            "resistance": "Resistência (Ω)",
            "inductance": "Indutância (H)",
            "capacitance": "Capacitância (F)",
            "t_close": "Tempo de fechamento (s)",
            "delay": "Atraso (s)",
            "amplitude": "Amplitude",
            "frequency": "Frequência (Hz)",
            "phase": "Fase (graus)",
        }
        return labels.get(param_name, param_name)

    def _extract_element_identifier(self, element: dict) -> str:
        """Extrai identificação amigável do elemento com base nos nós da linha RAW."""
        raw_line = str(element.get("raw_line", ""))
        first_token = raw_line.strip().split()[0] if raw_line.strip() else ""
        node_matches = re.findall(r"[A-Za-z]\d{4}[A-Za-z]", first_token)
        element_type = str(element.get("type", "")).lower()

        if element_type in {"branch", "switch"}:
            if len(node_matches) >= 2:
                return f"{node_matches[0]} → {node_matches[1]}"
            if len(first_token) >= 12:
                return f"{first_token[:6]} → {first_token[6:12]}"
        elif element_type == "source":
            if node_matches:
                return node_matches[0]

        return first_token if first_token else "Nao identificado"

    def _validate_numeric_input(self, new_value: str) -> bool:
        """Validação de entrada para bloquear caracteres inválidos em campos numéricos ATP."""
        if new_value == "":
            return True

        pattern = r"^-?\d*\.?\d*$"
        return re.match(pattern, new_value) is not None

    def _is_non_negative_parameter(self, param_name: str) -> bool:
        return param_name in {"resistance", "inductance", "capacitance", "amplitude", "frequency", "t_close"}

    def _get_param_step(self, param_name: str, original_value: float) -> float:
        if param_name == "resistance":
            return 1.0
        if param_name == "inductance":
            return 10.0 if abs(float(original_value)) >= 100.0 else 1.0
        if param_name == "amplitude":
            return 1.0
        if param_name == "frequency":
            return 5.0
        if param_name in {"t_close", "delay"}:
            return 0.1
        if param_name == "phase":
            return 1.0

        abs_value = abs(float(original_value))
        if abs_value >= 1000:
            return 10.0
        if abs_value >= 100:
            return 1.0
        if abs_value >= 10:
            return 0.1
        if abs_value >= 1:
            return 0.01
        if abs_value >= 0.1:
            return 0.001
        return 0.0001

    def _format_param_value(self, value: float, keep_trailing_dot: bool = False) -> str:
        value_f = float(value)
        if value_f.is_integer():
            return f"{value_f:.1f}"
        return f"{value_f:.10g}"

    def _refresh_atp_param_status(self):
        total = len(self._atp_param_rows)
        changed = len(self.parameter_overrides)
        invalid = sum(1 for row in self._atp_param_rows if bool(row.get("invalid", False)))

        if total == 0:
            self.atp_param_status_var.set("Nenhum parametro carregado")
            return

        status = f"Editaveis: {total} | Alterados: {changed}"
        if invalid:
            status += f" | Invalidos: {invalid}"
        self.atp_param_status_var.set(status)

    def _set_atp_row_visual_state(self, row: dict, state: str):
        entry = row["entry"]
        status_label = row.get("status_label")
        if state == "invalid":
            entry.configure(
                fg_color=row["default_fg_color"],
                border_color="#DC2626",
                border_width=2,
                text_color=row["default_text_color"],
            )
            if status_label is not None:
                status_label.configure(text="Inválido", text_color=("#B91C1C", "#FCA5A5"))
        elif state == "changed":
            entry.configure(
                fg_color=row["default_fg_color"],
                border_color="#D97706",
                border_width=2,
                text_color=row["default_text_color"],
            )
            if status_label is not None:
                status_label.configure(text="Alterado", text_color=("#B45309", "#FCD34D"))
        else:
            entry.configure(
                fg_color=row["default_fg_color"],
                border_color=row["default_border_color"],
                border_width=row["default_border_width"],
                text_color=row["default_text_color"],
            )
            if status_label is not None:
                status_label.configure(text="Original", text_color=("#4B5563", "#94A3B8"))

    def _update_atp_parameter_row(self, row: dict, source: str = "entry"):
        key = (int(row["line_index"]), str(row["parameter"]))
        raw = row["var"].get().strip()

        if raw == "":
            row["invalid"] = True
            self.parameter_overrides.pop(key, None)
            self._set_atp_row_visual_state(row, "invalid")
            self._refresh_atp_param_status()
            return

        normalized = raw.replace("D", "E").replace("d", "e")
        try:
            parsed_value = float(normalized)
        except ValueError:
            row["invalid"] = True
            self.parameter_overrides.pop(key, None)
            self._set_atp_row_visual_state(row, "invalid")
            self._refresh_atp_param_status()
            return

        row["invalid"] = False

        original_value = float(row["original_value"])
        if abs(parsed_value - original_value) > 1e-15:
            self.parameter_overrides[key] = float(parsed_value)
            self._set_atp_row_visual_state(row, "changed")
        else:
            self.parameter_overrides.pop(key, None)
            self._set_atp_row_visual_state(row, "normal")

        self._refresh_atp_param_status()

    def _on_atp_entry_changed(self, row: dict):
        if row.get("_updating", False):
            return
        self._update_atp_parameter_row(row, source="entry")

    def _adjust_atp_parameter(self, row: dict, direction: int):
        if row.get("_updating", False):
            return

        raw = row["var"].get().strip()
        normalized = raw.replace("D", "E").replace("d", "e")
        try:
            current = float(normalized)
        except ValueError:
            current = float(row["original_value"])

        new_value = current + (float(direction) * float(row["step"]))
        if self._is_non_negative_parameter(str(row.get("parameter", ""))):
            new_value = max(0.0, new_value)

        row["_updating"] = True
        row["var"].set(self._format_param_value(float(new_value), keep_trailing_dot=row["keep_trailing_dot"]))
        row["_updating"] = False
        self._update_atp_parameter_row(row, source="buttons")

    def _reset_single_atp_parameter(self, row: dict):
        """Reseta um único parâmetro para valor original e remove override correspondente."""
        row["_updating"] = True
        original = float(row["original_value"])
        row["var"].set(self._format_param_value(original, keep_trailing_dot=row["keep_trailing_dot"]))
        row["_updating"] = False
        row["invalid"] = False

        key = (int(row["line_index"]), str(row["parameter"]))
        self.parameter_overrides.pop(key, None)
        self._set_atp_row_visual_state(row, "normal")
        self._refresh_atp_param_status()

    def _repeat_adjust_tick(self, row: dict):
        direction = int(row.get("_repeat_direction", 0))
        if direction == 0:
            row["_repeat_job"] = None
            return

        self._adjust_atp_parameter(row, direction)
        row["_repeat_job"] = self.after(90, lambda r=row: self._repeat_adjust_tick(r))

    def _start_adjust_repeat(self, row: dict, direction: int):
        self._stop_adjust_repeat(row)
        row["_repeat_direction"] = int(direction)

        # Primeiro passo imediato, depois repete enquanto pressionado.
        self._adjust_atp_parameter(row, int(direction))
        row["_repeat_job"] = self.after(350, lambda r=row: self._repeat_adjust_tick(r))

    def _stop_adjust_repeat(self, row: dict):
        row["_repeat_direction"] = 0
        job = row.get("_repeat_job")
        if job is None:
            return
        try:
            self.after_cancel(job)
        except Exception:
            pass
        row["_repeat_job"] = None

    def _apply_atp_parameter_changes(self):
        """Valida campos e confirma alterações no estado da interface."""
        try:
            overrides = self._collect_atp_parameter_overrides()
        except Exception as e:
            self._show_error("Erro", "Existem valores invalidos no editor ATP.", details=[("Detalhes", str(e))])
            return

        if not overrides:
            self._show_info("Parametros ATP", "Nenhuma alteracao pendente para aplicar.")
            return

        self.log(f"[ATP] Alteracoes confirmadas na interface: {len(overrides)} parametro(s)")
        self._show_success(
            "Alteracoes aplicadas",
            "As alteracoes foram registradas e serao usadas no proximo Run Simulation.",
            details=[("Parametros alterados", str(len(overrides)))],
        )

    def _reset_atp_parameter_changes(self):
        """Restaura todos os campos para os valores originais detectados."""
        if not self._atp_param_rows:
            self._show_info("Parametros ATP", "Nenhum parametro carregado para resetar.")
            return

        for row in self._atp_param_rows:
            row["_updating"] = True
            original = float(row["original_value"])
            row["var"].set(self._format_param_value(original, keep_trailing_dot=row["keep_trailing_dot"]))
            row["_updating"] = False
            row["invalid"] = False
            self._set_atp_row_visual_state(row, "normal")

        self.parameter_overrides = {}
        self._refresh_atp_param_status()
        self.log("[ATP] Alteracoes de parametros resetadas para valores originais")

    def _build_atp_overrides_preview_table(self, overrides: list[dict]) -> str:
        header = f"{'Elemento':<30} {'Parametro':<22} {'Valor'}"
        lines = [header, "-" * len(header)]

        max_rows = 80
        for item in overrides[:max_rows]:
            name = str(item.get("name", ""))[:30]
            parameter = self._friendly_atp_param_label(str(item.get("parameter", "")))[:22]
            old_value = self._format_param_value(float(item.get("old_value", 0.0)))
            new_value = self._format_param_value(float(item.get("new_value", 0.0)))
            lines.append(f"{name:<30} {parameter:<22} {old_value} -> {new_value}")

        if len(overrides) > max_rows:
            lines.append(f"... +{len(overrides) - max_rows} alteracao(oes) adicionais")

        return "\n".join(lines)

    def _confirm_atp_overrides_preview(self, overrides: list[dict]) -> bool:
        """Mostra preview de alterações e pede confirmação para executar a simulação."""
        if not overrides:
            return True

        dialog = ctk.CTkToplevel(self)
        dialog.title("Preview de alteracoes ATP")
        dialog.transient(self)
        dialog.resizable(False, False)
        dialog.grab_set()

        result = {"confirmed": False}

        container = ctk.CTkFrame(dialog, corner_radius=12)
        container.pack(fill="both", expand=True, padx=14, pady=14)

        ctk.CTkLabel(
            container,
            text="Preview de alteracoes antes do Run Simulation",
            font=ctk.CTkFont(size=16, weight="bold"),
            anchor="w",
        ).pack(fill="x", padx=12, pady=(12, 6))

        ctk.CTkLabel(
            container,
            text=f"{len(overrides)} parametro(s) alterado(s). Confirma execucao com estas alteracoes?",
            anchor="w",
            justify="left",
        ).pack(fill="x", padx=12, pady=(0, 8))

        preview_text = self._build_atp_overrides_preview_table(overrides)
        preview_box = ctk.CTkTextbox(container, width=760, height=260, wrap="none")
        preview_box.pack(fill="both", expand=True, padx=12, pady=(0, 10))
        preview_box.insert("1.0", preview_text)
        preview_box.configure(state="disabled")

        buttons = ctk.CTkFrame(container, fg_color="transparent")
        buttons.pack(fill="x", padx=12, pady=(0, 12))

        def _cancel(_event=None):
            try:
                dialog.grab_release()
            except Exception:
                pass
            dialog.destroy()

        def _confirm(_event=None):
            result["confirmed"] = True
            _cancel()

        ctk.CTkButton(buttons, text="Cancelar", width=130, fg_color="#757575", hover_color="#616161", command=_cancel).pack(side="right")
        ctk.CTkButton(buttons, text="Continuar", width=130, fg_color="#2E7D32", hover_color="#1B5E20", command=_confirm).pack(side="right", padx=(0, 8))

        dialog.bind("<Escape>", _cancel)
        dialog.bind("<Return>", _confirm)

        dialog.update_idletasks()
        width = max(760, min(dialog.winfo_reqwidth(), 980))
        height = max(430, min(dialog.winfo_reqheight(), 700))

        parent_x = self.winfo_rootx()
        parent_y = self.winfo_rooty()
        parent_w = self.winfo_width()
        parent_h = self.winfo_height()
        pos_x = parent_x + (parent_w - width) // 2
        pos_y = parent_y + (parent_h - height) // 2
        dialog.geometry(f"{width}x{height}+{pos_x}+{pos_y}")
        dialog.lift()
        dialog.focus_force()
        self.wait_window(dialog)

        return bool(result["confirmed"])

    def _load_atp_parameters(self, show_dialog_errors: bool = True):
        """Lê o .atp atual e monta editor de parametros detectados automaticamente."""
        atp_file = self.atp_file_var.get().strip()
        atp_path = Path(atp_file)
        if not atp_file or not atp_path.exists():
            self._clear_atp_parameter_editor()
            if show_dialog_errors:
                self._show_error("Erro", "Arquivo .atp nao encontrado para detectar parametros.")
            return

        try:
            signature = self._get_atp_file_signature(atp_path)
            elements = parse_atp_file_cached(atp_path)
            editable = get_editable_parameters(elements)
            original_lines = self._get_cached_atp_original_lines(atp_path, signature)
        except Exception as e:
            self._clear_atp_parameter_editor()
            if show_dialog_errors:
                self._show_error("Erro", "Falha ao analisar arquivo .atp.", details=[("Detalhes", str(e))])
            return

        self._clear_atp_parameter_editor()
        self._atp_source_signature = signature
        self._atp_elements_cache = elements
        self._atp_original_lines_cache = original_lines

        if not editable:
            self.atp_param_status_var.set("Nenhum componente editavel detectado")
            self._refresh_atp_sweep_parameter_menu()
            if self.atp_params_scroll_frame is not None:
                ctk.CTkLabel(
                    self.atp_params_scroll_frame,
                    text="Nao foram encontrados componentes R/L/C/V/I editaveis neste arquivo.",
                    justify="left"
                ).pack(anchor="w", padx=6, pady=6)
            self.log("[ATP] Nenhum parametro editavel detectado no .atp selecionado")
            return

        group_titles = {
            "branch": "BRANCHES",
            "switch": "SWITCHES",
            "source": "SOURCES",
        }
        grouped_elements = {"branch": [], "switch": [], "source": []}

        for element in self._atp_elements_cache:
            etype = str(element.get("type", "")).lower()
            if etype not in grouped_elements:
                continue

            params = element.get("parameters", {})
            if not isinstance(params, dict):
                continue

            editable_params = []
            for param_name, meta in params.items():
                if not isinstance(meta, dict):
                    continue
                if not bool(meta.get("editable", True)):
                    continue
                editable_params.append((str(param_name), meta))

            if editable_params:
                grouped_elements[etype].append((element, editable_params))

        vcmd_numeric = (self.register(self._validate_numeric_input), "%P")
        total_cards = 0
        for etype in ("branch", "switch", "source"):
            items = grouped_elements.get(etype, [])
            if not items:
                continue

            section = ctk.CTkFrame(self.atp_params_scroll_frame, corner_radius=10)
            section.pack(fill="x", padx=4, pady=(6, 8))

            header = ctk.CTkFrame(section, fg_color="transparent")
            header.pack(fill="x", padx=12, pady=(10, 4))

            ctk.CTkLabel(
                header,
                text=group_titles[etype],
                font=ctk.CTkFont(size=16, weight="bold"),
                anchor="w",
            ).pack(side="left", fill="x", expand=True)

            toggle_btn = ctk.CTkButton(
                header,
                text="▾",
                width=34,
                height=28,
                font=ctk.CTkFont(size=16, weight="bold"),
                command=lambda key=etype: self._toggle_atp_section(key),
            )
            toggle_btn.pack(side="right")

            section_body = ctk.CTkFrame(section, fg_color="transparent")

            self._atp_section_widgets[etype] = {
                "section": section,
                "body": section_body,
                "toggle": toggle_btn,
                "cards": [],
                "section_visible": True,
            }
            self._apply_atp_section_visibility(etype)

            for element, editable_params in items:
                element_id = self._extract_element_identifier(element)
                title_prefix = etype.upper()
                card = ctk.CTkFrame(section_body, corner_radius=10, fg_color=("#EDEFF3", "#2B3138"))
                card.pack(fill="x", padx=10, pady=(8, 12))
                total_cards += 1

                inner_frame = ctk.CTkFrame(card, fg_color="transparent")
                inner_frame.pack(fill="x", padx=10, pady=8)

                ctk.CTkLabel(
                    inner_frame,
                    text=f"🔌 Linha {element_id}",
                    font=ctk.CTkFont(size=14, weight="bold"),
                    anchor="w",
                ).pack(fill="x", pady=(0, 6))

                card_search_parts = [title_prefix.lower(), element_id.lower(), f"linha {element_id.lower()}"]

                for param_name, meta in editable_params:
                    param_row = ctk.CTkFrame(inner_frame, fg_color="transparent")
                    param_row.pack(fill="x", pady=(2, 8))

                    ctk.CTkLabel(
                        param_row,
                        text=self._friendly_atp_param_label(param_name),
                        width=170,
                        anchor="w",
                    ).pack(side="left", padx=(0, 8))

                    control_frame = ctk.CTkFrame(param_row, fg_color="transparent")
                    control_frame.pack(side="left", fill="x", expand=True)

                    original_value = float(meta.get("original_value", meta.get("value", 0.0)))
                    line_index = int(element.get("line_index", -1))
                    field_start = int(meta.get("start", -1))
                    field_end = int(meta.get("end", -1))
                    raw_line = str(element.get("raw_line", ""))
                    field_text = raw_line[field_start:field_end] if 0 <= field_start < field_end <= len(raw_line) else ""
                    keep_trailing_dot = re.match(r"^[+-]?\d+\.$", field_text.strip()) is not None
                    value_var = tk.StringVar(value=self._format_param_value(original_value, keep_trailing_dot))

                    minus_btn = ctk.CTkButton(control_frame, text="-", width=34, height=30)
                    minus_btn.pack(side="left", padx=(0, 6))

                    entry = ctk.CTkEntry(control_frame, width=130, textvariable=value_var)
                    entry.pack(side="left", padx=(0, 6))
                    entry.configure(validate="key", validatecommand=vcmd_numeric)

                    plus_btn = ctk.CTkButton(control_frame, text="+", width=34, height=30)
                    plus_btn.pack(side="left", padx=(0, 6))

                    reset_btn = ctk.CTkButton(control_frame, text="↺", width=34, height=30)
                    reset_btn.pack(side="left", padx=(0, 8))

                    status_label = ctk.CTkLabel(
                        control_frame,
                        text="Original",
                        width=70,
                        anchor="center",
                        font=ctk.CTkFont(size=11, weight="bold"),
                    )
                    status_label.pack(side="left")

                    step = self._get_param_step(param_name, original_value)
                    card_search_parts.extend(
                        [
                            str(param_name).lower(),
                            self._friendly_atp_param_label(param_name).lower(),
                        ]
                    )

                    row_data = {
                        "line_index": line_index,
                        "name": f"Linha {element_id}",
                        "element_name": etype,
                        "parameter": param_name,
                        "editable": True,
                        "original_value": original_value,
                        "entry": entry,
                        "status_label": status_label,
                        "var": value_var,
                        "invalid": False,
                        "_updating": False,
                        "step": step,
                        "keep_trailing_dot": keep_trailing_dot,
                        "default_fg_color": entry.cget("fg_color"),
                        "default_border_color": entry.cget("border_color"),
                        "default_border_width": entry.cget("border_width"),
                        "default_text_color": entry.cget("text_color"),
                        "_repeat_job": None,
                        "_repeat_direction": 0,
                    }

                    minus_btn.bind("<ButtonPress-1>", lambda _e, r=row_data: self._start_adjust_repeat(r, -1))
                    minus_btn.bind("<ButtonRelease-1>", lambda _e, r=row_data: self._stop_adjust_repeat(r))
                    minus_btn.bind("<Leave>", lambda _e, r=row_data: self._stop_adjust_repeat(r))

                    plus_btn.bind("<ButtonPress-1>", lambda _e, r=row_data: self._start_adjust_repeat(r, +1))
                    plus_btn.bind("<ButtonRelease-1>", lambda _e, r=row_data: self._stop_adjust_repeat(r))
                    plus_btn.bind("<Leave>", lambda _e, r=row_data: self._stop_adjust_repeat(r))

                    reset_btn.configure(command=lambda r=row_data: self._reset_single_atp_parameter(r))

                    value_var.trace_add("write", lambda *_args, r=row_data: self._on_atp_entry_changed(r))

                    self._atp_param_rows.append(row_data)
                    self._set_atp_row_visual_state(row_data, "normal")

                self._atp_section_widgets[etype]["cards"].append(
                    {
                        "widget": card,
                        "search_text": " ".join(card_search_parts),
                        "visible": True,
                    }
                )

        if not self._atp_param_rows:
            self.atp_param_status_var.set("Nenhum componente editavel detectado")
            ctk.CTkLabel(
                self.atp_params_scroll_frame,
                text="Nao foram encontrados parametros editaveis para BRANCHES, SWITCHES ou SOURCES.",
                justify="left",
            ).pack(anchor="w", padx=8, pady=8)
            self.log("[ATP] Nenhum parametro editavel para interface por elementos")
            return

        self._refresh_atp_param_status()
        self._apply_atp_parameter_filter()
        self._refresh_atp_sweep_parameter_menu()
        self.log(
            f"[ATP] Parametros editaveis carregados: {len(self._atp_param_rows)} em {total_cards} elemento(s)"
        )

    def _collect_atp_parameter_overrides(self) -> list[dict]:
        """Coleta alterações de parametros ATP digitadas na GUI."""
        overrides = []
        invalid_items = []

        for row in self._atp_param_rows:
            if not bool(row.get("editable", True)):
                continue

            raw = row["var"].get().strip()
            name = row["name"]
            parameter = row["parameter"]
            line_index = row["line_index"]
            original_value = row["original_value"]

            if raw == "":
                invalid_items.append(f"{name} ({parameter}): valor vazio")
                continue

            normalized = raw.replace("D", "E").replace("d", "e")
            try:
                new_value = float(normalized)
            except ValueError:
                invalid_items.append(f"{name} ({parameter}): '{raw}'")
                continue

            if abs(new_value - original_value) <= 1e-15:
                continue

            self.parameter_overrides[(int(line_index), str(parameter))] = float(new_value)
            overrides.append(
                {
                    "line_index": line_index,
                    "name": name,
                    "parameter": parameter,
                    "new_value": new_value,
                    "old_value": original_value,
                }
            )

        # Remove chaves obsoletas (campo voltou ao valor original ou ficou inválido).
        valid_keys = {(int(o["line_index"]), str(o["parameter"])) for o in overrides}
        for key in list(self.parameter_overrides.keys()):
            if key not in valid_keys:
                self.parameter_overrides.pop(key, None)

        if invalid_items:
            details = "\n".join(invalid_items[:15])
            raise ValueError(f"Valores invalidos no editor de parametros ATP:\n{details}")

        return overrides

    def _build_atp_sweep_option_label(self, row: dict) -> str:
        line_number = int(row.get("line_index", -1)) + 1
        return (
            f"Linha {line_number} | {row.get('name', 'Elemento')} | "
            f"{self._friendly_atp_param_label(str(row.get('parameter', '')))}"
        )

    def _refresh_atp_sweep_parameter_menu(self):
        rows = [row for row in self._atp_param_rows if bool(row.get("editable", True))]
        self._atp_sweep_parameter_options = {}

        if self.atp_sweep_parameter_menu is None:
            return

        if not rows:
            placeholder = "Nenhum parametro carregado"
            self.atp_sweep_parameter_menu.configure(values=[placeholder], state="disabled")
            self.atp_sweep_parameter_var.set(placeholder)
            return

        values = []
        for row in rows:
            label = self._build_atp_sweep_option_label(row)
            values.append(label)
            self._atp_sweep_parameter_options[label] = row

        current = self.atp_sweep_parameter_var.get().strip()
        if current not in self._atp_sweep_parameter_options:
            current = values[0]

        self.atp_sweep_parameter_menu.configure(values=values, state="normal")
        self.atp_sweep_parameter_var.set(current)
        self._on_atp_sweep_parameter_selected(current)

    def _on_atp_sweep_parameter_selected(self, _selected: str | None = None):
        row = self._atp_sweep_parameter_options.get(self.atp_sweep_parameter_var.get().strip())
        if row is None:
            return

        self.atp_sweep_start_var.set(
            self._format_param_value(float(row["original_value"]), keep_trailing_dot=row["keep_trailing_dot"])
        )
        self.atp_sweep_stop_var.set(
            self._format_param_value(float(row["original_value"]), keep_trailing_dot=row["keep_trailing_dot"])
        )
        self.atp_sweep_step_var.set(self._format_param_value(float(row["step"])))

    def _parse_float_field(self, raw_value: str, field_name: str) -> float:
        raw = raw_value.strip()
        if not raw:
            raise ValueError(f"{field_name} nao informado")

        normalized = raw.replace("D", "E").replace("d", "e")
        try:
            return float(normalized)
        except ValueError as exc:
            raise ValueError(f"{field_name} invalido: {raw_value}") from exc

    def _postprocess_atp_sweep_lis(
        self,
        lis_path: Path,
        run_dir: Path,
        *,
        plot_options: dict,
        show_plots: bool,
        hide_errors: bool,
        only_comparative: bool,
        report_progress=None,
    ) -> dict[str, object]:
        parsed = parse_lis_once(lis_path, verbose=False)
        df, _stats_lines, summary = parsed.table_df, parsed.stats_lines, parsed.summary
        excel_path = None
        table_warning = None
        risk_record = None

        if df is None:
            table_warning = "Tabela de distribuicao de picos nao encontrada no .lis gerado"
            if report_progress is not None:
                report_progress(f"Aviso: {table_warning}")
        else:
            excel_path = run_dir / f"{lis_path.stem}.xlsx"
            save_df_to_excel_only(df, excel_path)

            computed_stats = None
            try:
                computed_stats = calcular_estatisticas_do_df(df)
                escrever_estatisticas_excel(excel_path, computed_stats, summary_from_lis=summary)
            except Exception as exc:
                if not hide_errors and report_progress is not None:
                    report_progress(f"Aviso: falha em estatisticas: {exc}")

            risk_config = plot_options.get("risk_config")
            if risk_config is not None and computed_stats is not None:
                try:
                    result = calculate_failure_risk(
                        computed_stats["mean"],
                        computed_stats["std_dev"],
                        risk_config,
                    )
                    risk_record = self._failure_risk_record(lis_path.name, result)
                except Exception as exc:
                    if not hide_errors and report_progress is not None:
                        report_progress(f"Aviso: falha no calculo de risco: {exc}")

            if not only_comparative:
                graph_name = f"grafico_{lis_path.stem}.png"
                self._criar_grafico_customizado(
                    excel_path,
                    run_dir,
                    graph_name,
                    plot_options,
                    mostrar=show_plots,
                )

        try:
            time_series_df = parsed.time_series_df
            if time_series_df is not None:
                if excel_path is None:
                    excel_path = run_dir / f"{lis_path.stem}.xlsx"
                save_time_series_to_excel(time_series_df, excel_path)
                criar_grafico_series_temporais(
                    time_series_df,
                    run_dir / f"series_temporais_{lis_path.stem}.png",
                    lis_name=lis_path.name,
                    mostrar=show_plots,
                )
        except Exception as exc:
            if not hide_errors and report_progress is not None:
                report_progress(f"Aviso: falha em series temporais: {exc}")

        return {
            "excel_path": str(excel_path) if excel_path else None,
            "table_warning": table_warning,
            "risk_record": risk_record,
        }

    def _cleanup_cancelled_atp_artifacts(
        self,
        *,
        source_atp: Path,
        execution_atp: Path,
        generated_lis: Path | None,
        simulation_dir: Path | None,
    ) -> None:
        """Remove somente artefatos incompletos da simulacao cancelada."""
        import shutil

        if simulation_dir is not None and simulation_dir.exists():
            shutil.rmtree(simulation_dir, ignore_errors=True)
        if generated_lis is not None and generated_lis.exists():
            try:
                generated_lis.unlink()
            except OSError:
                pass

        if execution_atp != source_atp and execution_atp.exists():
            try:
                execution_atp.unlink()
            except OSError:
                pass

    def _cancel_atp_execution(self):
        if not self._atp_running:
            self._show_info("Execucao ATP", "Nao existe simulacao ATP em andamento para cancelar.")
            return

        self._atp_cancel_event.set()
        if self.atp_cancel_button is not None:
            self.atp_cancel_button.configure(state="disabled", text="Cancelando...")
        self.atp_batch_status_var.set("Cancelamento solicitado")
        self._atp_runtime_status_text = "Cancelamento solicitado"
        self.atp_run_status_var.set("Status: cancelamento solicitado")
        self.log("[ATP] Cancelamento solicitado pelo usuario")

    def _on_atp_sweep_event(self, event: dict):
        message = str(event.get("message", "")).strip()
        if message:
            self.log(f"[ATP-SWEEP] {message}")
            self.atp_batch_status_var.set(message)
            self._atp_runtime_status_text = message

        total_runs = int(event.get("total_runs", 0) or 0)
        run_index = int(event.get("run_index", 0) or 0)
        event_type = str(event.get("type", ""))
        if self._atp_sweep_parallel_active and event_type in {
            "run_succeeded",
            "run_failed",
            "run_cancelled",
        }:
            self._atp_sweep_completed_runs.add(run_index)

        if total_runs:
            displayed_runs = (
                len(self._atp_sweep_completed_runs)
                if self._atp_sweep_parallel_active
                else min(run_index, total_runs)
            )
            self.atp_batch_progress_var.set(f"{displayed_runs} / {total_runs}")

        progress = event.get("progress")
        if progress is not None:
            if self._atp_sweep_parallel_active:
                if event_type in {"run_succeeded", "run_failed"} and total_runs:
                    self._set_atp_progress(len(self._atp_sweep_completed_runs) / total_runs)
                elif event_type in {"sweep_cancelled", "sweep_finished"}:
                    self._set_atp_progress(max(self._atp_progress_value, float(progress)))
            else:
                self._set_atp_progress(max(self._atp_progress_value, float(progress)))

    def _format_atp_sweep_summary(self, summary) -> str:
        output_label = (
            str(summary.output_dir)
            if summary.output_dir.exists()
            else "Nenhum resultado preservado"
        )
        lines = [
            f"Sweep finalizado em {summary.elapsed_seconds:.1f}s",
            f"Parametro: {summary.parameter.display_label}",
            f"Pasta de saida: {output_label}",
            (
                f"Total: {summary.total_runs} | Sucessos: {summary.success_count} | "
                f"Falhas: {summary.failure_count} | Canceladas: {summary.cancelled_count} | "
                f"Ignoradas: {summary.skipped_count}"
            ),
        ]

        if summary.cancelled:
            lines.append("Execucao interrompida por cancelamento do usuario.")
        if summary.stopped_on_error:
            lines.append("Execucao interrompida no primeiro erro, conforme configurado.")

        lines.append("")
        for result in summary.results:
            line = f"{result.run_index:03d} | valor={result.value:g} | {result.status.upper()}"
            if result.lis_path:
                line += f" | LIS={result.lis_path}"
            if result.error:
                line += f" | erro={result.error}"
            lines.append(line)

        return "\n".join(lines)

    def _on_atp_parameter_sweep_finished(self, summary):
        final_progress = min(1.0, summary.processed_count / max(1, summary.total_runs))
        completed_all = (
            not summary.cancelled
            and not summary.stopped_on_error
            and summary.processed_count == summary.total_runs
        )
        elapsed = self._set_atp_feedback_finished(success=completed_all, final_progress=final_progress)
        summary_text = self._format_atp_sweep_summary(summary)
        self._update_simulation_results(summary_text)

        risk_records = []
        for run_result in summary.results:
            analysis = run_result.analysis
            if not isinstance(analysis, dict) or not analysis.get("risk_record"):
                continue
            record = dict(analysis["risk_record"])
            record["label"] = (
                f"Execucao {run_result.run_index:03d} - "
                f"{summary.parameter.display_label} = {run_result.value:g}"
            )
            risk_records.append(record)
        if risk_records:
            self._write_failure_risk_report(
                summary.output_dir,
                "Sweep ATP",
                risk_records,
            )

        if summary.cancelled:
            self.status_var.set("Sweep cancelado")
            self.atp_batch_status_var.set("Sweep cancelado")
            self.atp_run_status_var.set(f"Status: cancelado ({elapsed:.1f}s)")
        elif summary.failure_count > 0:
            self.status_var.set("Sweep concluido com falhas")
            self.atp_batch_status_var.set("Sweep concluido com falhas")
            self.atp_run_status_var.set(f"Status: concluido com falhas ({elapsed:.1f}s)")
        else:
            self.status_var.set("Sweep concluido")
            self.atp_batch_status_var.set("Sweep concluido")
            self.atp_run_status_var.set(f"Status: concluido ({elapsed:.1f}s)")

        self.log(
            f"[ATP-SWEEP] Resumo final: {summary.success_count} sucesso(s), "
            f"{summary.failure_count} falha(s), {summary.cancelled_count} cancelada(s)"
        )

        if self.save_logs_var.get() and summary.output_dir.exists():
            try:
                log_file = summary.output_dir / f"log_sweep_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.txt"
                log_content = self.log_textbox.get("1.0", "end")
                log_file.write_text(log_content, encoding="utf-8")
                self.log(f"[ATP-SWEEP] Log salvo: {log_file.name}")
            except Exception as exc:
                self.log(f"[ATP-SWEEP] Aviso: nao foi possivel salvar log do sweep: {exc}")

        if self.open_output_var.get() and summary.output_dir.exists():
            _open_in_file_manager(summary.output_dir)

        if summary.cancelled:
            preserved_output = (
                str(summary.output_dir)
                if summary.output_dir.exists()
                else "Nenhum resultado preservado"
            )
            self._show_warning(
                "Sweep cancelado",
                f"O sweep foi interrompido apos {summary.processed_count} execucao(oes) concluida(s).",
                details=[("Resultados", preserved_output)],
            )
        elif summary.failure_count > 0:
            self._show_warning(
                "Sweep concluido com falhas",
                (
                    f"{summary.success_count} execucao(oes) concluida(s) e "
                    f"{summary.failure_count} falha(s)."
                ),
                details=[("Resultados", str(summary.output_dir))],
            )
        else:
            self._show_success(
                "Sweep concluido",
                f"{summary.success_count} execucao(oes) concluida(s) com sucesso.",
                details=[("Resultados", str(summary.output_dir))],
            )

    def _run_atp_parameter_sweep(self):
        if self._atp_running:
            self._show_info("Execucao ATP", "Ja existe uma execucao ATP em andamento.")
            return

        atp_file = self.atp_file_var.get().strip()
        if not atp_file or not Path(atp_file).exists():
            self._show_error("Erro", "Arquivo .atp nao encontrado.")
            return

        if not self._atp_param_rows:
            self._load_atp_parameters(show_dialog_errors=True)
            if not self._atp_param_rows:
                return

        selected_label = self.atp_sweep_parameter_var.get().strip()
        selected_row = self._atp_sweep_parameter_options.get(selected_label)
        if selected_row is None:
            self._show_error("Erro", "Selecione um parametro valido para o sweep.")
            return

        missing_insert_dependencies = get_missing_insert_dependencies(atp_file)
        if missing_insert_dependencies:
            details = [("Arquivo ATP", atp_file)]
            for line_no, target in missing_insert_dependencies[:10]:
                details.append((f"Linha {line_no}", target))
            self._show_error(
                "Dependencias ATP ausentes",
                "Nao e possivel iniciar o sweep. Arquivo(s) auxiliar(es) de $INSERT nao encontrado(s).",
                details=details,
            )
            return

        outdir_str = self.outdir_var.get().strip()
        if not outdir_str:
            self._show_error("Erro", "Pasta de saida nao informada.")
            return

        try:
            pending_overrides = self._collect_atp_parameter_overrides()
        except Exception as exc:
            self._show_error("Erro", "Nao foi possivel validar parametros ATP.", details=[("Detalhes", str(exc))])
            return

        if pending_overrides:
            confirm = messagebox.askyesno(
                "Ignorar alteracoes manuais",
                (
                    "A execucao em lote altera somente o parametro selecionado para o sweep.\n\n"
                    "As alteracoes manuais pendentes no editor ATP serao ignoradas nesta execucao.\n\n"
                    "Deseja continuar?"
                ),
            )
            if not confirm:
                return

        try:
            start = self._parse_float_field(self.atp_sweep_start_var.get(), "Inicio")
            stop = self._parse_float_field(self.atp_sweep_stop_var.get(), "Fim")
            step = self._parse_float_field(self.atp_sweep_step_var.get(), "Passo")
            sweep_values = generate_sweep_values(start, stop, step)
            risk_config = self._collect_risk_config()
        except Exception as exc:
            self._show_error("Erro", "Configuracao invalida para o sweep.", details=[("Detalhes", str(exc))])
            return

        show_plots = self.show_plots_var.get()
        hide_errors = self.hide_errors_var.get()
        only_comparative = self.only_comparative_var.get()
        parse_each_lis = self.atp_sweep_parse_lis_var.get()
        stop_on_error = self.atp_sweep_stop_on_error_var.get()
        plot_options = {
            'show_bars': self.plot_bars_var.get(),
            'show_points': self.plot_points_var.get(),
            'show_gaussian': self.plot_gaussian_var.get(),
            'show_cumulative': self.plot_cumulative_var.get(),
            'show_stats_box': self.plot_stats_box_var.get(),
            'risk_config': risk_config,
        }

        self._atp_cancel_event.clear()
        self._atp_sweep_completed_runs.clear()
        self.atp_batch_progress_var.set(f"0 / {len(sweep_values)}")
        self.atp_batch_elapsed_var.set("0s")
        self.atp_batch_status_var.set("Preparando sweep")
        self._set_atp_feedback_running(mode="batch", total_runs=len(sweep_values))
        self.status_var.set("Executando sweep ATP...")
        self._update_simulation_results("Executando sweep ATP em background... aguarde.\n")

        parameter_ref = SweepParameterRef(
            line_index=int(selected_row["line_index"]),
            parameter=str(selected_row["parameter"]),
            element_name=str(selected_row.get("element_name", "element")),
            label=selected_label,
        )

        cpu_count = os.cpu_count() or 1
        parallel_runs = 1
        if self.parallel_process_var.get():
            parallel_runs = min(
                len(sweep_values),
                max(1, min(4, (cpu_count + 1) // 2)),
            )
        self._atp_sweep_parallel_active = parallel_runs > 1

        def worker():
            def report_progress(message: str):
                self.after(0, lambda msg=message: self._on_atp_progress_message(msg))

            def event_callback(event: dict):
                self.after(0, lambda payload=event: self._on_atp_sweep_event(payload))

            def lis_parser(lis_path: Path, run_dir: Path, _value: float, _idx: int, _total: int):
                return self._postprocess_atp_sweep_lis(
                    lis_path,
                    run_dir,
                    plot_options=plot_options,
                    show_plots=show_plots,
                    hide_errors=hide_errors,
                    only_comparative=only_comparative,
                    report_progress=report_progress,
                )

            try:
                _ensure_pipeline_imports()
                summary = run_parameter_sweep(
                    base_atp_path=atp_file,
                    parameter_id=parameter_ref,
                    start=start,
                    stop=stop,
                    step=step,
                    output_dir=Path(outdir_str) / "batch_runs",
                    lis_parser=lis_parser if parse_each_lis else None,
                    solver_timeout=self._atp_timeout_sec,
                    continue_on_error=not stop_on_error,
                    cancel_event=self._atp_cancel_event,
                    event_callback=event_callback,
                    max_parallel_runs=parallel_runs,
                )
                self.after(0, lambda data=summary: self._on_atp_parameter_sweep_finished(data))
            except Exception as exc:
                self.after(0, lambda msg=str(exc): self._on_atp_simulation_finished(False, msg))

        threading.Thread(target=worker, daemon=True).start()

    def _export_atp_parameters_txt(self):
        """Exporta para TXT todos os parametros ATP detectados para conferência manual."""
        atp_file = self.atp_file_var.get().strip()
        if not atp_file or not Path(atp_file).exists():
            self._show_error("Erro", "Arquivo .atp nao encontrado.")
            return

        if not self._atp_elements_cache:
            self._load_atp_parameters(show_dialog_errors=False)

        if not self._atp_elements_cache:
            self._show_warning("Aviso", "Nenhum parametro detectado para exportar.")
            return

        editable = get_editable_parameters(self._atp_elements_cache)
        current_values = {
            (int(row.get("line_index", -1)), str(row.get("parameter", ""))): row["entry"].get().strip()
            for row in self._atp_param_rows
        }

        atp_path = Path(atp_file)
        default_name = f"{atp_path.stem}_parametros_detectados.txt"

        target = filedialog.asksaveasfilename(
            title="Salvar parametros detectados",
            defaultextension=".txt",
            initialfile=default_name,
            filetypes=[("Texto", "*.txt"), ("Todos", "*.*")],
        )
        if not target:
            return

        lines = [
            "RELATORIO DE PARAMETROS ATP DETECTADOS",
            f"Arquivo: {atp_path}",
            f"Gerado em: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
            "",
            f"Total de elementos detectados: {len(self._atp_elements_cache)}",
            f"Total de parametros editaveis: {len(editable)}",
            "",
        ]

        for idx, element in enumerate(self._atp_elements_cache):
            line_index = int(element.get("line_index", -1))
            etype = str(element.get("type", "")).upper()
            raw_line = str(element.get("raw_line", ""))
            lines.append(f"[{idx}] Linha {line_index + 1} - {etype}")
            lines.append(f"RAW: {raw_line}")

            element_params = [p for p in editable if int(p.get("element_index", -1)) == idx]
            if not element_params:
                lines.append("  (sem parametros editaveis)")
                lines.append("")
                continue

            for param in element_params:
                field = str(param.get("field", ""))
                label = str(param.get("label", field))
                detected = param.get("value")
                current = current_values.get((line_index, field), "")
                current_display = current if current != "" else str(detected)
                status = "editavel" if bool(param.get("editable", True)) else "valor_padrao"
                lines.append(f"  - {label} [{field}]")
                lines.append(f"    status: {status}")
                lines.append(f"    detectado: {detected}")
                lines.append(f"    campo_gui: {current_display}")

            lines.append("")

        try:
            Path(target).write_text("\n".join(lines), encoding="utf-8")
            self.log(f"[ATP] TXT de parametros exportado: {target}")
            self._show_success("Sucesso", "TXT de parametros exportado com sucesso.", details=[("Arquivo", target)])
        except Exception as e:
            self._show_error("Erro", "Falha ao salvar TXT de parametros.", details=[("Detalhes", str(e))])
    
    
    def _clear_filter(self):
        """Limpar filtro"""
        self.filter_var.set('')
        self.refresh_list()
    
    def _clear_output_folder(self):
        """Limpar todos os arquivos da pasta de saída"""
        outdir = Path(self.outdir_var.get())
        
        if not outdir.exists():
            self._show_info("Informação", "A pasta de saída não existe.", details=[("Pasta", outdir)])
            return
        
        # Contar arquivos
        all_files = list(outdir.rglob('*'))
        files_count = sum(1 for f in all_files if f.is_file())
        dirs_count = sum(1 for f in all_files if f.is_dir())
        
        if files_count == 0 and dirs_count == 0:
            self._show_info("Informação", "A pasta de saída já está vazia.")
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
            self._show_success(
                "Sucesso",
                "Pasta limpa com sucesso.",
                details=[("Resumo", f"{deleted_files} arquivo(s)\n{deleted_dirs} pasta(s)")],
            )
            
        except Exception as e:
            self.log(f"Erro ao limpar pasta: {e}")
            self._show_error("Erro", "Erro ao limpar pasta.", details=[("Detalhes", str(e))])

    def _cancel_processing(self):
        """Cancelar processamento em andamento"""
        self.cancel_event.set()
        self.log("Cancelamento solicitado...")

    def _clear_logs(self):
        """Limpar área de logs"""
        self.log_textbox.configure(state="normal")
        self.log_textbox.delete("1.0", "end")
        self.log_textbox.configure(state="disabled")
        self._log_entries_since_trim = 0
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
                self._show_success("Sucesso", "Logs salvos com sucesso.", details=[("Arquivo", file)])
            except Exception as e:
                self._show_error("Erro", "Falha ao salvar logs.", details=[("Detalhes", str(e))])
    
    def log(self, message: str):
        """Adiciona mensagem ao log"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        full_msg = f"[{timestamp}] {message}\n"
        self.log_textbox.configure(state="normal")
        self.log_textbox.insert("end", full_msg)

        self._log_entries_since_trim += 1
        if self._log_entries_since_trim >= self._log_trim_every:
            self._trim_log_if_needed()
            self._log_entries_since_trim = 0

        self.log_textbox.see("end")
        self.log_textbox.configure(state="disabled")

    def _trim_log_if_needed(self):
        """Mantem apenas as ultimas linhas do log para evitar crescimento indefinido."""
        try:
            line_idx = self.log_textbox.index("end-1c")
            line_count = int(line_idx.split(".")[0])
            overflow = line_count - self._log_max_lines
            if overflow > 0:
                self.log_textbox.delete("1.0", f"{overflow + 1}.0")
        except Exception:
            pass

    def _show_styled_dialog(self, title: str, message: str, level: str = "info", details: list | None = None):
        """Exibe dialogo modal customizado com layout mais organizado que messagebox."""
        palette = {
            "info": {"accent": "#2563eb"},
            "success": {"accent": "#15803d"},
            "warning": {"accent": "#b45309"},
            "error": {"accent": "#b91c1c"},
        }
        cfg = palette.get(level, palette["info"])

        dialog = ctk.CTkToplevel(self)
        dialog.title(title)
        dialog.transient(self)
        dialog.resizable(False, False)
        dialog.grab_set()

        container = ctk.CTkFrame(dialog, corner_radius=12)
        container.pack(fill="both", expand=True, padx=14, pady=14)

        header = ctk.CTkFrame(container, corner_radius=10, fg_color=cfg["accent"])
        header.pack(fill="x", padx=12, pady=(12, 8))

        ctk.CTkLabel(
            header,
            text=title,
            font=ctk.CTkFont(size=16, weight="bold"),
            text_color="white",
            anchor="center",
            justify="center",
        ).pack(fill="x", padx=12, pady=10)

        body = ctk.CTkFrame(container, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=14, pady=(4, 8))

        ctk.CTkLabel(
            body,
            text=message,
            justify="left",
            anchor="w",
            wraplength=620,
            font=ctk.CTkFont(size=13),
        ).pack(fill="x", pady=(0, 8))

        copy_idle_text = "⧉"
        copy_done_text = "✓"
        copy_idle_fg = ("#E6E8EB", "#2B3036")
        copy_idle_hover = ("#D9DDE2", "#353B43")
        copy_idle_text_color = ("#2D3748", "#DCE3EA")
        copy_idle_border = ("#C9D1D9", "#4A5568")
        copy_done_fg = ("#2F6F44", "#2F6F44")

        def _copy_detail(content: str, button=None):
            try:
                self.clipboard_clear()
                self.clipboard_append(content)
                self.update_idletasks()
            except Exception:
                return

            if button is not None:
                try:
                    button.configure(
                        text=copy_done_text,
                        fg_color=copy_done_fg,
                        hover_color=copy_done_fg,
                        text_color="#FFFFFF",
                        border_width=0,
                        border_color=copy_done_fg,
                    )
                    dialog.after(
                        1200,
                        lambda btn=button: btn.winfo_exists()
                        and btn.configure(
                            text=copy_idle_text,
                            fg_color=copy_idle_fg,
                            hover_color=copy_idle_hover,
                            text_color=copy_idle_text_color,
                            border_width=1,
                            border_color=copy_idle_border,
                        ),
                    )
                except Exception:
                    pass

        if details:
            for label, value in details:
                section = ctk.CTkFrame(body, corner_radius=8)
                section.pack(fill="x", pady=(6, 6))

                section_header = ctk.CTkFrame(section, fg_color="transparent")
                section_header.pack(fill="x", padx=10, pady=(8, 2))

                ctk.CTkLabel(
                    section_header,
                    text=label,
                    anchor="w",
                    font=ctk.CTkFont(size=12, weight="bold"),
                ).pack(side="left", fill="x", expand=True)

                copy_btn = ctk.CTkButton(
                    section_header,
                    text=copy_idle_text,
                    width=34,
                    height=24,
                    font=ctk.CTkFont(size=14, weight="bold"),
                    fg_color=copy_idle_fg,
                    hover_color=copy_idle_hover,
                    text_color=copy_idle_text_color,
                    border_width=1,
                    border_color=copy_idle_border,
                )
                copy_btn.pack(side="right")
                copy_btn.configure(command=lambda v=str(value), btn=copy_btn: _copy_detail(v, btn))

                text_box = ctk.CTkTextbox(section, height=62, wrap="char")
                text_box.pack(fill="x", padx=10, pady=(0, 8))
                text_box.insert("1.0", str(value))
                text_box.configure(state="disabled")

        buttons = ctk.CTkFrame(container, fg_color="transparent")
        buttons.pack(fill="x", padx=12, pady=(0, 12))

        def _close_dialog(_event=None):
            try:
                dialog.grab_release()
            except Exception:
                pass
            dialog.destroy()
            self.focus_force()

        ctk.CTkButton(
            buttons,
            text="OK",
            width=120,
            fg_color=cfg["accent"],
            hover_color=cfg["accent"],
            command=_close_dialog,
        ).pack(pady=2)

        dialog.bind("<Return>", _close_dialog)
        dialog.bind("<Escape>", _close_dialog)

        dialog.update_idletasks()
        width = max(560, min(dialog.winfo_reqwidth(), 860))
        height = max(220, min(dialog.winfo_reqheight(), 620))

        parent_x = self.winfo_rootx()
        parent_y = self.winfo_rooty()
        parent_w = self.winfo_width()
        parent_h = self.winfo_height()
        pos_x = parent_x + (parent_w - width) // 2
        pos_y = parent_y + (parent_h - height) // 2
        dialog.geometry(f"{width}x{height}+{pos_x}+{pos_y}")
        dialog.lift()
        dialog.focus_force()
        self.wait_window(dialog)

    def _show_info(self, title: str, message: str, details: list | None = None):
        self._show_styled_dialog(title, message, level="info", details=details)

    def _show_success(self, title: str, message: str, details: list | None = None):
        self._show_styled_dialog(title, message, level="success", details=details)

    def _show_warning(self, title: str, message: str, details: list | None = None):
        self._show_styled_dialog(title, message, level="warning", details=details)

    def _show_error(self, title: str, message: str, details: list | None = None):
        self._show_styled_dialog(title, message, level="error", details=details)

    def _run_atp_simulation(self):
        """Executa o solver ATP em background e atualiza a GUI ao finalizar."""
        if self._atp_running:
            self._show_info("Simulacao ATP", "Ja existe uma simulacao ATP em andamento.")
            return

        atp_file = self.atp_file_var.get().strip()
        if not atp_file or not Path(atp_file).exists():
            self._show_error("Erro", "Arquivo .atp nao encontrado.")
            return

        missing_insert_dependencies = get_missing_insert_dependencies(atp_file)
        if missing_insert_dependencies:
            preview_limit = 10
            preview = missing_insert_dependencies[:preview_limit]
            details = [
                ("Arquivo ATP", atp_file),
                ("Diretorio base", str(Path(atp_file).parent)),
                ("Dependencias ausentes", str(len(missing_insert_dependencies))),
            ]
            for line_no, target in preview:
                details.append((f"Linha {line_no}", target))
            if len(missing_insert_dependencies) > preview_limit:
                details.append(("Outros", f"+{len(missing_insert_dependencies) - preview_limit} item(ns)"))

            self.log("Validacao pre-run ATP falhou: dependencia(s) $INSERT ausente(s).")
            for line_no, target in missing_insert_dependencies:
                self.log(f" - Linha {line_no}: {target}")

            self._update_simulation_results(
                "Falha na validacao pre-run ATP:\n"
                + "\n".join([f"- Linha {line_no}: {target}" for line_no, target in missing_insert_dependencies])
            )
            self._show_error(
                "Dependencias ATP ausentes",
                "Nao e possivel iniciar a simulacao. Arquivo(s) auxiliar(es) de $INSERT nao encontrado(s).",
                details=details,
            )
            return

        outdir_str = self.outdir_var.get().strip()
        if not outdir_str:
            self._show_error("Erro", "Pasta de saida nao informada.")
            return

        show_plots = self.show_plots_var.get()
        hide_errors = self.hide_errors_var.get()
        only_comparative = self.only_comparative_var.get()
        try:
            risk_config = self._collect_risk_config()
        except ValueError as exc:
            self._show_error(
                "Risco de falha",
                "Os parametros do calculo de risco sao invalidos.",
                details=[("Detalhes", str(exc))],
            )
            return

        plot_options = {
            'show_bars': self.plot_bars_var.get(),
            'show_points': self.plot_points_var.get(),
            'show_gaussian': self.plot_gaussian_var.get(),
            'show_cumulative': self.plot_cumulative_var.get(),
            'show_stats_box': self.plot_stats_box_var.get(),
            'risk_config': risk_config,
        }

        try:
            atp_overrides = self._collect_atp_parameter_overrides()
        except Exception as e:
            self._show_error("Erro", "Nao foi possivel validar parametros ATP.", details=[("Detalhes", str(e))])
            return

        if atp_overrides and not self._confirm_atp_overrides_preview(atp_overrides):
            self.log("[ATP] Execucao cancelada pelo usuario no preview de alteracoes")
            return

        self._set_atp_feedback_running(mode="single", total_runs=1)
        self.status_var.set("Executando simulacao ATP...")
        self.log(f"Iniciando simulacao ATP para: {atp_file}")
        self._update_simulation_results("Executando simulacao ATP e pos-processamento... aguarde.\n")

        def worker():
            def report_progress(message: str):
                self.after(0, lambda msg=message: self._on_atp_progress_message(msg))

            parametrized_exec_atp = None
            generated_lis_path = None
            execution_atp_path = Path(atp_file)
            sim_outdir = None
            try:
                _ensure_pipeline_imports()
                import shutil

                if atp_overrides:
                    report_progress("Applying ATP parameter overrides...")
                    elements = parse_atp_file_cached(execution_atp_path)
                    original_lines = _read_text_lines_preserve_newlines(execution_atp_path)
                    for override in atp_overrides:
                        update_parameter(
                            elements,
                            element_name=override["name"],
                            new_value=override["new_value"],
                            line_index=override["line_index"],
                            parameter_name=override["parameter"],
                        )

                    base_param_name = f"{execution_atp_path.stem}_param{execution_atp_path.suffix}"
                    parametrized_exec_atp = execution_atp_path.parent / base_param_name
                    if parametrized_exec_atp.exists():
                        suffix_idx = 2
                        while True:
                            candidate = (
                                execution_atp_path.parent
                                / f"{execution_atp_path.stem}_param_{suffix_idx}{execution_atp_path.suffix}"
                            )
                            if not candidate.exists():
                                parametrized_exec_atp = candidate
                                break
                            suffix_idx += 1

                    write_atp_file(elements, original_lines, parametrized_exec_atp)
                    execution_atp_path = parametrized_exec_atp
                    report_progress(f"Parameterized ATP ready: {execution_atp_path.name}")

                report_progress("Running ATP solver...")
                generated_lis_path = Path(
                    run_atp_solver(
                        str(execution_atp_path),
                        timeout=self._atp_timeout_sec,
                        status_callback=report_progress,
                        cancel_event=self._atp_cancel_event,
                    )
                )

                if self._atp_cancel_event.is_set():
                    raise ATPExecutionCancelled("Simulacao ATP cancelada pelo usuario")

                report_progress("Preparing output folder...")
                base_outdir = Path(outdir_str)
                base_outdir.mkdir(parents=True, exist_ok=True)
                timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
                sim_outdir = base_outdir / timestamp
                sim_outdir.mkdir(parents=True, exist_ok=True)

                lis_target = sim_outdir / generated_lis_path.name
                if lis_target.exists():
                    lis_target = sim_outdir / f"{generated_lis_path.stem}_{timestamp}{generated_lis_path.suffix}"

                try:
                    if generated_lis_path.resolve() != lis_target.resolve():
                        lis_target = Path(shutil.move(str(generated_lis_path), str(lis_target)))
                except Exception:
                    # Se mover falhar, segue com o .lis no caminho original.
                    lis_target = generated_lis_path

                generated_atp_snapshot = None
                if parametrized_exec_atp is not None and parametrized_exec_atp.exists():
                    generated_atp_snapshot = sim_outdir / parametrized_exec_atp.name
                    try:
                        generated_atp_snapshot = Path(
                            shutil.move(str(parametrized_exec_atp), str(generated_atp_snapshot))
                        )
                    except Exception:
                        try:
                            shutil.copy2(str(parametrized_exec_atp), str(generated_atp_snapshot))
                        except Exception:
                            generated_atp_snapshot = None

                if atp_overrides:
                    metadata = {
                        "source_atp": atp_file,
                        "executed_atp": str(execution_atp_path),
                        "applied_overrides": atp_overrides,
                        "generated_atp_in_output": str(generated_atp_snapshot) if generated_atp_snapshot else None,
                    }
                    metadata_path = sim_outdir / "parametros_aplicados.json"
                    metadata_path.write_text(json.dumps(metadata, indent=2), encoding="utf-8")

                if self._atp_cancel_event.is_set():
                    raise ATPExecutionCancelled("Simulacao ATP cancelada pelo usuario")

                report_progress("Parsing LIS and generating tables...")
                parsed = parse_lis_once(lis_target, verbose=False)
                df, stats_lines, summary = parsed.table_df, parsed.stats_lines, parsed.summary
                excel_path = None
                table_warning = None
                risk_records = []
                if df is None:
                    table_warning = "Tabela de distribuicao de picos nao encontrada no .lis gerado"
                    report_progress(f"Warning: {table_warning}")
                else:
                    excel_path = sim_outdir / f"{lis_target.stem}.xlsx"
                    save_df_to_excel_only(df, excel_path)

                    computed_stats = None
                    try:
                        computed_stats = calcular_estatisticas_do_df(df)
                        escrever_estatisticas_excel(excel_path, computed_stats, summary_from_lis=summary)
                    except Exception as e:
                        if not hide_errors:
                            report_progress(f"Warning: falha em estatisticas: {e}")

                    if risk_config is not None and computed_stats is not None:
                        try:
                            result = calculate_failure_risk(
                                computed_stats["mean"],
                                computed_stats["std_dev"],
                                risk_config,
                            )
                            risk_records.append(
                                self._failure_risk_record(lis_target.name, result)
                            )
                        except Exception as exc:
                            if not hide_errors:
                                report_progress(f"Warning: falha no calculo de risco: {exc}")

                    if not only_comparative:
                        report_progress("Generating chart from analyzed data...")
                        graph_name = f"grafico_{lis_target.stem}.png"
                        self._criar_grafico_customizado(
                            excel_path,
                            sim_outdir,
                            graph_name,
                            plot_options,
                            mostrar=show_plots,
                        )
                    else:
                        report_progress("Skipping individual chart (only comparative option enabled).")

                if self._atp_cancel_event.is_set():
                    raise ATPExecutionCancelled("Simulacao ATP cancelada pelo usuario")

                report_progress("Processing time series...")
                try:
                    time_series_df = parsed.time_series_df
                    if time_series_df is not None:
                        if excel_path is None:
                            excel_path = sim_outdir / f"{lis_target.stem}.xlsx"
                        save_time_series_to_excel(time_series_df, excel_path)
                        criar_grafico_series_temporais(
                            time_series_df,
                            sim_outdir / f"series_temporais_{lis_target.stem}.png",
                            lis_name=lis_target.name,
                            mostrar=show_plots,
                        )
                except Exception as e:
                    if not hide_errors:
                        report_progress(f"Warning: falha em series temporais: {e}")

                if self._atp_cancel_event.is_set():
                    raise ATPExecutionCancelled("Simulacao ATP cancelada pelo usuario")

                risk_report_path = self._write_failure_risk_report(
                    sim_outdir,
                    "Simulacao ATP",
                    risk_records,
                )
                if self._atp_cancel_event.is_set():
                    raise ATPExecutionCancelled("Simulacao ATP cancelada pelo usuario")

                payload = {
                    "lis_path": str(lis_target),
                    "outdir": str(sim_outdir),
                    "excel_path": str(excel_path) if excel_path else None,
                    "applied_overrides": len(atp_overrides),
                    "table_warning": table_warning,
                    "risk_report_path": str(risk_report_path) if risk_report_path else None,
                }
                self.after(0, lambda data=payload: self._on_atp_simulation_finished(True, data))
            except ATPExecutionCancelled:
                self._cleanup_cancelled_atp_artifacts(
                    source_atp=Path(atp_file),
                    execution_atp=execution_atp_path,
                    generated_lis=generated_lis_path,
                    simulation_dir=sim_outdir,
                )
                self.after(0, self._on_atp_simulation_cancelled)
            except Exception as e:
                error_msg = str(e)
                self.after(0, lambda msg=error_msg: self._on_atp_simulation_finished(False, msg))
            finally:
                if parametrized_exec_atp is not None and parametrized_exec_atp.exists():
                    try:
                        parametrized_exec_atp.unlink()
                    except Exception:
                        pass

        threading.Thread(target=worker, daemon=True).start()

    def _on_atp_progress_message(self, message: str):
        """Recebe mensagens de progresso do runner ATP no thread principal."""
        self.log(f"[ATP] {message}")
        self.atp_batch_status_var.set(message)
        self._atp_runtime_status_text = message

        text = message.lower()
        if "process started" in text:
            self._set_atp_progress(0.03)
        elif "process finished" in text:
            self._set_atp_progress(0.88)
        elif "waiting for lis generation/stabilization" in text:
            self._set_atp_progress(0.92)
        elif "lis ready" in text:
            self._set_atp_progress(0.98)

    def _on_atp_simulation_cancelled(self):
        """Restaura a interface apos o usuario interromper a simulacao."""
        elapsed = self._set_atp_feedback_finished(success=False, final_progress=self._atp_progress_value)
        self.status_var.set("Simulacao cancelada")
        self.atp_batch_status_var.set("Simulacao cancelada")
        self.atp_run_status_var.set(f"Status: cancelado ({elapsed:.1f}s)")
        self.log(f"Simulacao ATP cancelada apos {elapsed:.1f}s")
        self._update_simulation_results(
            f"Simulacao ATP cancelada pelo usuario apos {elapsed:.1f}s."
        )

    def _on_atp_simulation_finished(self, success: bool, payload):
        """Atualiza a GUI quando a simulacao ATP termina (thread principal)."""
        elapsed = self._set_atp_feedback_finished(success=success)

        if success:
            lis_path = payload["lis_path"] if isinstance(payload, dict) else str(payload)
            outdir = payload.get("outdir") if isinstance(payload, dict) else None
            overrides_count = payload.get("applied_overrides", 0) if isinstance(payload, dict) else 0
            table_warning = payload.get("table_warning") if isinstance(payload, dict) else None
            risk_report_path = payload.get("risk_report_path") if isinstance(payload, dict) else None
            self.status_var.set("Simulation completed")
            self.atp_batch_progress_var.set("1 / 1")
            self.atp_batch_status_var.set("Simulacao concluida")
            self.log(f"Simulacao concluida. LIS gerado em: {lis_path}")
            if outdir:
                self.log(f"Resultados da analise salvos em: {outdir}")
            if overrides_count:
                self.log(f"Parametros ATP aplicados nesta execucao: {overrides_count}")
            if table_warning:
                self.log(f"Aviso: {table_warning}")

            self._update_simulation_results(
                f"Simulation completed in {elapsed:.1f}s\nLIS file: {lis_path}\nOutput folder: {outdir if outdir else '(nao informado)'}\nParameter overrides: {overrides_count}"
                + (f"\nRisk report: {risk_report_path}" if risk_report_path else "")
                + (f"\nWarning: {table_warning}" if table_warning else "")
            )

            if self.save_logs_var.get() and outdir:
                try:
                    log_file = Path(outdir) / f"log_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.txt"
                    log_content = self.log_textbox.get("1.0", "end")
                    log_file.write_text(log_content, encoding="utf-8")
                    self.log(f"Log salvo: {log_file.name}")
                except Exception as e:
                    self.log(f"Aviso: nao foi possivel salvar log da simulacao ATP: {e}")

            if self.open_output_var.get() and outdir:
                _open_in_file_manager(Path(outdir))

            self._show_success(
                "Simulacao concluida",
                f"Simulacao finalizada em {elapsed:.1f}s.",
                details=[
                    ("Resultados", outdir if outdir else "(nao informado)"),
                ],
            )
        else:
            self.status_var.set("Pronto")
            error_msg = str(payload)
            self.atp_batch_progress_var.set("0 / 1")
            self.atp_batch_status_var.set("Erro na simulacao")
            details = [("Detalhes", error_msg)]

            marker = "Trecho do LIS:"
            if marker in error_msg:
                head, excerpt = error_msg.split(marker, 1)
                details = [("Motivo", head.strip())]
                excerpt_text = excerpt.strip()
                if excerpt_text:
                    details.append(("Trecho do LIS", excerpt_text))

            self.log(f"Erro na simulacao ATP: {error_msg}")
            self._update_simulation_results(
                f"Erro na simulacao ATP apos {elapsed:.1f}s:\n{error_msg}"
            )
            self._show_error(
                "Erro na simulacao ATP",
                f"Falha apos {elapsed:.1f}s.",
                details=details,
            )

    def _set_atp_feedback_running(self, mode: str = "single", total_runs: int = 1):
        """Ativa indicadores visuais de simulacao ATP em andamento."""
        # Reset explicito para nova simulacao (evita manter 100% da execucao anterior).
        self._atp_running = False
        self._set_atp_progress(0.0)

        self._atp_running = True
        self._atp_active_mode = mode
        self._atp_started_at = time.time()
        self._atp_runtime_status_text = "Executando"
        self._atp_cancel_event.clear()
        self._set_atp_progress(0.0)
        self.atp_run_button.configure(state="disabled", text="Executando...")
        if self.atp_sweep_run_button is not None:
            self.atp_sweep_run_button.configure(state="disabled", text="Executando...")
        if self.atp_cancel_button is not None:
            self.atp_cancel_button.configure(state="normal", text="Cancelar")
        self.atp_batch_progress_var.set(f"0 / {max(1, int(total_runs))}")
        self.atp_batch_elapsed_var.set("0s")
        self.atp_batch_status_var.set("Executando")
        self.atp_run_status_var.set("Status: executando (0s)")
        self.after(1000, self._tick_atp_running_status)

    def _set_atp_feedback_finished(self, success: bool, final_progress: float | None = None) -> float:
        """Desativa indicadores visuais e retorna tempo decorrido."""
        elapsed = 0.0
        if self._atp_started_at is not None:
            elapsed = time.time() - self._atp_started_at

        self._atp_running = False
        self._atp_started_at = None
        self._atp_active_mode = None
        target_progress = final_progress
        if target_progress is None:
            target_progress = 1.0 if success else 0.0
        self._set_atp_progress(target_progress)
        self.atp_run_button.configure(state="normal", text="Executar Simulação")
        if self.atp_sweep_run_button is not None:
            self.atp_sweep_run_button.configure(state="normal", text="Executar em Lote")
        if self.atp_cancel_button is not None:
            self.atp_cancel_button.configure(state="disabled", text="Cancelar")
        self.atp_batch_elapsed_var.set(f"{elapsed:.1f}s")

        if success:
            self.atp_run_status_var.set(f"Status: concluido ({elapsed:.1f}s)")
        else:
            self.atp_run_status_var.set(f"Status: erro ({elapsed:.1f}s)")

        return elapsed

    def _tick_atp_running_status(self):
        """Atualiza o status de tempo da execucao ATP a cada segundo."""
        if not self._atp_running or self._atp_started_at is None:
            return

        elapsed_float = time.time() - self._atp_started_at
        elapsed = int(elapsed_float)
        self.atp_batch_elapsed_var.set(f"{elapsed}s")
        # Inicia do zero e acelera de forma suave apos o primeiro segundo.
        base_elapsed = max(0.0, elapsed_float - 1.0)
        timed_progress = 0.80 * (1.0 - math.exp(-base_elapsed / 45.0))
        timed_progress = min(0.85, timed_progress)
        self._set_atp_progress(timed_progress)
        status_text = self._atp_runtime_status_text or "executando"
        self.atp_run_status_var.set(f"Status: {status_text} ({elapsed}s)")
        self.after(1000, self._tick_atp_running_status)

    def _set_atp_progress(self, value: float):
        """Define progresso ATP entre 0 e 1, sem retroceder durante execucao."""
        clamped = max(0.0, min(1.0, float(value)))
        if self._atp_running:
            self._atp_progress_value = max(self._atp_progress_value, clamped)
        else:
            self._atp_progress_value = clamped
        self.atp_progress.set(self._atp_progress_value)

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
                self._show_info("Abrir arquivo", "Abra manualmente.", details=[("Arquivo", file_path)])
        except Exception as e:
            self._show_error("Erro ao abrir", "Falha ao abrir arquivo.", details=[("Detalhes", str(e))])
    
    def _process_selected(self):
        """Processar arquivos .lis selecionados"""
        # Coletar arquivos selecionados
        selected_files = []
        for file_str, var in self.file_selection_vars.items():
            if var.get():
                selected_files.append(Path(file_str))
        
        if not selected_files:
            self._show_warning("Aviso", "Nenhum arquivo selecionado.", details=[("Ação", "Marque os arquivos que deseja processar.")])
            return
        
        try:
            risk_config = self._collect_risk_config()
        except ValueError as exc:
            self._show_error(
                "Risco de falha",
                "Parametros invalidos para o calculo do risco.",
                details=[("Detalhes", str(exc))],
            )
            return

        self.log(f"Iniciando processamento de {len(selected_files)} arquivo(s)...")
        self.status_var.set(f"Processando {len(selected_files)} arquivo(s)...")
        
        # Mostrar botão de cancelar
        self._show_main_cancel()
        
        def worker():
            try:
                _ensure_pipeline_imports()
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
                    'risk_config': risk_config,
                }
                
                excel_paths = []  # Armazenar paths dos Excel para comparativo
                risk_records = []
                total = len(selected_files)
                
                # NOTA: Processamento paralelo (parallel_process) desabilitado por segurança
                # Para evitar conflitos de matplotlib em threads múltiplas
                
                for idx, lis_path in enumerate(selected_files, start=1):
                    if self.cancel_event.is_set():
                        self.log("Processamento cancelado pelo usuário")
                        break
                    
                    # Atualizar progresso
                    progress = idx / total
                    self._set_main_progress(progress)
                    self.status_var.set(f"Processando {lis_path.name}... ({idx}/{total})")
                    
                    self.log(f"Processando: {lis_path.name}")
                    
                    # Parse do .lis
                    parsed = parse_lis_once(lis_path, verbose=False)
                    df, stats_lines, summary = parsed.table_df, parsed.stats_lines, parsed.summary
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
                    computed_stats = None
                    try:
                        computed_stats = calcular_estatisticas_do_df(df)
                        escrever_estatisticas_excel(excel_path, computed_stats, summary_from_lis=summary)
                    except Exception as e:
                        error_msg = f"Erro ao calcular estatísticas: {e}"
                        if not self.hide_errors_var.get():
                            self.log(error_msg)
                            self.after(
                                0,
                                lambda msg=error_msg: self._show_warning(
                                    "Aviso",
                                    "Falha ao calcular estatisticas.",
                                    details=[("Detalhes", msg)],
                                ),
                            )

                    if risk_config is not None and computed_stats is not None:
                        try:
                            result = calculate_failure_risk(
                                computed_stats["mean"],
                                computed_stats["std_dev"],
                                risk_config,
                            )
                            risk_records.append(self._failure_risk_record(lis_path.name, result))
                        except Exception as exc:
                            if not self.hide_errors_var.get():
                                self.log(f"Falha ao calcular risco para {lis_path.name}: {exc}")
                    
                    # Gráfico individual (se não for modo "só comparativo")
                    if not self.only_comparative_var.get():
                        graph_name = f"grafico_{base_name}.png"
                        self._criar_grafico_customizado(excel_path, outdir, graph_name, plot_options, mostrar=self.show_plots_var.get())
                    
                    # Séries temporais
                    try:
                        time_series_df = parsed.time_series_df
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
                if risk_records:
                    self._write_failure_risk_report(
                        outdir,
                        "Analise de arquivos LIS",
                        risk_records,
                    )
                self._set_main_progress(1.0)
                self.status_var.set(f"Processamento concluído! {len(selected_files)} arquivo(s)")
                self.log(f"Processamento finalizado com sucesso em: {outdir.name}")
                
                if self.open_output_var.get():
                    _open_in_file_manager(outdir)
                
                self.after(
                    0,
                    lambda total=len(selected_files), result_outdir=str(outdir): self._show_success(
                        "Processamento concluido",
                        f"{total} arquivo(s) processado(s) com sucesso.",
                        details=[("Resultados", result_outdir)],
                    ),
                )
                
            except Exception as e:
                self.log(f"Erro durante processamento: {e}")
                self.after(
                    0,
                    lambda msg=str(e): self._show_error(
                        "Erro",
                        "Erro durante processamento.",
                        details=[("Detalhes", msg)],
                    ),
                )
            finally:
                self._hide_main_cancel()
                self.cancel_event.clear()
                self._set_main_progress(0)
                if not self.cancel_event.is_set():
                    self.status_var.set("Pronto")
        
        threading.Thread(target=worker, daemon=True).start()
    
    def _criar_grafico_customizado(self, excel_path: Path, outdir: Path, output_name: str, plot_options: dict, mostrar: bool = False):
        """Cria gráfico individual com opções customizadas de visualização"""
        try:
            _ensure_pipeline_imports()
            _ensure_plot_imports()
            
            res = obter_xy_e_stats_de_excel(excel_path)
            if res is None:
                self.log(f"Erro ao extrair dados de {excel_path.name}")
                return
            
            x, y, mu, sigma = res
            risk_result = None
            risk_config = plot_options.get("risk_config")
            if risk_config is not None:
                try:
                    risk_result = calculate_failure_risk(float(mu), float(sigma), risk_config)
                except Exception:
                    pass
            
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
                        f"μ = {_safe_float(mu):.6g}\n"
                        f"σ = {_safe_float(sigma):.6g}\n"
                        f"Mediana = {_safe_float(computed_stats.get('median', float('nan'))):.6g}\n"
                        f"CV = {_safe_float(computed_stats.get('cv', float('nan'))):.6g}\n"
                        f"R² = {_safe_float(computed_stats.get('r2', float('nan'))):.5g}"
                    )
                    bbox_props = dict(boxstyle="round,pad=0.6", fc="white", ec="0.4", alpha=0.9)
                    stats_y = 0.74 if risk_result is not None else 0.95
                    ax.text(0.98, stats_y, stats_text, transform=ax.transAxes, fontsize=9,
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
            
            if risk_result is not None:
                ax.text(
                    0.98,
                    0.96,
                    (
                        f"Risco de falha: {risk_result.risk * 100.0:.6g}%\n"
                        f"R = {risk_result.risk:.3e}"
                    ),
                    transform=ax.transAxes,
                    ha="right",
                    va="top",
                    fontsize=9,
                    fontweight="bold",
                    color="#444444",
                    zorder=20,
                    bbox=dict(
                        boxstyle="round,pad=0.5",
                        facecolor="white",
                        edgecolor="#666666",
                        alpha=0.96,
                    ),
                )

            fig.tight_layout()
            fig.savefig(out_png, dpi=220, bbox_inches='tight')
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
            _ensure_pipeline_imports()
            _ensure_plot_imports()
            
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
            
            # Caixa de estatisticas resumida.
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
    folder = Path(sys.argv[1]) if len(sys.argv) > 1 else Path.cwd() / "data" / "samples" / "ACP"
    outdir = Path(sys.argv[2]) if len(sys.argv) > 2 else Path.cwd() / "Simulation_Result"
    launch_gui(folder, outdir)
