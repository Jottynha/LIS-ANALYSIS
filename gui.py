import threading
import traceback
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
except Exception as e:
    # fallback: erro será exibido quando tentar abrir GUI via main
    raise


def _scan_lis(folder: Path):
    folder = Path(folder)
    return sorted(folder.glob('*.lis'), key=lambda f: f.stat().st_mtime, reverse=True)


def launch_gui(folder: Path, outdir: Path, start_index: int = 1):
    folder = Path(folder)
    outdir = Path(outdir)

    root = tk.Tk()
    root.title('LIS Analysis — Interface')

    folder_var = tk.StringVar(value=str(folder))
    outdir_var = tk.StringVar(value=str(outdir))
    start_idx_var = tk.IntVar(value=start_index)

    # Top: seleção de pasta e saída
    top = ttk.Frame(root, padding=10)
    top.pack(fill='x')

    ttk.Label(top, text='Pasta (.lis):').grid(row=0, column=0, sticky='w')
    ent_folder = ttk.Entry(top, textvariable=folder_var, width=60)
    ent_folder.grid(row=0, column=1, sticky='we', padx=6)
    def choose_folder():
        sel = filedialog.askdirectory(initialdir=folder_var.get() or '.')
        if sel:
            folder_var.set(sel)
            refresh_list()
    ttk.Button(top, text='Escolher...', command=choose_folder).grid(row=0, column=2, sticky='w')

    ttk.Label(top, text='Saída:').grid(row=1, column=0, sticky='w', pady=(6,0))
    ent_out = ttk.Entry(top, textvariable=outdir_var, width=60)
    ent_out.grid(row=1, column=1, sticky='we', padx=6, pady=(6,0))
    def choose_out():
        sel = filedialog.askdirectory(initialdir=outdir_var.get() or '.')
        if sel:
            outdir_var.set(sel)
    ttk.Button(top, text='Escolher...', command=choose_out).grid(row=1, column=2, sticky='w', pady=(6,0))

    ttk.Label(top, text='Índice inicial:').grid(row=2, column=0, sticky='w', pady=(6,0))
    ttk.Spinbox(top, from_=1, to=9999, textvariable=start_idx_var, width=6).grid(row=2, column=1, sticky='w', pady=(6,0))

    top.columnconfigure(1, weight=1)

    # Lista de arquivos
    mid = ttk.Frame(root, padding=(10, 0, 10, 10))
    mid.pack(fill='both', expand=True)
    ttk.Label(mid, text='Arquivos .lis encontrados:').pack(anchor='w')

    list_frame = ttk.Frame(mid)
    list_frame.pack(fill='both', expand=True)

    lb = tk.Listbox(list_frame, selectmode=tk.EXTENDED, height=14)
    lb.pack(side='left', fill='both', expand=True)
    sb = ttk.Scrollbar(list_frame, orient='vertical', command=lb.yview)
    sb.pack(side='right', fill='y')
    lb.configure(yscrollcommand=sb.set)

    btns = ttk.Frame(root, padding=(10, 0, 10, 10))
    btns.pack(fill='x')

    status_var = tk.StringVar(value='Pronto.')
    status = ttk.Label(root, textvariable=status_var, relief='sunken', anchor='w')
    status.pack(fill='x', padx=10, pady=(0,10))

    def refresh_list():
        lb.delete(0, tk.END)
        try:
            files = _scan_lis(Path(folder_var.get()))
        except Exception:
            files = []
        for f in files:
            lb.insert(tk.END, f.name)
        status_var.set(f"{len(files)} arquivo(s) encontrado(s).")

    def select_all():
        lb.select_set(0, tk.END)

    def clear_sel():
        lb.selection_clear(0, tk.END)

    def set_buttons_state(state: str):
        for w in (btn_refresh, btn_all, btn_clear, btn_proc, ent_folder, ent_out):
            try:
                w.configure(state=state)
            except Exception:
                pass

    def process_selected():
        files = _scan_lis(Path(folder_var.get()))
        sels = [lb.get(i) for i in lb.curselection()]
        if not sels:
            messagebox.showwarning('Aviso', 'Nenhum arquivo selecionado.')
            return
        paths = [next((p for p in files if p.name == name), None) for name in sels]
        paths = [p for p in paths if p is not None]
        if not paths:
            messagebox.showwarning('Aviso', 'Seleção inválida.')
            return

        outp = Path(outdir_var.get())
        start = int(start_idx_var.get()) if start_idx_var.get() else 1

        def worker():
            try:
                set_buttons_state('disabled')
                status_var.set('Processando...')
                outp.mkdir(parents=True, exist_ok=True)
                excel_paths = []
                idx = start
                for lp in paths:
                    status_var.set(f'Processando: {lp.name}')
                    # parse
                    df, stats_lines, summary_from_lis = parse_lis_table(lp)
                    if df is None:
                        continue
                    excel_path = outp / f"Resultados_Simulacao_{idx}.xlsx"
                    save_df_to_excel_only(df, excel_path)
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
                    criar_grafico_a_partir_do_excel(excel_path, outp, sim_index=idx, salvar_png=True, mostrar=False)
                    excel_paths.append(excel_path)
                    idx += 1
                if len(excel_paths) > 1:
                    status_var.set('Gerando gráfico comparativo...')
                    criar_grafico_comparativo(excel_paths, outp, mostrar=False)
                status_var.set('Concluído.')
                messagebox.showinfo('Concluído', f'Processo finalizado. Verifique: {outp}')
            except Exception:
                messagebox.showerror('Erro', traceback.format_exc())
            finally:
                set_buttons_state('normal')

        threading.Thread(target=worker, daemon=True).start()

    btn_refresh = ttk.Button(btns, text='Atualizar', command=refresh_list)
    btn_refresh.pack(side='left')
    btn_all = ttk.Button(btns, text='Selecionar tudo', command=select_all)
    btn_all.pack(side='left', padx=6)
    btn_clear = ttk.Button(btns, text='Limpar seleção', command=clear_sel)
    btn_clear.pack(side='left')
    btn_proc = ttk.Button(btns, text='Processar selecionados', command=process_selected)
    btn_proc.pack(side='right')

    refresh_list()
    root.mainloop()
