# GUIA DE APRESENTAÇÃO: `gui.py`

## Índice para Apresentação

1. [Visão Geral](#visão-geral)
2. [Conceitos Básicos](#conceitos-básicos)
3. [Estrutura do Arquivo](#estrutura-do-arquivo)
4. [Linha por Linha - Explicação Detalhada](#linha-por-linha)
5. [Recursos Visuais para Apresentação](#recursos-visuais)
6. [Demonstração Prática](#demonstração-prática)
7. [Perguntas Frequentes](#perguntas-frequentes)

---

## Visão Geral

### O que é o `gui.py`?

**Definição simples:** É o arquivo que cria a **janela gráfica** do programa. É o que o usuário vê e interage!
**Analogia:** 
- `main.py` = Motor de um carro (faz o trabalho pesado, mas você não vê)
- `gui.py` = Painel e volante (interface que você usa para controlar o carro)

### Por que precisamos dele?

Sem `gui.py`: O usuário precisaria digitar comandos no terminal
```bash
python main.py --arquivo "caso1.lis" --output "resultado.xlsx"
```

Com `gui.py`: O usuário clica em botões e seleciona arquivos
```
[Botão: Escolher Arquivo] → Clique → Janela de seleção
```

---

## Conceitos Básicos (para quem não sabe Python)

### 1. **Imports (Linhas 1-35)**

**O que é?** Importar = "pegar emprestado" código de outros lugares

```python
import tkinter as tk
```

**Explicação:**
- `tkinter` = Biblioteca que cria janelas gráficas
- `as tk` = Apelido (para não escrever "tkinter" toda hora)

**Analogia:**
```
Você vai fazer um bolo:
- import farinha     → Pega farinha do armário
- import ovos        → Pega ovos da geladeira
- import tkinter     → Pega ferramentas para criar janelas
```

### 2. **Classes (Linha 138+)**

**O que é?** Um "molde" para criar objetos

```python
class LisAnalysisApp:
    def __init__(self, root):
        # Código de inicialização
```

**Explicação:**
- `class` = Define o molde
- `LisAnalysisApp` = Nome do molde
- `def __init__` = Função que roda quando criamos a janela

**Analogia:**
```
class Bolo:                    class LisAnalysisApp:
    sabor = "chocolate"    →       cor = "azul"
    peso = "2kg"                   título = "LIS Analysis"
```

### 3. **Variáveis (Linhas 153-172)**

**O que é?** Uma "caixa" que guarda informações

```python
self.folder_var = tk.StringVar(value="C:/pasta")
```

**Explicação:**
- `self.folder_var` = Nome da caixa
- `StringVar` = Tipo especial de variável do tkinter
- `value=...` = Valor inicial

**Analogia:**
```
Caixa de Correio:
- Nome: "folder_var"
- Conteúdo: "C:/pasta"
- Tipo: StringVar (texto)
```

---

## Estrutura do Arquivo

### Organização Geral (1640 linhas!)

```
gui.py
│
├─ IMPORTS (1-35)           → "Ingredientes"
├─ FUNÇÕES AUXILIARES       → "Ferramentas auxiliares"
│  ├─ _scan_lis()          (42-60)
│  ├─ _scan_acp()          (63-81)
│  ├─ _fmt_size()          (84-90)
│  └─ _open_in_file_manager() (93-103)
│
├─ CLASSE Tooltip (106-134) → "Dicas que aparecem ao passar mouse"
│
└─ CLASSE LisAnalysisApp (138-1640) → "O PROGRAMA PRINCIPAL!"
   │
   ├─ __init__()           (139-183) → "Construtor - cria a janela"
   ├─ _load_prefs()        (186-205) → "Carrega configurações salvas"
   ├─ _build_menu()        (223-241) → "Cria o menu superior"
   ├─ _build_ui()          (243-575) → "Cria TODA a interface visual"
   │  ├─ Configurações
   │  ├─ Opções de processamento
   │  ├─ Variáveis do .lis
   │  ├─ Controle inteligente
   │  ├─ Filtros
   │  ├─ Botões de ação
   │  ├─ Lista de arquivos
   │  └─ Barra de progresso
   │
   ├─ _bind_shortcuts()    (577-583) → "Atalhos de teclado (Ctrl+S, etc)"
   ├─ _choose_folder()     (586-591) → "Escolher pasta de entrada"
   ├─ refresh_list()       (622-637) → "Atualizar lista de arquivos"
   ├─ _detect_variables()  (639-720) → "Detectar variáveis dos arquivos"
   ├─ _detect_control_parameters() (758-890) → "Detectar RPI, RF, etc"
   └─ process_selected()   (1500+)   → "Processar arquivos selecionados"
```

---

## Linha por Linha - Explicação Detalhada

### PARTE 1: Imports (Linhas 1-35)

#### Linhas 1-11: Bibliotecas Básicas

```python
import threading      # Para rodar coisas em paralelo
import traceback      # Para capturar erros detalhados
import json           # Para salvar/ler configurações
import sys            # Informações do sistema
import os             # Operações com pastas
import subprocess     # Executar programas externos
from datetime import datetime  # Data e hora
from pathlib import Path       # Trabalhar com caminhos de arquivo
import tkinter as tk           # CRIAR JANELAS!
from tkinter import ttk, filedialog, messagebox  # Widgets e diálogos
```

**Demonstração Visual:**
```
┌─────────────────────────────────────────┐
│ 📦 Caixa de Ferramentas                │
├─────────────────────────────────────────┤
│ 🧵 threading   → Multitarefa           │
│ 🐛 traceback   → Captura erros          │
│ 📄 json        → Salva configurações    │
│ 💻 sys/os      → Sistema operacional    │
│ 🪟 tkinter     → Cria janelas          │
└─────────────────────────────────────────┘
```

#### Linhas 14-35: Imports do Projeto

```python
try:
    from main import (
        parse_lis_table,           # Ler arquivos .lis
        save_df_to_excel_only,     # Salvar no Excel
        criar_grafico_...           # Criar gráficos
    )
    from acp_parser import (       # Trabalhar com arquivos .acp
        AcpParser,
        AtpRunner,
        modify_acp_rpi
    )
    from control_detector import ( # Detectar parâmetros RPI, RF
        ControlDetector,
        FileControlInfo
    )
except Exception:
    raise  # Se der erro, para tudo!
```

**Demonstração Visual:**
```
gui.py precisa de:
    ↓
    ├─ main.py         → Análise de dados
    ├─ acp_parser.py   → Modificar arquivos ATP
    └─ control_detector.py → Detectar parâmetros
```

---

### PARTE 2: Funções Auxiliares (Linhas 42-103)

#### Função `_scan_lis()` (Linhas 42-60)

```python
def _scan_lis(folder: Path):
    """Retorna arquivos .lis/.LIS ordenados por modificação (desc)."""
    folder = Path(folder)
    files = list(folder.glob('*.lis')) + list(folder.glob('*.LIS'))
    # ... código de duplicação ...
    files.sort(key=lambda f: f.stat().st_mtime, reverse=True)
    return files
```

**Explicação Simples:**
1. Recebe uma pasta
2. Procura todos os arquivos `.lis` (maiúscula ou minúscula)
3. Remove duplicatas
4. Ordena do mais recente para o mais antigo
5. Retorna a lista


**Demonstração Visual:**
```
Pasta "Listas ATP":
┌─────────────────────┐
│ 📄 caso1.lis       │ ← Encontra
│ 📄 CASO2.LIS       │ ← Encontra (maiúscula)
│ 📊 caso3.xlsx      │ ← Ignora (não é .lis)
│ 📄 caso4.lis       │ ← Encontra
└─────────────────────┘
        ↓
Ordena por data ↓
        ↓
Retorna: [caso4.lis, caso1.lis, CASO2.LIS]
```

#### Função `_fmt_size()` (Linhas 84-90)

```python
def _fmt_size(nbytes: int) -> str:
    for unit in ['B','KB','MB','GB','TB']:
        if nbytes < 1024:
            return f"{nbytes:.0f} {unit}"
        nbytes /= 1024
    return f"{nbytes:.0f} PB"
```

**Explicação Simples:**
Converte bytes em formato legível


**Demonstração:**
```python
_fmt_size(1500)      → "1500 B"   (Bytes)
_fmt_size(15000)     → "15 KB"    (Kilobytes)
_fmt_size(1500000)   → "1 MB"     (Megabytes)
_fmt_size(1500000000)→ "1 GB"     (Gigabytes)
```

---

### PARTE 3: Classe Tooltip (Linhas 106-134)

```python
class _Tooltip:
    def __init__(self, widget, text: str, delay_ms: int = 500):
        self.widget = widget
        self.text = text
        self.delay = delay_ms
        widget.bind('<Enter>', self._schedule)  # Mouse entra
        widget.bind('<Leave>', self._hide)      # Mouse sai
```

**O que faz?**
Cria aquelas "dicas" que aparecem quando você passa o mouse sobre um botão!


**Demonstração Visual:**
```
      [Botão: Salvar]
           ↓
    Mouse entra (espera 500ms)
           ↓
    ┌──────────────────────┐
    │ 💾 Salvar arquivo    │ ← Tooltip aparece!
    └──────────────────────┘
```

**Exemplo de Uso:**
```python
botao = ttk.Button(root, text="Salvar")
Tooltip(botao, "Salva o arquivo no disco")
```

---

### PARTE 4: Classe Principal `LisAnalysisApp` (Linha 138+)

#### `__init__()` - O Construtor (Linhas 139-183)

```python
class LisAnalysisApp:
    def __init__(self, root: tk.Tk, folder: Path, outdir: Path, start_index: int = 1):
        self.root = root
        self.root.title('LIS Analysis — Interface')
```

**O que acontece aqui?**

1. **Recebe parâmetros:**
   - `root` = Janela principal
   - `folder` = Pasta com arquivos .lis
   - `outdir` = Pasta para salvar resultados
   - `start_index` = Número inicial (padrão = 1)

2. **Configura a janela:**
   ```python
   self.root.title('LIS Analysis — Interface')  # Título
   ```

3. **Define o estilo visual:**
   ```python
   self.style = ttk.Style()
   self.style.theme_use('clam')  # Tema moderno
   ```

4. **Cria variáveis de controle:**
   ```python
   self.folder_var = tk.StringVar(value=str(folder))    # Pasta de entrada
   self.outdir_var = tk.StringVar(value=str(outdir))    # Pasta de saída
   self.status_var = tk.StringVar(value='Pronto.')      # Mensagem de status
   ```

**Demonstração Visual:**
```
Construtor __init__():
├─ 1. Cria janela principal        [✓]
├─ 2. Define título                [✓] "LIS Analysis"
├─ 3. Configura estilo/cores       [✓] Tema azul claro
├─ 4. Cria variáveis de controle   [✓] folder_var, status_var...
├─ 5. Carrega preferências salvas  [✓] _load_prefs()
├─ 6. Constrói menu superior       [✓] _build_menu()
├─ 7. Constrói interface completa  [✓] _build_ui()
├─ 8. Define atalhos de teclado    [✓] _bind_shortcuts()
└─ 9. Carrega lista de arquivos    [✓] refresh_list()
```

#### Variáveis de Controle (Linhas 153-172)

```python
# Checkboxes de opções (8 no total)
self.show_plots_var = tk.BooleanVar(value=False)      # Mostrar gráficos?
self.open_output_var = tk.BooleanVar(value=True)      # Abrir pasta?
self.only_comparative_var = tk.BooleanVar(value=False)# Só comparativo?
self.save_logs_var = tk.BooleanVar(value=True)        # Salvar logs?
```

**Demonstração Visual:**
```
Interface:                    Variável:
[✓] Mostrar gráficos    ←→   show_plots_var = True
[ ] Abrir pasta         ←→   open_output_var = False
[✓] Salvar logs         ←→   save_logs_var = True
```

---

### PARTE 5: Construção da Interface `_build_ui()` (Linhas 243-575)

**Esta é a função mais importante!** Ela cria TUDO que você vê na tela.

#### Estrutura Geral:

```python
def _build_ui(self):
    # 1. Canvas com scrollbar (linhas 244-283)
    # 2. Seção "Configurações" (linhas 285-310)
    # 3. Seção "Opções de Processamento" (linhas 312-371)
    # 4. Seção "Variáveis do .lis" (linhas 373-396)
    # 5. Seção "Controle Inteligente" (linhas 398-461)
    # 6. Filtros e tipo de arquivo (linhas 463-478)
    # 7. Botões de ação (linhas 480-498)
    # 8. Lista de arquivos (Treeview) (linhas 500-521)
    # 9. Barra de progresso (linhas 523-532)
```

#### 5.1: Canvas com Scrollbar (Linhas 244-283)

**Por que precisamos?**
A interface é muito grande! Precisamos de uma **barra de rolagem**.

```python
# Frame principal
main_frame = ttk.Frame(self.root)
main_frame.pack(fill='both', expand=True)

# Canvas (área rolável)
canvas = tk.Canvas(main_frame, highlightthickness=0)
canvas.pack(side='left', fill='both', expand=True)

# Scrollbar
scrollbar = ttk.Scrollbar(main_frame, orient='vertical', command=canvas.yview)
scrollbar.pack(side='right', fill='y')
```

**Demonstração Visual:**
```
┌──────────────────────┐
│ ╔════════════════╗ ▲ │ ← Scrollbar
│ ║ Configurações  ║ █ │
│ ║ Opções         ║ █ │
│ ║ Variáveis      ║ ▼ │
│ ║ Controle       ║   │
│ ║ ↓↓ (rola) ↓↓   ║   │
│ ╚════════════════╝   │
└──────────────────────┘
```

#### 5.2: Seção "Configurações" (Linhas 285-310)

```python
row1 = ttk.LabelFrame(container, text='⚙️ Configurações', padding=(10,8))
row1.pack(fill='x')

# Campo: Pasta (.lis)
ttk.Label(row1, text='Pasta (.lis):').grid(row=0, column=0, sticky='w')
self.ent_folder = ttk.Entry(row1, textvariable=self.folder_var)
self.ent_folder.grid(row=0, column=1, sticky='we', padx=6)
btn_folder = ttk.Button(row1, text='Escolher…', command=self._choose_folder)
btn_folder.grid(row=0, column=2, sticky='w')
```

**O que acontece:**
1. Cria um **quadro com título** "⚙️ Configurações"
2. Adiciona **3 linhas:**
   - Linha 1: Pasta de entrada (.lis)
   - Linha 2: Pasta de saída
   - Linha 3: Índice inicial

**Demonstração Visual:**
```
┌─────────────────────────────────────────┐
│ ⚙️ Configurações                        │
├─────────────────────────────────────────┤
│ Pasta (.lis): [C:/Listas ATP    ] [📁] │
│ Saída:        [C:/Resultados    ] [📁] │
│ Índice:       [1              ] ↑↓      │
└─────────────────────────────────────────┘
```

**Código simplificado para explicar:**
```python
# Criar rótulo (Label)
ttk.Label(row1, text='Pasta (.lis):')

# Criar campo de texto (Entry)
# Conectado à variável folder_var
self.ent_folder = ttk.Entry(row1, textvariable=self.folder_var)

# Criar botão
# Quando clica, executa _choose_folder()
btn_folder = ttk.Button(row1, text='Escolher…', command=self._choose_folder)
```

#### 5.3: Seção "Opções de Processamento" (Linhas 312-371)

```python
row1_5 = ttk.LabelFrame(container, text='⚙️ Opções de Processamento')
row1_5.pack(fill='x', pady=(8,0))

# Criar 4 colunas
chk_col1 = ttk.Frame(row1_5)
chk_col1.pack(side='left', fill='both', expand=True)
# ... chk_col2, chk_col3, chk_col4 ...

# Coluna 1
chk1 = ttk.Checkbutton(chk_col1, text='📊 Mostrar gráficos', 
                       variable=self.show_plots_var)
chk1.pack(anchor='w', pady=2)
```

**O que faz:**
Cria **8 checkboxes** organizados em **4 colunas**:


**Demonstração Visual:**
```
┌──────────────────────────────────────────────────────┐
│ ⚙️ Opções de Processamento                           │
├──────────────────────────────────────────────────────┤
│ Col 1          Col 2          Col 3         Col 4    │
│ [✓] Gráficos   [✓] Abrir     [ ] Só comp.  [✓] Logs │
│ [ ] Ocultar    [✓] Paralelo  [✓] Organizar [✓] Sobre│
└──────────────────────────────────────────────────────┘
```

**Relação Checkbox ↔ Variável:**
```python
Checkbox visual             Variável Python
[✓] Mostrar gráficos  ←→   show_plots_var = True
[ ] Abrir pasta       ←→   open_output_var = False
[✓] Salvar logs       ←→   save_logs_var = True
```

#### 5.4: Seção "Variáveis do .lis" (Linhas 373-396)

**NOVA FUNCIONALIDADE!** Detecta automaticamente as variáveis do arquivo.

```python
row1_7 = ttk.LabelFrame(container, text='📊 Variáveis do Arquivo .lis')
row1_7.pack(fill='x', pady=(8,0))

# Frame para checkboxes dinâmicos
self.variables_frame = ttk.Frame(row1_7)
self.variables_frame.pack(fill='both', expand=True)

# Mensagem inicial
self.variables_label = ttk.Label(
    self.variables_frame, 
    text='💡 Selecione um arquivo .lis para detectar variáveis',
    foreground='gray'
)
self.variables_label.pack(pady=10)

# Botão para detectar
btn_detect = ttk.Button(row1_7, text='🔍 Detectar Variáveis', 
                        command=self._detect_variables)
btn_detect.pack(pady=(5,0))
```

**Como funciona:**

1. **Inicialmente:** Mostra apenas mensagem "Selecione um arquivo..."
2. **Usuário clica "Detectar Variáveis":**
   - Sistema lê o arquivo .lis
   - Encontra todas as variáveis (ex: V_A, V_B, I_LOAD)
   - Cria checkboxes automaticamente
3. **Usuário escolhe** quais variáveis analisar


**Demonstração Visual (ANTES):**
```
┌─────────────────────────────────────────┐
│ 📊 Variáveis do Arquivo .lis            │
├─────────────────────────────────────────┤
│ 💡 Selecione um arquivo para detectar   │
│    [🔍 Detectar Variáveis]              │
└─────────────────────────────────────────┘
```

**Demonstração Visual (DEPOIS de detectar):**
```
┌─────────────────────────────────────────┐
│ 📊 Variáveis do Arquivo .lis            │
├─────────────────────────────────────────┤
│ ✅ 6 variáveis detectadas:              │
│                                         │
│ [✓] 📌 V_A    [✓] 📌 V_B               │
│ [✓] 📌 V_C    [✓] 📌 I_LOAD            │
│ [✓] 📌 P_ACT  [ ] 📌 P_REACT           │
│                                         │
│ [✓ Selecionar Todas] [✗ Desmarcar]     │
└─────────────────────────────────────────┘
```

#### 5.5: Seção "Controle Inteligente" (Linhas 398-461)

**A FUNCIONALIDADE MAIS AVANÇADA!**

```python
row1_8 = ttk.LabelFrame(container, text='🎯 Controle Inteligente de Parâmetros')
row1_8.pack(fill='x', pady=(8,0))

# Frame dinâmico
self.control_frame = ttk.Frame(row1_8)
self.control_frame.pack(fill='both', expand=True)

# Mensagem inicial
self.control_label = ttk.Label(
    self.control_frame,
    text='💡 Selecione arquivos .lis/.acp para detectar parâmetros (RPI, RF, etc)',
    foreground='gray'
)
self.control_label.pack(pady=10)

# Botões
self.btn_detect_controls = ttk.Button(
    control_buttons, 
    text='🔍 Detectar Parâmetros', 
    command=self._detect_control_parameters
)
```

**Como funciona:**

1. **Detecta automaticamente** do nome do arquivo:
   ```
   Arquivo: "Caso0_Convenc_RPI=100 e RF=30.lis"
   Detecta: RPI=100Ω, RF=30Ω
   ```

2. **Cria controles dinâmicos:**
   ```python
   RPI: [150]Ω  [100][200][300] ← Botões de sugestão
   RF:  [30]Ω   [10][20][30]
   ```

3. **Permite modificar** e gerar novos arquivos


**Demonstração Visual (Fluxo completo):**
```
PASSO 1: Selecionar arquivo
┌─────────────────────────────────────┐
│ [✓] Caso0_RPI=100_RF=30.lis        │ ← Usuário seleciona
└─────────────────────────────────────┘

PASSO 2: Clicar "Detectar Parâmetros"
         ↓
Sistema analisa nome do arquivo
         ↓
Detecta: RPI=100, RF=30
         ↓

PASSO 3: Interface atualiza automaticamente
┌─────────────────────────────────────────────┐
│ 🎯 Controle Inteligente                     │
├─────────────────────────────────────────────┤
│ ✅ 1 arquivo com controle detectado         │
│                                             │
│ RPI (Resistência Pré-Inserção):            │
│ 📌 Detectado: 100Ω                          │
│    Novo valor: [150]Ω [100][200][300][400] │
│                                             │
│ RF (Resistor de Falta):                     │
│ 📌 Detectado: 30Ω                           │
│    Novo valor: [30]Ω [10][20][30][40]      │
│                                             │
│ [🔧 Aplicar] [📊 Resumo]                   │
└─────────────────────────────────────────────┘

PASSO 4: Usuário modifica RPI para 150
         ↓
PASSO 5: Clica "Aplicar"
         ↓
Sistema cria: "Caso0_RPI=150_RF=30.acp"
```

#### 5.6: Lista de Arquivos (Treeview) (Linhas 500-521)

```python
columns = ('nome', 'tamanho', 'modificado')
self.tv = ttk.Treeview(row3, columns=columns, show='headings', 
                       selectmode='extended')

self.tv.heading('nome', text='Nome', command=lambda: self._sort_by('nome'))
self.tv.heading('tamanho', text='Tamanho')
self.tv.heading('modificado', text='Modificado')

self.tv.column('nome', anchor='w', width=420)
self.tv.column('tamanho', anchor='center', width=120)
self.tv.column('modificado', anchor='center', width=180)
```

**O que é Treeview?**
Uma tabela onde você pode selecionar múltiplos arquivos.

**Demonstração Visual:**
```
┌───────────────────────────────────────────────────────────┐
│ 📋 Arquivos encontrados                                   │
├─────────────────────┬──────────┬───────────────────────────┤
│ Nome ↓              │ Tamanho  │ Modificado                │
├─────────────────────┼──────────┼───────────────────────────┤
│ Caso0_RPI=100.lis   │ 2.3 MB   │ 2025-11-16 10:30         │
│ Caso0_RPI=200.lis   │ 2.4 MB   │ 2025-11-16 10:25         │
│ Caso0_RPI=300.lis   │ 2.2 MB   │ 2025-11-16 10:20         │
└─────────────────────┴──────────┴───────────────────────────┘
        ↑                              ↑
   Clica para ordenar          Selecionar múltiplos com Ctrl
```

**Funcionalidades:**
- **Ordenação:** Clique no título para ordenar
- **Seleção múltipla:** Ctrl+Click
- **Seleção tudo:** Ctrl+A
- **Scroll:** Se tiver muitos arquivos

---

### PARTE 6: Métodos de Ação

#### `_choose_folder()` (Linhas 586-591)

```python
def _choose_folder(self):
    sel = filedialog.askdirectory(initialdir=self.folder_var.get() or '.')
    if sel:
        self.folder_var.set(sel)
        self._save_prefs()
        self.refresh_list()
```

**O que faz:**
1. Abre diálogo para escolher pasta
2. Se escolher, atualiza `folder_var`
3. Salva preferência
4. Atualiza lista de arquivos

**Demonstração Visual:**
```
Usuário clica: [Escolher pasta]
        ↓
   _choose_folder() executa
        ↓
Abre janela do sistema:
┌────────────────────────────────┐
│ Selecionar Pasta               │
│ ┌────────────────────────────┐ │
│ │ 📁 Documentos              │ │
│ │ 📁 Downloads               │ │
│ │ 📁 Listas ATP         ← ✓  │ │
│ │ 📁 Resultados              │ │
│ └────────────────────────────┘ │
│ [Cancelar]  [Selecionar]       │
└────────────────────────────────┘
        ↓
folder_var = "C:/Listas ATP"
        ↓
Salva preferência (JSON)
        ↓
Atualiza lista de arquivos
```

#### `refresh_list()` (Linhas 622-637)

```python
def refresh_list(self):
    folder = Path(self.folder_var.get()).expanduser()
    try:
        ftype = (self.filetype_var.get() or '.lis').strip().lower()
        if ftype == '.acp':
            self._files_cache = _scan_acp(folder)
        elif ftype == 'ambos':
            self._files_cache = _scan_lis(folder) + _scan_acp(folder)
        else:
            self._files_cache = _scan_lis(folder)
    except Exception:
        self._files_cache = []
    self._populate_tree()
    self.status_var.set(f"{len(self._files_cache)} arquivo(s) encontrado(s)")
```

**O que faz:**
1. Lê a pasta escolhida
2. Procura arquivos do tipo selecionado (.lis, .acp ou ambos)
3. Armazena na cache
4. Atualiza a tabela visual


**Demonstração (Passo a Passo):**
```
1. Pega pasta: folder_var.get() → "C:/Listas ATP"

2. Verifica tipo selecionado:
   filetype_var = ".lis"

3. Chama função auxiliar:
   _scan_lis("C:/Listas ATP")
   
4. Recebe lista:
   ["caso1.lis", "caso2.lis", "caso3.lis"]

5. Atualiza interface:
   Treeview mostra os 3 arquivos

6. Atualiza status:
   "3 arquivo(s) encontrado(s)"
```

#### `_detect_variables()` (Linhas 639-720)

**MÉTODO IMPORTANTE!** Detecta variáveis automaticamente.

```python
def _detect_variables(self):
    from main import parse_lis_output_variables
    
    # Pegar arquivo selecionado
    sels = self.tv.selection()
    if sels:
        lis_path = Path(sels[0])
    else:
        messagebox.showwarning('Aviso', 'Selecione um arquivo!')
        return
    
    # Detectar variáveis
    variables = parse_lis_output_variables(lis_path)
    
    # Criar checkboxes dinamicamente
    for var in variables:
        var_checkbox = tk.BooleanVar(value=True)
        self.variable_checkboxes[var] = var_checkbox
        chk = ttk.Checkbutton(frame, text=f'📌 {var}', variable=var_checkbox)
        chk.pack()
```

**Fluxo Completo:**


**Demonstração :**
```
ETAPA 1: Usuário seleciona arquivo
[✓] Caso0_RPI=100.lis

ETAPA 2: Clica "🔍 Detectar Variáveis"
         ↓
_detect_variables() executa

ETAPA 3: Lê o arquivo .lis
┌─────────────────────────────┐
│ Arquivo: Caso0_RPI=100.lis │
│ ...                         │
│ Column headings:            │
│ V_A  V_B  V_C  I_LOAD P_ACT│ ← Encontra variáveis!
│ ...                         │
└─────────────────────────────┘

ETAPA 4: Cria lista
variables = ["V_A", "V_B", "V_C", "I_LOAD", "P_ACT"]

ETAPA 5: Cria checkboxes (LOOP)
Para cada variável em variables:
    1. Cria BooleanVar (True)
    2. Cria Checkbutton
    3. Adiciona à interface

ETAPA 6: Interface atualiza!
┌─────────────────────────────┐
│ ✅ 5 variáveis detectadas   │
│ [✓] 📌 V_A                  │
│ [✓] 📌 V_B                  │
│ [✓] 📌 V_C                  │
│ [✓] 📌 I_LOAD               │
│ [✓] 📌 P_ACT                │
└─────────────────────────────┘
```

#### `_detect_control_parameters()` (Linhas 758-890)

**O MÉTODO MAIS COMPLEXO E IMPORTANTE!**

```python
def _detect_control_parameters(self):
    # 1. Pegar arquivos selecionados
    sels = self.tv.selection()
    selected_files = [Path(item) for item in sels]
    
    # 2. Detectar parâmetros usando ControlDetector
    self.detected_controls = ControlDetector.detect_from_files(selected_files)
    
    # 3. Filtrar arquivos COM controle
    with_control = [info for info in self.detected_controls if info.has_control]
    
    # 4. Coletar parâmetros únicos
    all_params = {}  # {param_name: [values]}
    for info in with_control:
        for param in info.parameters:
            if param.name not in all_params:
                all_params[param.name] = set()
            all_params[param.name].add(param.value)
    
    # 5. Criar controles dinâmicos
    for param_name in sorted(all_params.keys()):
        # Criar label, spinbox, botões de sugestão...
```

**Fluxo Completo (Simplificado):**

```
ENTRADA:
Arquivos selecionados:
- Caso0_RPI=100_RF=30.lis
- Caso0_RPI=200_RF=30.lis

    ↓

PROCESSAMENTO:
1. ControlDetector analisa nomes
2. Detecta: RPI=[100, 200], RF=[30]
3. Gera sugestões: RPI=[50,100,150,200,250...]

    ↓

SAÍDA (Interface):
┌─────────────────────────────────────┐
│ RPI (Resistência Pré-Inserção):    │
│ 📌 Detectado: 100Ω, 200Ω           │
│ Novo: [100]Ω [50][100][150][200]   │
│                                     │
│ RF (Resistor de Falta):             │
│ 📌 Detectado: 30Ω                   │
│ Novo: [30]Ω [10][20][30][40]       │
└─────────────────────────────────────┘
```


---

## Recursos Visuais 

### 1. Fluxograma Geral

```
┌─────────────┐
│ Usuário     │
└──────┬──────┘
       │ Abre programa
       ↓
┌─────────────────────────────┐
│ main.py                     │
│ └→ Importa gui.py           │
│ └→ Cria janela tkinter      │
│ └→ Instancia LisAnalysisApp │
└──────────┬──────────────────┘
           │
           ↓
┌──────────────────────────────────────┐
│ LisAnalysisApp.__init__()            │
│ ├─ Cria variáveis de controle        │
│ ├─ Carrega preferências (_load_prefs)│
│ ├─ Constrói menu (_build_menu)       │
│ ├─ Constrói UI (_build_ui) ★★★★★    │
│ └─ Atualiza lista (refresh_list)     │
└──────────┬───────────────────────────┘
           │
           ↓
┌──────────────────────────┐
│ Interface exibida!       │
│ Usuário interage...      │
└──────────────────────────┘
```

### 2. Mapa de Interações

```
INTERFACE                          MÉTODO CHAMADO
┌────────────────┐
│ [Escolher...]  │ ──────────→   _choose_folder()
└────────────────┘
┌────────────────┐
│ [Detectar Var] │ ──────────→   _detect_variables()
└────────────────┘
┌────────────────┐
│ [Detectar Par] │ ──────────→   _detect_control_parameters()
└────────────────┘
┌────────────────┐
│ [Aplicar Mod]  │ ──────────→   _apply_control_modifications()
└────────────────┘
┌────────────────┐
│ [Processar]    │ ──────────→   process_selected()
└────────────────┘
```

### 3. Arquitetura MVC (Simplificada)

```
┌──────────────────────────────────────┐
│           MODEL (Dados)              │
│  • folder_var (pasta de entrada)    │
│  • status_var (mensagens)            │
│  • detected_controls (parâmetros)    │
└─────────────┬────────────────────────┘
              │ Atualiza
              ↓
┌──────────────────────────────────────┐
│          VIEW (Visual)               │
│  • Botões, Labels, Entries           │
│  • Treeview (lista de arquivos)      │
│  • LabelFrames (seções)              │
└─────────────┬────────────────────────┘
              │ Eventos (cliques)
              ↓
┌──────────────────────────────────────┐
│       CONTROLLER (Lógica)            │
│  • _choose_folder()                  │
│  • _detect_variables()               │
│  • _detect_control_parameters()      │
│  • process_selected()                │
└──────────────────────────────────────┘
```

---

## Demonstração Prática (Script)

### Demonstração 1: "Do Zero ao Olá Mundo"

**Arquivo: `demo1_basico.py`**

```python
import tkinter as tk
from tkinter import ttk

# ETAPA 1: Criar janela
root = tk.Tk()
root.title("Minha Primeira Janela")
root.geometry("400x200")

# ETAPA 2: Adicionar label
label = ttk.Label(root, text="Olá Mundo!", font=("Arial", 20))
label.pack(pady=50)

# ETAPA 3: Adicionar botão
def ao_clicar():
    label.config(text="Você clicou!")

botao = ttk.Button(root, text="Clique Aqui", command=ao_clicar)
botao.pack()

# ETAPA 4: Iniciar loop
root.mainloop()
```

**Execução:**
```bash
python demo1_basico.py
```

### Demonstração 2: "Variáveis de Controle"

**Arquivo: `demo2_variaveis.py`**

```python
import tkinter as tk
from tkinter import ttk

root = tk.Tk()
root.title("Variáveis de Controle")

# CRIAR VARIÁVEL
nome_var = tk.StringVar(value="João")
idade_var = tk.IntVar(value=20)
ativo_var = tk.BooleanVar(value=True)

# CONECTAR À INTERFACE
ttk.Label(root, text="Nome:").pack()
ttk.Entry(root, textvariable=nome_var).pack()

ttk.Label(root, text="Idade:").pack()
ttk.Spinbox(root, from_=1, to=100, textvariable=idade_var).pack()

ttk.Checkbutton(root, text="Ativo?", variable=ativo_var).pack()

# MOSTRAR VALORES
def mostrar():
    msg = f"Nome: {nome_var.get()}\n"
    msg += f"Idade: {idade_var.get()}\n"
    msg += f"Ativo: {ativo_var.get()}"
    print(msg)

ttk.Button(root, text="Mostrar Valores", command=mostrar).pack(pady=20)

root.mainloop()
```

**Explicação:**
> "Variáveis de controle são a ponte entre a interface e os dados do programa. Quando você digita no campo, a variável é atualizada automaticamente!"

### Demonstração 3: "Construção Dinâmica"

**Arquivo: `demo3_dinamico.py`**

```python
import tkinter as tk
from tkinter import ttk

root = tk.Tk()
root.title("Interface Dinâmica")

frame = ttk.Frame(root, padding=20)
frame.pack()

# SIMULAR DETECÇÃO DE VARIÁVEIS
def detectar():
    # Limpar frame
    for widget in frame.winfo_children():
        widget.destroy()
    
    # "Detectar" variáveis (simulado)
    variaveis = ["V_A", "V_B", "V_C", "I_LOAD"]
    
    ttk.Label(frame, text=f"✅ {len(variaveis)} variáveis detectadas:",
              font=("Arial", 12, "bold")).pack(anchor='w')
    
    # Criar checkboxes dinamicamente
    for var in variaveis:
        chk = ttk.Checkbutton(frame, text=f"📌 {var}")
        chk.pack(anchor='w', pady=2)

ttk.Button(root, text="🔍 Detectar Variáveis", 
           command=detectar).pack(pady=10)

root.mainloop()
```

**Explicação:**
> "Este exemplo mostra como criar elementos da interface DINAMICAMENTE, igual ao sistema de detecção de variáveis do LIS-ANALYSIS!"

---

## Perguntas Frequentes (para preparar respostas)

### P1: "Por que o arquivo tem 1640 linhas?"

**R:** O `gui.py` é responsável por TODA a interface visual. Cada botão, cada campo, cada seção precisa ser criada e configurada. É como construir uma casa: precisa de muitos tijolos!

**Comparação:**
- Casa simples (100 linhas) = Janela com 2 botões
- Casa grande (1640 linhas) = Interface completa com menu, tabelas, checkboxes, detecção automática, etc.

### P2: "O que é `tk` e `ttk`?"

**R:**
- `tk` = Biblioteca básica do Python para criar interfaces gráficas (GUI)
- `ttk` = Versão **moderna** dos widgets do tk (visual mais bonito)

```python
# Antigo (tk)
botao = tk.Button(root, text="Clique")  # Feio

# Moderno (ttk)
botao = ttk.Button(root, text="Clique")  # Bonito!
```

### P3: "Por que usar `self.` em tudo?"

**R:** `self.` significa "desta instância". É como dizer "MEU botão", "MINHA variável".

```python
class Carro:
    def __init__(self):
        self.cor = "vermelho"  # MEU carro é vermelho
        self.marca = "Fiat"    # MEU carro é Fiat

carro1 = Carro()
carro1.cor = "azul"  # Agora MEU carro (carro1) é azul

carro2 = Carro()
carro2.cor = "verde"  # Outro carro (carro2) é verde

# carro1 e carro2 são diferentes!
```

### P4: "O que são `lambda` functions?"

**R:** São funções **anônimas** de uma linha só. Atalho para funções simples.

```python
# Sem lambda (forma normal)
def ao_clicar():
    self.refresh_list()

botao = ttk.Button(root, text="Atualizar", command=ao_clicar)

# Com lambda (atalho)
botao = ttk.Button(root, text="Atualizar", 
                   command=lambda: self.refresh_list())
```

### P5: "Como funciona o `pack()` e `grid()`?"

**R:** São gerenciadores de layout (como os widgets são organizados).

**pack()** = Empilha elementos (um embaixo do outro)
```python
label1.pack()   # Topo
label2.pack()   # Abaixo de label1
button.pack()   # Abaixo de label2
```

**grid()** = Organiza em grade (linhas e colunas)
```python
label.grid(row=0, column=0)   # Linha 0, Coluna 0
entry.grid(row=0, column=1)   # Linha 0, Coluna 1
button.grid(row=1, column=0)  # Linha 1, Coluna 0
```

**Visual:**
```
pack():              grid():
┌──────┐            ┌────┬─────┐
│ Label│            │Lab │Entry│ ← Linha 0
├──────┤            ├────┴─────┤
│ Label│            │Button    │ ← Linha 1
├──────┤            └──────────┘
│Button│              ↑     ↑
└──────┘            Col 0  Col 1
```

---


