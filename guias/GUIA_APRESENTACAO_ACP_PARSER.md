# GUIA DE APRESENTAÇÃO - `acp_parser.py` (ainda em desenvolvimento)

## Objetivo Deste Guia

Este documento serve como **roteiro de apresentação** para explicar o arquivo `acp_parser.py` para pessoas que **não conhecem Python profundamente**, focando em como o código **modifica** e **executa** simulações ATP automaticamente.

---

## ÍNDICE

1. [Visão Geral](#1-visão-geral)
2. [O Que Este Arquivo Faz](#2-o-que-este-arquivo-faz)
3. [Conceitos Fundamentais](#3-conceitos-fundamentais)
4. [Estrutura do Código](#4-estrutura-do-código)
5. [Explicação Detalhada](#5-explicação-detalhada)
6. [Demonstração Prática](#6-demonstração-prática)
7. [Pontos-Chave](#7-pontos-chave-para-apresentação)

---

## 1. VISÃO GERAL

### **O que é este arquivo?**

```
acp_parser.py = "Cirurgião de Arquivos ATP" 🔧
```

Este arquivo é um **modificador cirúrgico** que:

- **Abre arquivos `.acp`** (formato compactado do ATPDraw)
- **Extrai** o conteúdo ATP interno
- **Modifica** valores de parâmetros (RPI, RF, etc)
- **Salva** novo arquivo modificado
- **Executa** simulações ATP automaticamente
- **Coleta** resultados `.lis` gerados

### **Por que existe?**

**Problema do Mundo Real:**

```
Você tem arquivo: Caso0_RPI=100.acp
Precisa criar:     Caso0_RPI=150.acp
                   Caso0_RPI=200.acp
                   Caso0_RPI=250.acp
                   ... (mais 10 variações)

SEM este código:
├─ Abrir ATPDraw manualmente
├─ Modificar valor de RPI (encontrar na interface)
├─ Salvar com novo nome
├─ Repetir 10 vezes
└─ Tempo: ~30 minutos

COM este código:
├─ Loop modificando RPI
├─ Salvar automaticamente
└─ Tempo: ~5 segundos
```

**Solução:**

```python
from acp_parser import modify_acp_rpi

# Uma linha de código faz TUDO!
modify_acp_rpi(
    "Caso0_RPI=100.acp", 
    new_rpi=150, 
    output="Caso0_RPI=150.acp"
)
```

### **Analogia do Mundo Real**

Imagine que você tem um **documento Word compactado em ZIP**:

```
documento.docx (é um arquivo ZIP!)
├─ content.xml         ← Texto do documento
├─ styles.xml          ← Formatação
└─ media/imagem.png    ← Imagens
```

Este código faz o mesmo com arquivos `.acp`:

```python
arquivo.acp (é um arquivo ZIP!)
├─ arquivo.$$$    ← Código ATP (texto)
├─ arquivo.dwg    ← Desenho gráfico
└─ metadados.xml  ← Configurações
        ↓
   Abrir ZIP
        ↓
   Modificar arquivo.$$$
        ↓
   Recompactar ZIP
```

---

## 2. O QUE ESTE ARQUIVO FAZ

### **Fluxo Completo**

```
┌──────────────────────────────────────────────────┐
│  ENTRADA: Arquivo .acp                           │
│  "Caso0_Convenc_RPI=100.acp"                    │
└────────────┬─────────────────────────────────────┘
             │
             ↓
┌──────────────────────────────────────────────────┐
│  ETAPA 1: Extrair Conteúdo                      │
│  ├─ Abrir como arquivo ZIP                      │
│  ├─ Encontrar arquivo .$$$                      │
│  └─ Ler texto ATP                                │
└────────────┬─────────────────────────────────────┘
             │
             ↓
┌──────────────────────────────────────────────────┐
│  ETAPA 2: Buscar Parâmetros                     │
│  ├─ Procurar linhas com "RPI"                   │
│  ├─ Encontrar valores numéricos                 │
│  └─ Mapear posições no arquivo                  │
└────────────┬─────────────────────────────────────┘
             │
             ↓
┌──────────────────────────────────────────────────┐
│  ETAPA 3: Modificar Valores                     │
│  ├─ Substituir RPI=100 por RPI=150              │
│  ├─ Manter formatação original                  │
│  └─ Preservar outros parâmetros                 │
└────────────┬─────────────────────────────────────┘
             │
             ↓
┌──────────────────────────────────────────────────┐
│  ETAPA 4: Salvar Modificado                     │
│  ├─ Copiar estrutura original                   │
│  ├─ Substituir arquivo .$$$                     │
│  └─ Recompactar como .acp                       │
└────────────┬─────────────────────────────────────┘
             │
             ↓
┌──────────────────────────────────────────────────┐
│  ETAPA 5: Executar ATP (opcional)               │
│  ├─ Chamar executável tpbig/atpmingw            │
│  ├─ Aguardar simulação                          │
│  └─ Coletar arquivo .lis gerado                 │
└────────────┬─────────────────────────────────────┘
             │
             ↓
┌──────────────────────────────────────────────────┐
│  SAÍDA: Novo arquivo .acp + .lis                │
│  "Caso0_Convenc_RPI=150.acp"                    │
│  "Caso0_Convenc_RPI=150.lis" (após simulação)   │
└──────────────────────────────────────────────────┘
```

### **Operações Principais**

| Operação | Descrição | Tempo |
|----------|-----------|-------|
| **Extrair** | Descompactar .acp e ler ATP | ~0.1s |
| **Buscar** | Encontrar parâmetros no texto | ~0.05s |
| **Modificar** | Alterar valores | ~0.02s |
| **Salvar** | Recompactar arquivo | ~0.2s |
| **Executar** | Rodar simulação ATP | ~5-30s |

**Total sem simulação:** ~0.4 segundos  
**Total com simulação:** ~5-30 segundos

---

## 3. CONCEITOS FUNDAMENTAIS

### **3.1 O Que É um Arquivo .acp?**

**.acp = Arquivo de Projeto do ATPDraw**

```
┌─────────────────────────────────────────┐
│ Arquivo .acp (formato ZIP!)             │
├─────────────────────────────────────────┤
│                                         │
│    projeto.$$$                          │
│    └─ Código ATP em texto               │
│       (instruções da simulação)         │
│                                         │
│    projeto.dwg                          │
│    └─ Desenho do circuito               │
│       (representação visual)            │
│                                         │
│    projeto.xml                          │
│    └─ Metadados e configurações         │
│                                         │
│   resources/                            │
│    └─ Bibliotecas e componentes         │
└─────────────────────────────────────────┘
```

**Prova:**

```bash
# Renomeie .acp para .zip e abra!
cp arquivo.acp arquivo.zip
unzip arquivo.zip
# Você verá todos os arquivos internos!
```

### **3.2 Estrutura do Arquivo ATP (.$$$)**

**Arquivo ATP = "Receita de Simulação"**

```
BEGIN NEW DATA CASE
C -------------------------------------
C Comentário: Simulação de Re-energização
C -------------------------------------
C <dt><tmax><xopt><copt>
  5.0E-05    0.500   60.    60.
C
C Chaves e resistores de pré-inserção
$VINTAGE, 1,
$ERASE
C RPI1  (Resistência de Pré-Inserção = 100 Ohms)
  BUSA  BUSRPI1         100.000              1
C RPI2  
  BUSB  BUSRPI2         100.000              1
C Switch de fechamento
C <busa><busb>  <tclose><topen><ie>
$SWITCH
  BUSA  BUSRPI1     0.100  0.500  1
C ... mais configurações ...
BLANK card terminating branches
BEGIN NEW DATA CASE
BLANK
```

**Partes Importantes:**

1. **dt, tmax**: Tempo de simulação
2. **Comentários (C)**: Descrições
3. **Resistores**: Valores de RPI, RF, etc
4. **Switches**: Tempos de chaveamento

### **3.3 Encoding de Caracteres**

**Por que `windows-1252`?**

ATP foi desenvolvido nos anos 80/90 quando Windows usava encoding diferente de UTF-8:

```python
# Tentar ler com UTF-8 → Pode dar erro!
text = file.read().decode('utf-8')  

# Correto: usar encoding do Windows
text = file.read().decode('windows-1252')  
```

**Analogia:** É como traduzir um livro antigo que usa alfabeto diferente do atual.

### **3.4 Subprocess - Executar Programas Externos**

**subprocess = "Chamar outro programa via Python"**

```python
import subprocess

# Python chama ATP para executar simulação
subprocess.run(['tpbig', 'arquivo.atp'])

# É como digitar no terminal:
# $ tpbig arquivo.atp
```

**Fluxo Visual:**

```
Python (seu código)
    ↓ subprocess.run()
Terminal/CMD
    ↓ executa
ATP (tpbig.exe)
    ↓ processa
Arquivo .lis (resultado)
    ↓ retorna
Python (coleta resultado)
```

---

## 4. ESTRUTURA DO CÓDIGO

### **Organização do Arquivo**

```python
acp_parser.py (896 linhas)
│
├─ BLOCO 1: Imports (linhas 1-15)
│  └─ zipfile, subprocess, pathlib, etc
│
├─ BLOCO 2: Classe AcpParser (linhas 17-280)
│  │
│  ├─ __init__()                   → Inicialização
│  ├─ extract_atp_from_acp()       → Extrair ATP do ZIP
│  ├─ find_control_parameters()    → Buscar RPI, RF, etc
│  ├─ modify_rpi_value()           → Modificar valores
│  ├─ save_modified_acp()          → Salvar modificado
│  └─ print_summary()              → Resumo visual
│
├─ BLOCO 3: Classe AtpRunner (linhas 282-450)
│  │
│  ├─ __init__()                   → Configurar executável
│  ├─ _find_atp_executable()       → Auto-detectar ATP
│  ├─ run_simulation()             → Executar ATP
│  └─ collect_results()            → Coletar .lis
│
├─ BLOCO 4: Funções Auxiliares (linhas 452-550)
│  ├─ modify_acp_rpi()     → Atalho para modificar RPI
│  ├─ run_acp_simulation() → Atalho para simular
│  └─ batch_process_acp()  → Processar múltiplos
│
└─ BLOCO 5: Testes e Demos (linhas 552-896)
   └─ if __name__ == "__main__": → Exemplos de uso
```

### **Hierarquia de Classes**

```
AcpParser (manipulação de arquivos)
    ├─ Responsabilidades:
    │  ├─ Abrir/fechar .acp
    │  ├─ Extrair ATP
    │  ├─ Modificar conteúdo
    │  └─ Salvar modificações
    │
    └─ Métodos principais:
       ├─ extract_atp_from_acp()
       ├─ modify_rpi_value()
       └─ save_modified_acp()

AtpRunner (execução de simulações)
    ├─ Responsabilidades:
    │  ├─ Encontrar executável ATP
    │  ├─ Executar simulação
    │  └─ Coletar resultados
    │
    └─ Métodos principais:
       ├─ _find_atp_executable()
       ├─ run_simulation()
       └─ collect_results()
```

---

## 5. EXPLICAÇÃO DETALHADA

### **5.1 Extrair ATP do .acp: `extract_atp_from_acp()`**

**Localização:** Linhas 23-62

**O Que Faz:** Abre arquivo .acp como ZIP e extrai o código ATP.

```python
def extract_atp_from_acp(self) -> Optional[str]:
    """
    ENTRADA: Path("Caso0_RPI=100.acp")
    SAÍDA:   String com conteúdo ATP (texto completo)
    """
```

**Fluxo Interno:**

```
1. Verificar se arquivo existe
   ↓
2. Abrir como ZIP
   with zipfile.ZipFile(acp_path, 'r') as zip_ref:
   ↓
3. Listar arquivos internos
   files = zip_ref.namelist()
   # ['arquivo.$$$', 'arquivo.dwg', 'arquivo.xml']
   ↓
4. Encontrar arquivo que termina com .$$$
   atp_file = [f for f in files if f.endswith('.$$$')][0]
   ↓
5. Ler conteúdo
   content = zip_ref.read(atp_file)
   ↓
6. Decodificar para texto
   text = content.decode('windows-1252')
   ↓
7. Armazenar e retornar
   self.atp_text = text
   return text
```

**Exemplo Passo a Passo:**

```python
# Arquivo: Caso0_RPI=100.acp

# PASSO 1: Abrir como ZIP
zip_ref = zipfile.ZipFile("Caso0_RPI=100.acp", 'r')

# PASSO 2: Listar conteúdo
zip_ref.namelist()
# Resultado: ['caso0.$$$', 'caso0.dwg', 'caso0.xml']

# PASSO 3: Ler arquivo ATP
atp_bytes = zip_ref.read('caso0.$$$')
# Resultado: b'BEGIN NEW DATA CASE\nC ...'

# PASSO 4: Decodificar
atp_text = atp_bytes.decode('windows-1252')
# Resultado: "BEGIN NEW DATA CASE\nC ..."

# PASSO 5: Pronto!
print(atp_text[:100])
# "BEGIN NEW DATA CASE
#  C Simulação de Re-energização
#  C RPI = 100 Ohms
#  ..."
```

**Por Que Funciona:**

- `.acp` **é realmente um ZIP** (pode renomear para `.zip` e abrir!)
- Python tem biblioteca `zipfile` nativa para manipular ZIPs
- Arquivo ATP (.$$$) é **texto puro**, fácil de manipular

### **5.2 Buscar Parâmetros: `find_control_parameters()`**

**Localização:** Linhas 64-130

**O Que Faz:** Procura por RPI, RF e outros parâmetros no texto ATP.

```python
def find_control_parameters(self) -> Dict:
    """
    ENTRADA: self.atp_text (texto ATP completo)
    SAÍDA:   {
        'rpi_values': [{'line': 10, 'value': 100}, ...],
        'switch_times': [...],
        'dt': 5e-05,
        'tmax': 0.5
    }
    """
```

**Estratégia de Busca:**

```python
# PASSO 1: Dividir texto em linhas
lines = atp_text.split('\n')

# PASSO 2: Para cada linha
for i, line in enumerate(lines):
    
    # PASSO 3: Procurar por "RPI"
    if 'RPI' in line.upper():
        
        # PASSO 4: Extrair números da linha
        numbers = re.findall(r'[-+]?\d*\.\d+|\d+', line)
        # Encontra: ['100', '1']
        
        # PASSO 5: Salvar informação
        rpi_info = {
            'line': i,
            'value': float(numbers[0]),  # 100.0
            'original_line': line
        }
```

**Exemplo Visual:**

```
Texto ATP:
┌────────────────────────────────────────┐
│ Linha 5:  C RPI1 = 100 Ohms           │ ← "RPI" encontrado!
│ Linha 6:  BUSA BUSRPI1  100.000  1    │ ← Valor numérico
│ Linha 10: C RPI2 = 200 Ohms           │ ← Outro RPI
│ Linha 11: BUSB BUSRPI2  200.000  1    │
└────────────────────────────────────────┘
                ↓
Resultado:
{
  'rpi_values': [
    {'line': 5, 'value': 100},
    {'line': 10, 'value': 200}
  ]
}
```

### **5.3 Modificar Valores: `modify_rpi_value()`**

**Localização:** Linhas 132-175

**O Que Faz:** Substitui valores de RPI no texto ATP.

```python
def modify_rpi_value(self, new_rpi: float) -> bool:
    """
    ENTRADA: new_rpi = 150
    AÇÃO:    RPI=100 → RPI=150 no texto ATP
    SAÍDA:   True se modificado com sucesso
    """
```

**Lógica de Modificação:**

```python
# ANTES:
linha = "  BUSA  BUSRPI1         100.000              1"
                                  ↑↑↑
                                valor antigo

# PASSO 1: Encontrar linha com "RPI"
if 'RPI' in linha.upper():
    
    # PASSO 2: Separar em partes
    parts = linha.split()
    # ['BUSA', 'BUSRPI1', '100.000', '1']
    
    # PASSO 3: Encontrar valor numérico
    for part in parts:
        if part == '100.000':
            
            # PASSO 4: Substituir
            nova_linha = linha.replace('100.000', '150.000')
            
# DEPOIS:
linha = "  BUSA  BUSRPI1         150.000              1"
                                  ↑↑↑
                                novo valor
```

**Cuidados Importantes:**

1. **Preservar formatação:** Manter espaços e alinhamento
2. **Precisão numérica:** Usar formato `f"{new_rpi:.6f}"` (6 casas decimais)
3. **Validação:** Verificar se valor é razoável (0.1 - 100000 Ω)

### **5.4 Salvar Modificado: `save_modified_acp()`**

**Localização:** Linhas 177-240

**O Que Faz:** Reconstrói arquivo .acp com ATP modificado.

```python
def save_modified_acp(self, output_path: Path) -> bool:
    """
    ENTRADA: output_path = "Caso0_RPI=150.acp"
    AÇÃO:    Criar novo .acp com ATP modificado
    SAÍDA:   True se salvo com sucesso
    """
```

**Processo de Salvamento:**

```
1. Copiar .acp original para novo arquivo
   shutil.copy2(original.acp, novo.acp)
   ↓
2. Abrir novo .acp como ZIP (modo escrita)
   with zipfile.ZipFile(novo.acp, 'w') as new_zip:
   ↓
3. Copiar todos arquivos EXCETO .$$$
   for file in original_files:
       if not file.endswith('.$$$'):
           new_zip.writestr(file, original_content)
   ↓
4. Escrever ATP modificado
   new_zip.writestr(
       'arquivo.$$$',
       atp_text_modificado.encode('windows-1252')
   )
   ↓
5. Fechar ZIP
   Arquivo novo.acp está pronto!
```

**Desafio Técnico:**

Python's `zipfile` não permite **deletar** arquivos de um ZIP. Solução:

```python
# Criar ZIP temporário
temp_zip = Path('temp.zip')

# Copiar tudo MENOS o arquivo a ser substituído
with zipfile.ZipFile(original, 'r') as old:
    with zipfile.ZipFile(temp_zip, 'w') as new:
        for item in old.infolist():
            if item.filename != 'arquivo.$$$':
                new.writestr(item, old.read(item.filename))
            else:
                # Escrever versão modificada
                new.writestr(item.filename, modified_text)

# Substituir original por temporário
temp_zip.replace(output_path)
```

### **5.5 Executar ATP: `AtpRunner.run_simulation()`**

**Localização:** Linhas 300-380

**O Que Faz:** Executa o ATP para rodar a simulação.

```python
def run_simulation(self, acp_path: Path) -> Optional[Path]:
    """
    ENTRADA: Path("Caso0_RPI=150.acp")
    AÇÃO:    Executa ATP
    SAÍDA:   Path("Caso0_RPI=150.lis") se sucesso
    """
```

**Fluxo de Execução:**

```
1. Encontrar executável ATP
   atp_exe = '/usr/local/bin/tpbig'
   ↓
2. Preparar comando
   comando = ['tpbig', 'arquivo.atp']
   ↓
3. Executar via subprocess
   result = subprocess.run(
       comando,
       capture_output=True,
       timeout=300  # 5 minutos máximo
   )
   ↓
4. Verificar sucesso
   if result.returncode == 0:
       print("✅ Simulação concluída")
   ↓
5. Procurar arquivo .lis gerado
   lis_file = acp_path.with_suffix('.lis')
   ↓
6. Retornar caminho do .lis
   return lis_file
```

**Exemplo Completo:**

```python
# Criar runner
runner = AtpRunner()

# Executar simulação
lis_file = runner.run_simulation(Path("Caso0_RPI=150.acp"))

if lis_file and lis_file.exists():
    print(f"✅ Resultado: {lis_file}")
    # Agora pode processar com main.py!
else:
    print("❌ Simulação falhou")
```

---
