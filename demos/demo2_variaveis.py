#!/usr/bin/env python3
"""
DEMONSTRAÇÃO 2: Variáveis de Controle
======================================

Este script demonstra o conceito de "variáveis de controle" do tkinter:
- StringVar (texto)
- IntVar (números inteiros)
- BooleanVar (verdadeiro/falso)
- DoubleVar (números decimais)

A magia: quando você edita o campo, a variável é atualizada automaticamente!
E vice-versa: quando você muda a variável, o campo é atualizado!

Use durante a apresentação para explicar a ponte entre interface e dados.
"""

import tkinter as tk
from tkinter import ttk

# ============================================================
# Criar janela
# ============================================================
root = tk.Tk()
root.title("Variáveis de Controle 🎛️")
root.geometry("500x450")

# Frame principal com padding
main_frame = ttk.Frame(root, padding=20)
main_frame.pack(fill='both', expand=True)

# Título
titulo = ttk.Label(
    main_frame,
    text="🎓 Demonstração: Variáveis de Controle",
    font=("Arial", 16, "bold"),
    foreground="#2F75B5"
)
titulo.pack(pady=(0, 20))

# ============================================================
# CRIAR VARIÁVEIS DE CONTROLE
# ============================================================

# StringVar - para textos
nome_var = tk.StringVar(value="João Silva")

# IntVar - para números inteiros
idade_var = tk.IntVar(value=20)

# BooleanVar - para verdadeiro/falso
ativo_var = tk.BooleanVar(value=True)

# DoubleVar - para números decimais
altura_var = tk.DoubleVar(value=1.75)

# ============================================================
# CONECTAR VARIÁVEIS À INTERFACE
# ============================================================

# Seção 1: Nome (StringVar)
secao1 = ttk.LabelFrame(main_frame, text="1️⃣ StringVar (Texto)", padding=10)
secao1.pack(fill='x', pady=5)

ttk.Label(secao1, text="Nome:").pack(side='left')
entry_nome = ttk.Entry(secao1, textvariable=nome_var, width=30)
entry_nome.pack(side='left', padx=10)

# Seção 2: Idade (IntVar)
secao2 = ttk.LabelFrame(main_frame, text="2️⃣ IntVar (Número Inteiro)", padding=10)
secao2.pack(fill='x', pady=5)

ttk.Label(secao2, text="Idade:").pack(side='left')
spin_idade = ttk.Spinbox(secao2, from_=1, to=120, textvariable=idade_var, width=10)
spin_idade.pack(side='left', padx=10)

# Seção 3: Ativo (BooleanVar)
secao3 = ttk.LabelFrame(main_frame, text="3️⃣ BooleanVar (Verdadeiro/Falso)", padding=10)
secao3.pack(fill='x', pady=5)

chk_ativo = ttk.Checkbutton(secao3, text="Usuário Ativo?", variable=ativo_var)
chk_ativo.pack(anchor='w')

# Seção 4: Altura (DoubleVar)
secao4 = ttk.LabelFrame(main_frame, text="4️⃣ DoubleVar (Número Decimal)", padding=10)
secao4.pack(fill='x', pady=5)

ttk.Label(secao4, text="Altura (m):").pack(side='left')
scale_altura = ttk.Scale(
    secao4, 
    from_=1.0, 
    to=2.5, 
    variable=altura_var,
    orient='horizontal',
    length=200
)
scale_altura.pack(side='left', padx=10)
label_altura = ttk.Label(secao4, text=f"{altura_var.get():.2f}m")
label_altura.pack(side='left')

# Atualizar label da altura quando mudar
def atualizar_altura(*args):
    label_altura.config(text=f"{altura_var.get():.2f}m")

altura_var.trace('w', atualizar_altura)

# ============================================================
# BOTÕES DE AÇÃO
# ============================================================

botoes_frame = ttk.Frame(main_frame)
botoes_frame.pack(pady=20)

# Botão para MOSTRAR valores
def mostrar_valores():
    """Mostra os valores atuais das variáveis"""
    msg = "📊 VALORES ATUAIS:\n"
    msg += "=" * 40 + "\n"
    msg += f"Nome:   {nome_var.get()}\n"
    msg += f"Idade:  {idade_var.get()} anos\n"
    msg += f"Ativo:  {'✅ Sim' if ativo_var.get() else '❌ Não'}\n"
    msg += f"Altura: {altura_var.get():.2f}m\n"
    msg += "=" * 40
    
    print("\n" + msg)
    
    # Mostrar em popup também
    tk.messagebox.showinfo("Valores Atuais", msg)

btn_mostrar = ttk.Button(
    botoes_frame,
    text="📊 Mostrar Valores",
    command=mostrar_valores
)
btn_mostrar.pack(side='left', padx=5)

# Botão para MODIFICAR valores programaticamente
def modificar_valores():
    """Modifica valores das variáveis via código"""
    nome_var.set("Maria Santos")
    idade_var.set(25)
    ativo_var.set(False)
    altura_var.set(1.65)
    
    print("\n✏️ Valores modificados programaticamente!")
    print("Note como os campos da interface foram atualizados automaticamente!")

btn_modificar = ttk.Button(
    botoes_frame,
    text="✏️ Modificar Valores",
    command=modificar_valores
)
btn_modificar.pack(side='left', padx=5)

# Botão para RESETAR
def resetar():
    """Reseta para valores padrão"""
    nome_var.set("João Silva")
    idade_var.set(20)
    ativo_var.set(True)
    altura_var.set(1.75)
    print("\n🔄 Valores resetados!")

btn_resetar = ttk.Button(
    botoes_frame,
    text="🔄 Resetar",
    command=resetar
)
btn_resetar.pack(side='left', padx=5)

# ============================================================
# EXPLICAÇÃO
# ============================================================

explicacao = ttk.LabelFrame(main_frame, text="💡 Entenda", padding=10)
explicacao.pack(fill='both', expand=True, pady=(10, 0))

texto_explicacao = """
🎯 VARIÁVEIS DE CONTROLE:

• StringVar, IntVar, BooleanVar, DoubleVar são "pontes"
  entre a interface e os dados do programa

• Quando você EDITA um campo → Variável é ATUALIZADA
• Quando você MUDA a variável → Campo é ATUALIZADO

🔗 CONEXÃO:
  Entry/Spinbox ←→ textvariable=nome_var ←→ Código Python

📝 EXEMPLO:
  nome_var.set("Novo Nome")  → Campo atualiza na tela!
  Digite no campo            → nome_var.get() retorna o novo valor!

💡 NO LIS-ANALYSIS:
  Usamos isso para pasta de entrada, pasta de saída,
  checkboxes de opções, campos de parâmetros, etc!
"""

label_explicacao = ttk.Label(
    explicacao,
    text=texto_explicacao,
    justify='left',
    font=("Courier", 9)
)
label_explicacao.pack(anchor='w')

# ============================================================
# Console
# ============================================================

print("=" * 60)
print("🎓 DEMONSTRAÇÃO 2: VARIÁVEIS DE CONTROLE")
print("=" * 60)
print("\n✨ Interface aberta!")
print("\n📋 INSTRUÇÕES:")
print("   1. Edite os campos na interface")
print("   2. Clique em 'Mostrar Valores' para ver os valores atuais")
print("   3. Clique em 'Modificar Valores' para mudar via código")
print("   4. Observe como interface e variáveis se sincronizam!")
print("\n" + "=" * 60)

# Iniciar loop
root.mainloop()

print("\n👋 Janela fechada! Fim da demonstração 2.")
