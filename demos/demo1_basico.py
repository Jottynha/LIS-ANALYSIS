#!/usr/bin/env python3
"""
DEMONSTRAÇÃO 1: Janela Básica com Tkinter
==========================================

Este script mostra o básico do tkinter:
- Criar janela
- Adicionar label
- Adicionar botão
- Interação simples

Use durante a apresentação para explicar os fundamentos!
"""

import tkinter as tk
from tkinter import ttk

# ============================================================
# ETAPA 1: Criar janela principal
# ============================================================
print("✅ ETAPA 1: Criando janela...")

root = tk.Tk()
root.title("Minha Primeira Janela 🪟")
root.geometry("400x250")

# ============================================================
# ETAPA 2: Adicionar label (texto)
# ============================================================
print("✅ ETAPA 2: Adicionando label...")

label = ttk.Label(
    root, 
    text="Olá Mundo! 👋", 
    font=("Arial", 24, "bold"),
    foreground="#2F75B5"
)
label.pack(pady=30)

# Label de instrução
instrucao = ttk.Label(
    root,
    text="Clique no botão abaixo para interagir!",
    font=("Arial", 10),
    foreground="gray"
)
instrucao.pack(pady=10)

# ============================================================
# ETAPA 3: Adicionar botão com função
# ============================================================
print("✅ ETAPA 3: Adicionando botão...")

# Variável de controle para contar cliques
contador = 0

def ao_clicar():
    """Função executada quando o botão é clicado"""
    global contador
    contador += 1
    
    # Atualizar texto do label
    label.config(text=f"Você clicou {contador} vez(es)! 🎉")
    
    # Mudar cor conforme número de cliques
    if contador < 5:
        label.config(foreground="#2F75B5")  # Azul
    elif contador < 10:
        label.config(foreground="#C55A11")  # Laranja
    else:
        label.config(foreground="#548235")  # Verde

botao = ttk.Button(
    root, 
    text="🖱️ Clique Aqui", 
    command=ao_clicar
)
botao.pack(pady=10)

# Botão para resetar
def resetar():
    global contador
    contador = 0
    label.config(text="Olá Mundo! 👋", foreground="#2F75B5")

botao_reset = ttk.Button(
    root,
    text="🔄 Resetar",
    command=resetar
)
botao_reset.pack(pady=5)

# ============================================================
# ETAPA 4: Iniciar loop de eventos
# ============================================================
print("✅ ETAPA 4: Iniciando interface...\n")
print("=" * 60)
print("🎓 CONCEITOS DEMONSTRADOS:")
print("=" * 60)
print("1. tk.Tk()         → Cria janela principal")
print("2. ttk.Label()     → Cria texto na interface")
print("3. ttk.Button()    → Cria botão clicável")
print("4. command=...     → Conecta botão a uma função")
print("5. .pack()         → Posiciona elementos na tela")
print("6. root.mainloop() → Mantém janela aberta")
print("=" * 60)
print("\n✨ Janela aberta! Interaja com ela...")

# Mantém a janela aberta até ser fechada
root.mainloop()

print("\n👋 Janela fechada! Fim da demonstração 1.")
