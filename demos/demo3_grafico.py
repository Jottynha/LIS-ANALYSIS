"""
📈 DEMO 3: Criação de Gráfico com Ajuste Gaussiano

Este script demonstra como criamos gráficos profissionais
a partir dos dados processados
"""

from pathlib import Path
from main import (
    parse_lis_table,
    save_df_to_excel_only,
    calcular_estatisticas_do_df,
    escrever_estatisticas_excel,
    criar_grafico_a_partir_do_excel
)
import time

print("╔═══════════════════════════════════════════════════════════╗")
print("║   DEMO 3: Criação de Gráfico com Ajuste Gaussiano        ║")
print("╚═══════════════════════════════════════════════════════════╝\n")

# Arquivo de exemplo
arquivo_lis = Path("caso0_convenc_semcontrole.lis")
pasta_saida = Path("Simulation_Result")
arquivo_excel = pasta_saida / "Demo_Grafico.xlsx"

if not arquivo_lis.exists():
    print(f"❌ Arquivo não encontrado: {arquivo_lis}")
    print("💡 Execute este script na pasta raiz do projeto")
    exit(1)

print(f"📄 Processando: {arquivo_lis.name}\n")

# Passo 1: Ler dados
print("1️⃣  Lendo arquivo .lis...")
df, stats_lines, summary_atp = parse_lis_table(arquivo_lis)

if df is None:
    print("❌ Erro ao ler arquivo")
    exit(1)

print(f"   ✅ {len(df)} pontos extraídos")
time.sleep(0.5)

# Passo 2: Calcular estatísticas
print("\n2️⃣  Calculando estatísticas...")
try:
    stats = calcular_estatisticas_do_df(df)
    print(f"   ✅ Média: {stats['mean']:.4f}, σ: {stats['std_dev']:.4f}")
except Exception as e:
    print(f"   ❌ Erro: {e}")
    exit(1)

time.sleep(0.5)

# Passo 3: Salvar em Excel
print("\n3️⃣  Salvando em Excel...")
try:
    # Criar pasta de saída
    pasta_saida.mkdir(parents=True, exist_ok=True)
    
    # Salvar aba "Dados"
    save_df_to_excel_only(df, arquivo_excel, sheet_name='Dados')
    print(f"   ✅ Aba 'Dados' criada")
    
    # Salvar aba "Estatisticas"
    escrever_estatisticas_excel(arquivo_excel, stats, summary_atp)
    print(f"   ✅ Aba 'Estatisticas' criada")
    print(f"   📁 {arquivo_excel}")
except Exception as e:
    print(f"   ❌ Erro: {e}")
    exit(1)

time.sleep(0.5)

# Passo 4: Criar gráfico
print("\n4️⃣  Criando gráfico...")
try:
    png_path = criar_grafico_a_partir_do_excel(
        excel_path=arquivo_excel,
        outdir=pasta_saida,
        sim_index=999,
        salvar_png=True,
        mostrar=False  # Não mostrar ainda
    )
    
    if png_path and png_path.exists():
        print(f"   ✅ Gráfico salvo")
        print(f"   📁 {png_path}")
    else:
        print("   ⚠️  Gráfico não foi criado")
        
except Exception as e:
    print(f"   ❌ Erro: {e}")
    import traceback
    traceback.print_exc()
    exit(1)

# Resumo
print("\n" + "═" * 60)
print("✅ ARQUIVOS GERADOS")
print("═" * 60)
print(f"\n📊 Excel:   {arquivo_excel}")
print(f"📈 Gráfico: {png_path}")

print("\n" + "═" * 60)
print("📊 ELEMENTOS DO GRÁFICO")
print("═" * 60)

print("""
O gráfico contém:

1. 📊 BARRAS AZUIS
   └─ Mostram a frequência de cada valor de tensão
   
2. 🔵 PONTOS AZUIS
   └─ Dados reais plotados individualmente
   
3. 🟠 CURVA LARANJA (Ajuste Gaussiano)
   └─ Distribuição normal ajustada aos dados
   └─ Fórmula: f(x) = (1/(σ√2π)) * e^(-(x-μ)²/2σ²)
   
4. 🟢 LINHA VERDE (Acumulado %)
   └─ Eixo secundário mostrando percentual acumulado
   
5. 📝 CAIXA DE ESTATÍSTICAS
   └─ μ, σ, mediana, moda, R², etc
""")

print("\n" + "═" * 60)
print("💡 INTERPRETANDO O GRÁFICO")
print("═" * 60)

mu = stats['mean']
sigma = stats['std_dev']
r2 = stats['r2']

print(f"\n📍 Centro da distribuição (μ): {mu:.4f} pu")
print(f"   → A tensão mais provável")

print(f"\n📏 Dispersão (σ): {sigma:.4f} pu")
print(f"   → Quão espalhados estão os dados")
print(f"   → 68% dos valores entre {mu-sigma:.4f} e {mu+sigma:.4f}")

print(f"\n🎯 Qualidade do ajuste (R²): {r2:.4f}")
if r2 > 0.95:
    print(f"   → Excelente! Curva gaussiana explica {r2*100:.1f}% dos dados")
elif r2 > 0.90:
    print(f"   → Bom ajuste ({r2*100:.1f}%)")
else:
    print(f"   → Ajuste moderado ({r2*100:.1f}%)")

# Mostrar comando para abrir o gráfico
print("\n" + "═" * 60)
print("👁️  VISUALIZAR GRÁFICO")
print("═" * 60)

import sys
if sys.platform.startswith('linux'):
    cmd = f"xdg-open '{png_path}'"
elif sys.platform == 'darwin':
    cmd = f"open '{png_path}'"
else:
    cmd = f"start '{png_path}'"

print(f"\n💡 Para abrir o gráfico, execute:")
print(f"   {cmd}")

# Opção de mostrar agora
print("\n" + "─" * 60)
resposta = input("❓ Deseja abrir o gráfico agora? (s/N): ").strip().lower()

if resposta in ['s', 'sim', 'y', 'yes']:
    print("\n📈 Abrindo gráfico...")
    try:
        if sys.platform.startswith('linux'):
            import subprocess
            subprocess.Popen(['xdg-open', str(png_path)])
        elif sys.platform == 'darwin':
            import subprocess
            subprocess.Popen(['open', str(png_path)])
        elif sys.platform == 'win32':
            import os
            os.startfile(str(png_path))
        
        print("✅ Gráfico aberto!")
    except Exception as e:
        print(f"⚠️  Não foi possível abrir automaticamente: {e}")
        print(f"   Abra manualmente: {png_path}")

print("\n" + "═" * 60)
print("✨ DEMO CONCLUÍDA!")
print("═" * 60)

print("\n💡 Próximos passos:")
print("   1. Examine o Excel gerado")
print("   2. Analise o gráfico em detalhes")
print("   3. Teste com seus próprios arquivos .lis")
print("   4. Use a GUI: python main.py --gui")

print("\n📚 Para entender melhor:")
print("   Veja: GUIA_APRESENTACAO_MAIN.md")
