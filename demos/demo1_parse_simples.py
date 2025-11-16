"""
📊 DEMO 1: Leitura Simples de Arquivo .lis

Este script demonstra como o programa lê e extrai dados de um arquivo .lis
"""

from pathlib import Path
from main import parse_lis_table

print("╔═══════════════════════════════════════════════════════════╗")
print("║   DEMO 1: Como o Programa Lê um Arquivo .lis             ║")
print("╚═══════════════════════════════════════════════════════════╝\n")

# Arquivo de exemplo
arquivo = Path("caso0_convenc_semcontrole.lis")

if not arquivo.exists():
    print(f"❌ Arquivo não encontrado: {arquivo}")
    print("💡 Execute este script na pasta raiz do projeto")
    exit(1)

print(f"📄 Lendo arquivo: {arquivo.name}\n")
print("⏳ Processando...")

# Ler arquivo usando a função principal
df, stats_lines, summary = parse_lis_table(arquivo)

if df is None:
    print("❌ Nenhuma tabela encontrada no arquivo")
    exit(1)

print("✅ Arquivo lido com sucesso!\n")

# Mostrar informações da tabela
print("─" * 60)
print("📊 TABELA EXTRAÍDA (primeiras 5 linhas):")
print("─" * 60)
print(df.head())

print(f"\n📈 Total de linhas: {len(df)}")
print(f"📊 Colunas disponíveis: {list(df.columns)}")

# Mostrar estatísticas do ATP (se houver)
if summary:
    print("\n" + "─" * 60)
    print("📈 ESTATÍSTICAS EXTRAÍDAS DO ATP:")
    print("─" * 60)
    for key, (grouped, ungrouped) in summary.items():
        if grouped is not None:
            print(f"  {key.upper()}: {grouped}")
            if ungrouped is not None:
                print(f"    (Ungrouped: {ungrouped})")
else:
    print("\n⚠️  Nenhuma estatística encontrada no arquivo")

# Mostrar algumas informações úteis
print("\n" + "─" * 60)
print("🔍 ANÁLISE RÁPIDA DOS DADOS:")
print("─" * 60)
print(f"  Tensão mínima: {df['Voltage_per_unit'].min():.3f} pu")
print(f"  Tensão máxima: {df['Voltage_per_unit'].max():.3f} pu")
print(f"  Frequência total: {df['Frequency'].sum():.0f}")

print("\n" + "═" * 60)
print("✨ DEMO CONCLUÍDA!")
print("═" * 60)
print("\n💡 Próximos passos:")
print("   1. Execute: python demo2_estatisticas.py")
print("   2. Execute: python demo3_grafico.py")
print("   3. Ou use a GUI: python main.py --gui")
