"""
🧮 DEMO 2: Cálculo de Estatísticas

Este script demonstra como calculamos estatísticas avançadas
a partir dos dados extraídos do arquivo .lis
"""

from pathlib import Path
from main import parse_lis_table, calcular_estatisticas_do_df

print("╔═══════════════════════════════════════════════════════════╗")
print("║   DEMO 2: Cálculo de Estatísticas Avançadas              ║")
print("╚═══════════════════════════════════════════════════════════╝\n")

# Arquivo de exemplo
arquivo = Path("caso0_convenc_semcontrole.lis")

if not arquivo.exists():
    print(f"❌ Arquivo não encontrado: {arquivo}")
    print("💡 Execute este script na pasta raiz do projeto")
    exit(1)

print(f"📄 Processando: {arquivo.name}\n")

# 1. Ler dados
print("1️⃣  Lendo arquivo...")
df, stats_lines, summary_atp = parse_lis_table(arquivo)

if df is None:
    print("❌ Erro ao ler arquivo")
    exit(1)

print(f"   ✅ {len(df)} linhas extraídas\n")

# 2. Calcular estatísticas
print("2️⃣  Calculando estatísticas...")
try:
    stats = calcular_estatisticas_do_df(df)
    print("   ✅ Estatísticas calculadas\n")
except Exception as e:
    print(f"   ❌ Erro: {e}")
    exit(1)

# 3. Mostrar resultados
print("═" * 60)
print("📊 ESTATÍSTICAS CALCULADAS")
print("═" * 60)

print("\n🎯 MEDIDAS DE TENDÊNCIA CENTRAL:")
print(f"   Média (μ):        {stats['mean']:.6f}")
print(f"   Mediana:          {stats['median']:.6f}")
print(f"   Moda:             {stats['mode']:.6f}")

print("\n📏 MEDIDAS DE DISPERSÃO:")
print(f"   Variância (σ²):   {stats['variance']:.6e}")
print(f"   Desvio Padrão (σ): {stats['std_dev']:.6f}")
print(f"   Coef. Variação:   {stats['cv']:.4f}")

print("\n📈 MEDIDAS DE FORMA:")
print(f"   Assimetria:       {stats['skewness']:.4f}")
print(f"   Curtose:          {stats['kurtosis']:.4f}")

print("\n🎲 INFORMAÇÕES ADICIONAIS:")
print(f"   Total Frequências: {stats['total_freq']:.0f}")
print(f"   R² (Ajuste):      {stats['r2']:.4f}")
print(f"   Método usado:     {stats['freq_method']}")

# 4. Comparar com estatísticas do ATP (se houver)
if summary_atp:
    print("\n" + "═" * 60)
    print("🔬 COMPARAÇÃO: Nossos Cálculos vs ATP")
    print("═" * 60)
    
    comparacao = {
        'mean': ('Média', 'mean'),
        'std_dev': ('Desvio Padrão', 'std_dev')
    }
    
    for key_atp, (nome, key_stats) in comparacao.items():
        if key_atp in summary_atp:
            valor_atp = summary_atp[key_atp][0]  # grouped value
            valor_calc = stats.get(key_stats)
            
            if valor_atp is not None and valor_calc is not None:
                diferenca = abs(valor_atp - valor_calc)
                percentual = (diferenca / valor_atp) * 100 if valor_atp != 0 else 0
                
                print(f"\n{nome}:")
                print(f"   ATP:    {valor_atp:.6f}")
                print(f"   Nosso:  {valor_calc:.6f}")
                print(f"   Dif:    {diferenca:.6e} ({percentual:.2f}%)")

# 5. Interpretação visual
print("\n" + "═" * 60)
print("💡 INTERPRETAÇÃO DOS RESULTADOS")
print("═" * 60)

# Interpretação do R²
if stats['r2'] > 0.95:
    r2_msg = "Excelente! Dados seguem distribuição gaussiana"
elif stats['r2'] > 0.90:
    r2_msg = "Bom ajuste à distribuição gaussiana"
elif stats['r2'] > 0.80:
    r2_msg = "Ajuste razoável"
else:
    r2_msg = "Dados não seguem bem uma gaussiana"

print(f"\n📊 R² = {stats['r2']:.4f}")
print(f"   → {r2_msg}")

# Interpretação da assimetria
skew = stats['skewness']
if abs(skew) < 0.5:
    skew_msg = "Distribuição aproximadamente simétrica"
elif skew > 0:
    skew_msg = "Distribuição com cauda à direita (assimetria positiva)"
else:
    skew_msg = "Distribuição com cauda à esquerda (assimetria negativa)"

print(f"\n📈 Assimetria = {skew:.4f}")
print(f"   → {skew_msg}")

# Interpretação da curtose
kurt = stats['kurtosis']
if abs(kurt) < 0.5:
    kurt_msg = "Distribuição mesocúrtica (normal)"
elif kurt > 0:
    kurt_msg = "Distribuição leptocúrtica (concentrada)"
else:
    kurt_msg = "Distribuição platicúrtica (achatada)"

print(f"\n📊 Curtose = {kurt:.4f}")
print(f"   → {kurt_msg}")

# Visualização ASCII da curva gaussiana
print("\n" + "═" * 60)
print("📈 REPRESENTAÇÃO VISUAL DA DISTRIBUIÇÃO")
print("═" * 60)

mu = stats['mean']
sigma = stats['std_dev']

print(f"\n       μ-2σ    μ-σ     μ      μ+σ    μ+2σ")
print(f"       {mu-2*sigma:.2f}  {mu-sigma:.2f}  {mu:.2f}  {mu+sigma:.2f}  {mu+2*sigma:.2f}")
print( "         |      |       |       |       |")
print( "         ╲      ╱       |       ╲      ╱")
print( "          ╲    ╱        |        ╲    ╱")
print( "           ╲  ╱         |         ╲  ╱")
print( "            ╲╱          ▼          ╲╱")
print( "     ────────●──────────●──────────●────────")
print(f"     ← 68.3% →          ← 95.4% →")

print("\n💡 68.3% dos valores estão entre μ±σ")
print(f"   ({mu-sigma:.3f} e {mu+sigma:.3f})")

print("\n" + "═" * 60)
print("✨ DEMO CONCLUÍDA!")
print("═" * 60)
print("\n💡 Próximo passo:")
print("   Execute: python demo3_grafico.py")
