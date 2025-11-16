#!/usr/bin/env python3
"""
🎬 EXECUÇÃO COMPLETA DAS DEMOS

Este script executa todas as demonstrações em sequência,
perfeito para apresentações ao vivo!
"""

import subprocess
import sys
from pathlib import Path

def print_header(text):
    print("\n" + "═" * 70)
    print(f"  {text}")
    print("═" * 70 + "\n")

def run_demo(script, description):
    print_header(description)
    
    print(f"📄 Executando: {script}\n")
    
    try:
        result = subprocess.run(
            [sys.executable, script],
            check=True,
            text=True
        )
        
        print("\n✅ Demo concluída com sucesso!")
        return True
        
    except subprocess.CalledProcessError as e:
        print(f"\n❌ Erro ao executar demo: {e}")
        return False
    except FileNotFoundError:
        print(f"\n❌ Arquivo não encontrado: {script}")
        return False

def main():
    print("╔════════════════════════════════════════════════════════════════╗")
    print("║                                                                ║")
    print("║     🎬 DEMONSTRAÇÃO COMPLETA - LIS-ANALYSIS                   ║")
    print("║                                                                ║")
    print("║     Sistema de Análise Automática de Resultados ATP           ║")
    print("║                                                                ║")
    print("╚════════════════════════════════════════════════════════════════╝")
    
    # Verificar se estamos na pasta correta
    if not Path("caso0_convenc_semcontrole.lis").exists():
        print("\n❌ ERRO: Execute este script na pasta raiz do projeto!")
        print("   (Onde está o arquivo caso0_convenc_semcontrole.lis)")
        sys.exit(1)
    
    demos = [
        ("demo1_parse_simples.py", "DEMO 1: Leitura de Arquivos .lis"),
        ("demo2_estatisticas.py", "DEMO 2: Cálculo de Estatísticas"),
        ("demo3_grafico.py", "DEMO 3: Criação de Gráficos"),
    ]
    
    print("\n📋 DEMONSTRAÇÕES DISPONÍVEIS:\n")
    for i, (script, desc) in enumerate(demos, 1):
        print(f"   {i}. {desc}")
    
    print("\n" + "─" * 70)
    
    choice = input("\nExecutar todas as demos em sequência? (S/n): ").strip().lower()
    
    if choice in ['n', 'nao', 'não', 'no']:
        print("\n❌ Execução cancelada pelo usuário")
        return
    
    # Executar demos
    results = []
    
    for script, description in demos:
        if Path(script).exists():
            success = run_demo(script, description)
            results.append((description, success))
            
            if not success:
                print("\n⚠️  Erro encontrado. Continuar? (S/n): ", end="")
                if input().strip().lower() in ['n', 'nao', 'não', 'no']:
                    break
            
            print("\n" + "─" * 70)
            input("Pressione ENTER para continuar...")
        else:
            print(f"\n⚠️  Script não encontrado: {script}")
            results.append((description, False))
    
    # Resumo final
    print_header("📊 RESUMO DA EXECUÇÃO")
    
    for desc, success in results:
        status = "✅" if success else "❌"
        print(f"{status} {desc}")
    
    total = len(results)
    sucessos = sum(1 for _, s in results if s)
    
    print(f"\n📈 Taxa de sucesso: {sucessos}/{total} ({sucessos/total*100:.0f}%)")
    
    if sucessos == total:
        print("\n🎉 TODAS AS DEMOS EXECUTADAS COM SUCESSO!")
    else:
        print("\n⚠️  Algumas demos falharam. Verifique os erros acima.")
    
    # Arquivos gerados
    print("\n" + "═" * 70)
    print("📁 ARQUIVOS GERADOS")
    print("═" * 70)
    
    saida = Path("Simulation_Result")
    if saida.exists():
        excel_files = list(saida.glob("*.xlsx"))
        png_files = list(saida.glob("*.png"))
        
        print(f"\n📊 Excel: {len(excel_files)} arquivo(s)")
        for f in excel_files:
            print(f"   • {f.name}")
        
        print(f"\n📈 Gráficos: {len(png_files)} arquivo(s)")
        for f in png_files:
            print(f"   • {f.name}")
    else:
        print("\n⚠️  Pasta de saída não encontrada")
    
    print("\n" + "═" * 70)
    print("💡 PRÓXIMOS PASSOS")
    print("═" * 70)
    
    print("""
1. 📖 Leia o guia completo:
   • GUIA_APRESENTACAO_MAIN.md
   
2. 🎮 Experimente a GUI:
   • python main.py --gui
   
3. 📊 Processe seus próprios arquivos:
   • python main.py --folder "sua_pasta"
   
4. 📚 Veja outros guias:
   • GUIA_APRESENTACAO_GUI.md
   • GUIA_CONTROLE_INTELIGENTE.md
""")
    
    print("═" * 70)
    print("✨ DEMONSTRAÇÃO CONCLUÍDA!")
    print("═" * 70)

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n⚠️  Execução interrompida pelo usuário (Ctrl+C)")
        sys.exit(130)
    except Exception as e:
        print(f"\n❌ ERRO INESPERADO: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
