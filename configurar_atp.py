#!/usr/bin/env python3
"""
Script auxiliar para detectar e configurar o ATP no sistema
"""

import os
import sys
import shutil
from pathlib import Path

def check_wine():
    """Verifica se Wine está instalado"""
    print("\n" + "="*80)
    print("🍷 Verificando Wine")
    print("="*80)
    
    wine_path = shutil.which('wine')
    if wine_path:
        print(f"✅ Wine instalado: {wine_path}")
        
        # Tentar obter versão
        try:
            import subprocess
            result = subprocess.run(['wine', '--version'], capture_output=True, text=True)
            print(f"   Versão: {result.stdout.strip()}")
        except:
            pass
        
        return True
    else:
        print("❌ Wine NÃO instalado")
        print("\n💡 Para instalar:")
        print("   sudo apt update")
        print("   sudo apt install wine wine64 winetricks")
        return False

def find_atpdraw_exe():
    """Procura por Atpdraw.exe no sistema"""
    print("\n" + "="*80)
    print("🔍 Procurando Atpdraw.exe")
    print("="*80)
    
    search_paths = [
        Path.home() / 'ATPDraw',
        Path.home() / '.wine' / 'drive_c' / 'ATP',
        Path.home() / '.wine' / 'drive_c' / 'Program Files' / 'ATP',
        Path.home() / '.wine' / 'drive_c' / 'Program Files (x86)' / 'ATP',
        Path('/opt/atpdraw'),
        Path('/opt/ATP'),
    ]
    
    found_files = []
    
    for search_path in search_paths:
        if search_path.exists():
            # Procurar recursivamente por Atpdraw.exe
            for exe_file in search_path.rglob('Atpdraw.exe'):
                found_files.append(exe_file)
                print(f"✅ Encontrado: {exe_file}")
                print(f"   Tamanho: {exe_file.stat().st_size / 1024:.1f} KB")
                print(f"   Permissões: {oct(exe_file.stat().st_mode)[-3:]}")
                
                # Verificar se tem permissão de execução
                if os.access(exe_file, os.X_OK):
                    print(f"   ✅ Executável: SIM")
                else:
                    print(f"   ⚠️  Executável: NÃO (use: chmod +x {exe_file})")
    
    if not found_files:
        print("❌ Nenhum Atpdraw.exe encontrado")
        print("\n💡 Possíveis soluções:")
        print("   1. Instale ATPDraw no Wine")
        print("   2. Copie Atpdraw.exe para ~/ATPDraw/")
        print("   3. Use versão nativa do ATP para Linux")
    
    return found_files

def find_native_atp():
    """Procura por executáveis ATP nativos (Linux)"""
    print("\n" + "="*80)
    print("🐧 Procurando ATP Nativo (Linux)")
    print("="*80)
    
    executables = ['tpbig', 'atpmingw', 'atp']
    found = []
    
    for exe in executables:
        path = shutil.which(exe)
        if path:
            found.append(path)
            print(f"✅ {exe}: {path}")
            
            # Tentar obter versão
            try:
                import subprocess
                result = subprocess.run([path, '--version'], capture_output=True, text=True, timeout=2)
                if result.stdout:
                    print(f"   Versão: {result.stdout.strip()[:100]}")
            except:
                pass
    
    if not found:
        print("❌ Nenhum executável ATP nativo encontrado")
        print("\n💡 Para instalar ATP nativo:")
        print("   - Visite: https://www.atpdraw.net/")
        print("   - Baixe versão para Linux")
        print("   - Instale seguindo instruções")
    
    return found

def suggest_configuration():
    """Sugere configuração baseada no que foi encontrado"""
    print("\n" + "="*80)
    print("💡 SUGESTÕES DE CONFIGURAÇÃO")
    print("="*80)
    
    # Verificar Wine
    has_wine = check_wine()
    
    # Procurar executáveis
    native_execs = find_native_atp()
    atpdraw_files = find_atpdraw_exe()
    
    print("\n" + "="*80)
    print("📝 RECOMENDAÇÃO")
    print("="*80)
    
    if native_execs:
        print("\n✅ MELHOR OPÇÃO: Use ATP Nativo")
        print(f"\nNo campo 'Executável ATP' da interface, configure:")
        print(f"   {native_execs[0]}")
    elif has_wine and atpdraw_files:
        atpdraw = atpdraw_files[0]
        print("\n⚠️  OPÇÃO ALTERNATIVA: Use Wine + Atpdraw.exe")
        
        # Verificar permissão
        if not os.access(atpdraw, os.X_OK):
            print(f"\n1. Primeiro, dê permissão de execução:")
            print(f"   chmod +x {atpdraw}")
        
        print(f"\n2. No campo 'Executável ATP' da interface, configure:")
        print(f"   wine {atpdraw}")
        
        print(f"\n3. OU crie um script wrapper:")
        wrapper_path = Path.home() / 'run_atpdraw.sh'
        print(f"   echo '#!/bin/bash' > {wrapper_path}")
        print(f"   echo 'wine {atpdraw} \"$@\"' >> {wrapper_path}")
        print(f"   chmod +x {wrapper_path}")
        print(f"\n   E configure na interface: {wrapper_path}")
    else:
        print("\n❌ NENHUMA OPÇÃO DISPONÍVEL")
        print("\nVocê precisa:")
        
        if not has_wine:
            print("\n1. Instalar Wine:")
            print("   sudo apt update")
            print("   sudo apt install wine wine64")
        
        if not atpdraw_files:
            print("\n2. Instalar ATPDraw ou copiar para ~/ATPDraw/")
        
        print("\nOU")
        print("\n3. Instalar ATP nativo para Linux")
        print("   Visite: https://www.atpdraw.net/")

def test_execution(executable_path):
    """Testa execução do ATP"""
    print("\n" + "="*80)
    print(f"🧪 Testando Execução: {executable_path}")
    print("="*80)
    
    try:
        import subprocess
        
        # Preparar comando
        if executable_path.startswith('wine '):
            cmd = executable_path.split(' ', 1)
        else:
            cmd = [executable_path]
        
        print(f"Comando: {cmd}")
        
        # Tentar executar com --help ou --version
        result = subprocess.run(
            cmd + ['--help'],
            capture_output=True,
            text=True,
            timeout=5
        )
        
        print(f"Código de retorno: {result.returncode}")
        
        if result.stdout:
            print(f"Stdout: {result.stdout[:200]}")
        
        if result.stderr:
            print(f"Stderr: {result.stderr[:200]}")
        
        print("\n✅ Executável responde a comandos")
        
    except Exception as e:
        print(f"\n❌ Erro ao testar execução: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    print("\n" + "="*80)
    print("🔧 DIAGNÓSTICO E CONFIGURAÇÃO DO ATP")
    print("="*80)
    
    suggest_configuration()
    
    print("\n" + "="*80)
    print("📖 PRÓXIMOS PASSOS")
    print("="*80)
    
    print("\n1. Siga as recomendações acima")
    print("2. Abra a interface: python3 main.py --gui")
    print("3. Configure o campo 'Executável ATP'")
    print("4. Selecione arquivos .acp e clique em '🚀 Rodar ATP'")
    
    print("\n📚 Documentação completa: INSTRUCOES_ATP.md")
    print("\n" + "="*80 + "\n")
