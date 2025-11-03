"""
Parser e modificador de arquivos .acp do ATPDraw.
Permite extrair, modificar e executar simulações ATP.
"""

import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Optional, Dict, List, Tuple
import re
import subprocess
import shutil
from datetime import datetime

class AcpParser:
    """Parser para arquivos .acp (ATPDraw)"""
    
    def __init__(self, acp_path: Path):
        self.acp_path = Path(acp_path)
        self.atp_text = None
        self.temp_dir = None
        
    def extract_atp_from_acp(self) -> Optional[str]:
        """
        Extrai o arquivo ATP de dentro do .acp (arquivo ZIP).
        
        Returns:
            String com conteúdo do arquivo ATP ou None
        """
        if not self.acp_path.exists():
            print(f"❌ Arquivo não encontrado: {self.acp_path}")
            return None
        
        try:
            # .acp é um arquivo ZIP
            with zipfile.ZipFile(self.acp_path, 'r') as zip_ref:
                # Procurar por arquivo .$$$
                files = zip_ref.namelist()
                atp_file = None
                
                for f in files:
                    if f.endswith('.$$$'):
                        atp_file = f
                        break
                
                if not atp_file:
                    print(f"❌ Arquivo ATP (.$$$) não encontrado em {self.acp_path.name}")
                    return None
                
                # Ler conteúdo
                with zip_ref.open(atp_file) as f:
                    # ATP usa encoding windows-1252 ou latin-1
                    content = f.read()
                    try:
                        self.atp_text = content.decode('windows-1252')
                    except:
                        self.atp_text = content.decode('latin-1', errors='ignore')
                
                print(f"✅ ATP extraído de {self.acp_path.name} ({len(self.atp_text)} chars)")
                return self.atp_text
        
        except Exception as e:
            print(f"❌ Erro ao extrair ATP: {e}")
            return None
    
    def find_control_parameters(self) -> Dict[str, any]:
        """
        Procura por parâmetros de controle no arquivo ATP.
        Foca em RPI (Resistência de Pré-Inserção).
        
        Returns:
            Dict com parâmetros encontrados
        """
        if not self.atp_text:
            self.extract_atp_from_acp()
        
        if not self.atp_text:
            return {}
        
        params = {
            'rpi_values': [],
            'rpi_lines': [],
            'switch_times': [],
            'dt': None,
            'tmax': None
        }
        
        lines = self.atp_text.split('\n')
        
        for i, line in enumerate(lines):
            # Procurar por resistores de pré-inserção (comentários ou nomes)
            if 'RPI' in line.upper() or 'PRE-INS' in line.upper():
                # Tentar extrair valores numéricos da próxima linha ou mesma linha
                numbers = re.findall(r'[-+]?\d*\.\d+|\d+', line)
                if numbers:
                    params['rpi_values'].append({
                        'line': i,
                        'value': float(numbers[0]),
                        'original_line': line
                    })
                    params['rpi_lines'].append(i)
            
            # Procurar por tempos de chaveamento
            if 'TCLOSE' in line.upper() or 'MEASURING' in line.upper():
                numbers = re.findall(r'[-+]?\d*\.\d+E[+-]?\d+|\d*\.\d+', line)
                if numbers:
                    params['switch_times'].append({
                        'line': i,
                        'time': float(numbers[0]),
                        'original_line': line
                    })
            
            # Procurar dt e tmax (primeiras linhas)
            if i < 10 and re.search(r'\d+\.\d*E[+-]?\d+', line):
                numbers = re.findall(r'[-+]?\d*\.\d+E[+-]?\d+', line)
                if len(numbers) >= 2 and params['dt'] is None:
                    params['dt'] = float(numbers[0])
                    params['tmax'] = float(numbers[1])
        
        return params
    
    def modify_rpi_value(self, new_rpi: float, node_identifier: str = None) -> bool:
        """
        Modifica o valor de RPI no arquivo ATP.
        
        Args:
            new_rpi: Novo valor de RPI em Ohms
            node_identifier: Identificador do nó (opcional, para múltiplos RPIs)
        
        Returns:
            True se modificado com sucesso
        """
        if not self.atp_text:
            self.extract_atp_from_acp()
        
        if not self.atp_text:
            return False
        
        lines = self.atp_text.split('\n')
        modified = False
        
        # Estratégia: procurar linhas com RPI e modificar valores numéricos
        for i, line in enumerate(lines):
            if 'RPI' in line.upper() or 'PRE-INS' in line.upper():
                # Verificar se há identificador de nó
                if node_identifier and node_identifier not in line:
                    continue
                
                # Tentar encontrar e substituir valores numéricos
                # Formato típico: "  NODE1 NODE2        R_VALUE    L_VALUE    C_VALUE"
                parts = line.split()
                
                # Procurar por números em formato científico ou decimal
                new_line = line
                for j, part in enumerate(parts):
                    try:
                        old_value = float(part)
                        # Se é um valor razoável para resistência (1-10000 Ohms)
                        if 0.1 <= old_value <= 100000:
                            # Substituir mantendo formato
                            new_line = line.replace(part, f"{new_rpi:.6f}", 1)
                            lines[i] = new_line
                            modified = True
                            print(f"✅ RPI modificado na linha {i+1}: {old_value} → {new_rpi} Ω")
                            break
                    except ValueError:
                        continue
        
        if modified:
            self.atp_text = '\n'.join(lines)
            return True
        
        print(f"⚠️ Nenhum RPI encontrado para modificar")
        return False
    
    def save_modified_acp(self, output_path: Path) -> bool:
        """
        Salva o arquivo .acp modificado.
        
        Args:
            output_path: Caminho do novo arquivo .acp
        
        Returns:
            True se salvo com sucesso
        """
        if not self.atp_text:
            print("❌ Nenhum conteúdo ATP para salvar")
            return False
        
        try:
            # Copiar .acp original
            shutil.copy2(self.acp_path, output_path)
            
            # Abrir como ZIP e substituir arquivo $$$
            with zipfile.ZipFile(output_path, 'a') as zip_ref:
                # Encontrar nome do arquivo $$$
                files = zip_ref.namelist()
                atp_file = None
                
                for f in files:
                    if f.endswith('.$$$'):
                        atp_file = f
                        break
                
                if not atp_file:
                    print("❌ Arquivo $$$ não encontrado")
                    return False
                
                # Remover arquivo antigo
                # (Nota: zipfile não suporta remoção direta, precisamos recriar)
                temp_zip = output_path.with_suffix('.tmp')
                
                with zipfile.ZipFile(temp_zip, 'w') as new_zip:
                    for item in zip_ref.infolist():
                        if item.filename != atp_file:
                            # Copiar outros arquivos
                            new_zip.writestr(item, zip_ref.read(item.filename))
                        else:
                            # Escrever ATP modificado
                            new_zip.writestr(
                                atp_file, 
                                self.atp_text.encode('windows-1252', errors='ignore')
                            )
                
                # Substituir arquivo
                temp_zip.replace(output_path)
            
            print(f"✅ Arquivo modificado salvo: {output_path}")
            return True
        
        except Exception as e:
            print(f"❌ Erro ao salvar .acp modificado: {e}")
            return False
    
    def print_summary(self):
        """Imprime resumo dos parâmetros encontrados"""
        params = self.find_control_parameters()
        
        print("\n" + "="*80)
        print("📋 RESUMO DO ARQUIVO ATP")
        print("="*80)
        print(f"Arquivo: {self.acp_path.name}")
        
        if params['dt'] and params['tmax']:
            print(f"\n⚙️  Configuração de Tempo:")
            print(f"   dT   = {params['dt']:.6E} s")
            print(f"   Tmax = {params['tmax']:.6f} s")
        
        if params['rpi_values']:
            print(f"\n🔌 Resistências de Pré-Inserção (RPI): {len(params['rpi_values'])}")
            for rpi in params['rpi_values']:
                print(f"   Linha {rpi['line']+1}: {rpi['value']:.2f} Ω")
        
        if params['switch_times']:
            print(f"\n🔀 Tempos de Chaveamento: {len(params['switch_times'])}")
            for sw in params['switch_times'][:5]:  # Mostrar até 5
                print(f"   Linha {sw['line']+1}: {sw['time']:.6f} s")
        
        print("\n" + "="*80 + "\n")


class AtpRunner:
    """Executor de simulações ATP"""
    
    def __init__(self, atpdraw_path: str = None):
        """
        Args:
            atpdraw_path: Caminho para executável do ATP (tpbig, atpmingw, etc)
        """
        self.atpdraw_path = atpdraw_path or self._find_atp_executable()
    
    def _find_atp_executable(self) -> Optional[str]:
        """Tenta encontrar executável do ATP no sistema"""
        import sys
        import os
        
        # Caminhos possíveis para executáveis ATP nativos (Linux/Unix)
        possible_paths = [
            '/usr/local/bin/tpbig',
            '/usr/bin/tpbig',
            '/opt/atp/tpbig',
            'tpbig',
            'atpmingw',
            '/usr/local/bin/atpmingw',
            '/usr/local/bin/atp',
            '/usr/bin/atp'
        ]
        
        # Procurar executáveis nativos primeiro
        for path in possible_paths:
            found = shutil.which(path)
            if found:
                print(f"✅ Executável ATP encontrado: {found}")
                return found
        
        # Se não encontrou nativo, tentar Wine + Atpdraw.exe (apenas Linux)
        if sys.platform.startswith('linux'):
            wine_path = shutil.which('wine')
            if wine_path:
                # Procurar Atpdraw.exe em locais comuns
                atpdraw_locations = [
                    Path.home() / 'ATPDraw' / 'Atpdraw.exe',
                    Path.home() / '.wine' / 'drive_c' / 'ATP' / 'Atpdraw.exe',
                    Path('/opt/atpdraw/Atpdraw.exe'),
                    Path('C:/ATP/Atpdraw.exe')  # Caminho Wine
                ]
                
                for atpdraw in atpdraw_locations:
                    if atpdraw.exists() and os.access(atpdraw, os.R_OK):
                        print(f"✅ ATP via Wine encontrado: wine {atpdraw}")
                        # Retornar comando completo com wine
                        return f"wine {atpdraw}"
        
        print("❌ Nenhum executável ATP encontrado no sistema")
        return None
    
    def run_simulation(self, acp_path: Path, output_dir: Path = None) -> Optional[Path]:
        """
        Executa simulação ATP e retorna caminho do arquivo .lis gerado.
        
        Args:
            acp_path: Caminho do arquivo .acp
            output_dir: Diretório para salvar .lis (padrão: mesmo do .acp)
        
        Returns:
            Path do arquivo .lis gerado ou None
        """
        if not self.atpdraw_path:
            print("❌ Executável do ATP não encontrado!")
            print("\n💡 Opções para resolver:")
            print("   1. Instale o ATP nativo para Linux (tpbig)")
            print("   2. Use Wine + ATPDraw para Windows")
            print("   3. Configure o caminho manualmente na interface")
            return None
        
        if not acp_path.exists():
            print(f"❌ Arquivo não encontrado: {acp_path}")
            return None
        
        # Extrair ATP do .acp
        parser = AcpParser(acp_path)
        atp_text = parser.extract_atp_from_acp()
        
        if not atp_text:
            return None
        
        # Criar arquivo temporário .atp no mesmo diretório do .acp
        temp_atp = acp_path.with_suffix('.atp')
        
        # Escrever conteúdo ATP no arquivo temporário
        try:
            with open(temp_atp, 'w', encoding='windows-1252', errors='ignore') as f:
                f.write(atp_text)
            print(f"📝 Arquivo .atp temporário criado: {temp_atp.name}")
        except Exception as e:
            print(f"❌ Erro ao criar arquivo .atp temporário: {e}")
            return None
        
        print(f"🚀 Executando simulação ATP: {acp_path.name}")
        
        try:
            # Preparar comando baseado no executável
            timeout = 120  # Timeout padrão de 2 minutos
            
            if 'wine' in self.atpdraw_path.lower():
                # Comando com Wine - extrair caminho do executável
                exe_path = self.atpdraw_path.replace('wine', '').strip()
                exe_dir = Path(exe_path).parent
                
                # PRIMEIRO: Procurar por executável de linha de comando (tpbig.exe, ATP.exe, etc)
                print(f"   🔍 Procurando executável CLI em: {exe_dir}")
                cli_executables = [
                    exe_dir / 'tpbig.exe',
                    exe_dir / 'ATP.exe', 
                    exe_dir / 'atpmingw.exe',
                    exe_dir / 'runATP.exe',
                    exe_dir / 'atpdraw' / 'tpbig.exe',
                    exe_dir / 'bin' / 'tpbig.exe',
                ]
                
                cli_found = None
                for cli_exe in cli_executables:
                    if cli_exe.exists():
                        print(f"   ✅ Encontrado executável CLI: {cli_exe}")
                        cli_found = cli_exe
                        break
                
                if cli_found:
                    # Usar executável de linha de comando (melhor opção)
                    cmd = ['wine', str(cli_found), str(temp_atp.absolute())]
                    timeout = 120  # 2 minutos para CLI
                else:
                    # FALLBACK: Usar Atpdraw.exe com opções não-interativas
                    print(f"   ⚠️  Executável CLI não encontrado, usando Atpdraw.exe")
                    print(f"   ℹ️  ATENÇÃO: Atpdraw.exe pode não funcionar via linha de comando")
                    print(f"\n   💡 RECOMENDAÇÃO:")
                    print(f"      Procure por 'tpbig.exe' ou 'ATP.exe' em {exe_dir}")
                    print(f"      e configure manualmente na interface\n")
                    
                    # Tentar com Xvfb + Wine
                    if shutil.which('xvfb-run'):
                        print(f"   ℹ️  Usando Xvfb para executar interface gráfica")
                        cmd = [
                            'xvfb-run', '-a',
                            'wine', exe_path, 
                            str(temp_atp.absolute())
                        ]
                        timeout = 60  # 1 minuto (mais curto porque provavelmente vai travar)
                    else:
                        print(f"   ❌ Xvfb não instalado e executável CLI não encontrado")
                        print(f"   📦 Instale Xvfb: sudo apt install xvfb")
                        temp_atp.unlink(missing_ok=True)
                        return None
                        
            else:
                # Comando nativo
                cmd = [self.atpdraw_path, str(temp_atp.absolute())]
                timeout = 120
            
            print(f"   Comando: {' '.join(cmd)}")
            print(f"   Diretório de trabalho: {acp_path.parent}")
            print(f"   Timeout: {timeout}s")
            print(f"   ⏳ Executando... (aguarde)")
            
            # Executar ATP
            result = subprocess.run(
                cmd,
                cwd=str(acp_path.parent),
                capture_output=True,
                text=True,
                timeout=timeout
            )
            
            print(f"   ✅ Processo finalizado!")
            print(f"   Código de retorno: {result.returncode}")
            
            # Procurar arquivo .lis gerado (pode ter diferentes nomes)
            possible_lis_files = [
                temp_atp.with_suffix('.lis'),  # temp_arquivo.lis
                acp_path.with_suffix('.lis'),   # arquivo_original.lis
                temp_atp.with_suffix('.LIS'),   # Maiúscula
                acp_path.with_suffix('.LIS'),
            ]
            
            # Procurar também por qualquer .lis novo no diretório
            import time
            current_time = time.time()
            
            lis_path = None
            for possible_lis in possible_lis_files:
                if possible_lis.exists():
                    # Verificar se foi modificado recentemente (últimos 30 segundos)
                    if current_time - possible_lis.stat().st_mtime < 30:
                        lis_path = possible_lis
                        print(f"   📄 Arquivo .lis encontrado: {lis_path.name}")
                        break
            
            if not lis_path:
                # Procurar qualquer .lis criado/modificado recentemente
                print(f"   🔍 Procurando arquivos .lis recentes...")
                for lis_file in list(acp_path.parent.glob('*.lis')) + list(acp_path.parent.glob('*.LIS')):
                    if current_time - lis_file.stat().st_mtime < 30:
                        lis_path = lis_file
                        print(f"   📄 Arquivo .lis recente encontrado: {lis_path.name}")
                        break
            
            if lis_path and lis_path.exists():
                # Mover para output_dir se especificado
                if output_dir:
                    output_dir = Path(output_dir)
                    output_dir.mkdir(parents=True, exist_ok=True)
                    
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    new_lis = output_dir / f"{acp_path.stem}_{timestamp}.lis"
                    
                    shutil.copy2(lis_path, new_lis)  # Copiar ao invés de mover
                    
                    # Limpar .lis original se for diferente do .acp original
                    if lis_path != acp_path.with_suffix('.lis'):
                        lis_path.unlink(missing_ok=True)
                    
                    lis_path = new_lis
                
                print(f"✅ Simulação concluída: {lis_path}")
                
                # Limpar arquivos temporários
                temp_atp.unlink(missing_ok=True)
                
                return lis_path
            else:
                print(f"⚠️  Simulação executada mas .lis não foi gerado")
                print(f"   Código de retorno: {result.returncode}")
                if result.stdout:
                    print(f"   Stdout: {result.stdout[:1000]}")
                if result.stderr:
                    print(f"   Stderr: {result.stderr[:1000]}")
                
                # Listar arquivos no diretório para debug
                print(f"\n   📁 Arquivos no diretório (com timestamps):")
                import time
                current_time = time.time()
                for f in sorted(acp_path.parent.glob('*'), key=lambda x: x.stat().st_mtime, reverse=True)[:15]:
                    age = current_time - f.stat().st_mtime
                    print(f"      - {f.name} (modificado há {age:.1f}s)")
                
                return None
        
        except subprocess.TimeoutExpired:
            print(f"❌ Timeout: simulação demorou mais de {timeout} segundos")
            print(f"\n   💡 Possíveis causas:")
            print(f"      1. Atpdraw.exe está esperando interação do usuário")
            print(f"      2. Não existe executável CLI (tpbig.exe)")
            print(f"\n   📦 Procure por 'tpbig.exe' em /home/pedro/ATPDraw/")
            print(f"      e configure-o no campo 'Executável ATP'")
            
            # Tentar limpar processo travado
            try:
                temp_atp.unlink(missing_ok=True)
            except:
                pass
            
            return None
        except PermissionError as e:
            print(f"❌ Erro de permissão ao executar ATP: {e}")
            print(f"\n💡 Possíveis soluções:")
            print(f"   1. Verifique se o arquivo tem permissão de execução:")
            print(f"      chmod +x {self.atpdraw_path}")
            print(f"   2. Se estiver usando Wine, certifique-se que o Wine está instalado:")
            print(f"      sudo apt install wine wine64")
            print(f"   3. Configure um executável válido no campo 'Executável ATP'")
            return None
        except FileNotFoundError as e:
            print(f"❌ Executável não encontrado: {e}")
            print(f"\n💡 O caminho '{self.atpdraw_path}' não existe ou não é válido")
            print(f"   Configure o caminho correto no campo 'Executável ATP' da interface")
            return None
        except Exception as e:
            print(f"❌ Erro ao executar ATP: {e}")
            print(f"\n💡 Detalhes do erro:")
            import traceback
            traceback.print_exc()
            return None
        finally:
            # Limpar arquivos temporários
            temp_atp.unlink(missing_ok=True)


# ==================== FUNÇÕES DE CONVENIÊNCIA ====================

def modify_acp_rpi(acp_path: Path, new_rpi: float, output_path: Path = None) -> Optional[Path]:
    """
    Função simplificada para modificar RPI em um arquivo .acp.
    
    Args:
        acp_path: Arquivo .acp original
        new_rpi: Novo valor de RPI em Ohms
        output_path: Arquivo de saída (padrão: adiciona "_RPI{valor}")
    
    Returns:
        Path do arquivo modificado ou None
    """
    if output_path is None:
        output_path = acp_path.with_name(f"{acp_path.stem}_RPI{int(new_rpi)}.acp")
    
    parser = AcpParser(acp_path)
    parser.extract_atp_from_acp()
    
    if parser.modify_rpi_value(new_rpi):
        if parser.save_modified_acp(output_path):
            return output_path
    
    return None


def run_acp_simulation(acp_path: Path, output_dir: Path = None) -> Optional[Path]:
    """
    Função simplificada para executar simulação ATP.
    
    Args:
        acp_path: Arquivo .acp para simular
        output_dir: Diretório para salvar .lis
    
    Returns:
        Path do arquivo .lis gerado
    """
    runner = AtpRunner()
    return runner.run_simulation(acp_path, output_dir)


# ==================== EXEMPLO DE USO ====================

if __name__ == "__main__":
    # Exemplo de uso
    acp_file = Path("Caso0_Convenc_SemControle.acp")
    
    if acp_file.exists():
        print("="*80)
        print("🔍 ANALISANDO ARQUIVO .acp")
        print("="*80)
        
        # 1. Extrair e analisar
        parser = AcpParser(acp_file)
        parser.extract_atp_from_acp()
        parser.print_summary()
        
        # 2. Modificar RPI
        print("\n" + "="*80)
        print("🔧 MODIFICANDO RPI")
        print("="*80)
        
        new_rpi = 100.0
        output_file = acp_file.with_name(f"{acp_file.stem}_RPI{int(new_rpi)}.acp")
        
        if parser.modify_rpi_value(new_rpi):
            if parser.save_modified_acp(output_file):
                print(f"\n✅ Arquivo modificado salvo: {output_file}")
        
        # 3. Executar simulação (se ATP estiver instalado)
        print("\n" + "="*80)
        print("🚀 EXECUTANDO SIMULAÇÃO")
        print("="*80)
        
        runner = AtpRunner()
        if runner.atpdraw_path:
            lis_result = runner.run_simulation(output_file, output_dir=Path("Simulation_Result"))
            if lis_result:
                print(f"\n✅ Simulação completa! Resultado: {lis_result}")
        else:
            print("⚠️ ATP não encontrado - pulando simulação")
    else:
        print(f"❌ Arquivo não encontrado: {acp_file}")
