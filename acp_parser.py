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
        possible_paths = [
            '/usr/local/bin/tpbig',
            '/usr/bin/tpbig',
            '/opt/atp/tpbig',
            'tpbig',
            'atpmingw',
            '/usr/local/bin/atpmingw'
        ]
        
        for path in possible_paths:
            if shutil.which(path):
                return path
        
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
            print("💡 Configure o caminho manualmente: AtpRunner('/caminho/para/tpbig')")
            return None
        
        if not acp_path.exists():
            print(f"❌ Arquivo não encontrado: {acp_path}")
            return None
        
        # Extrair ATP do .acp
        parser = AcpParser(acp_path)
        atp_text = parser.extract_atp_from_acp()
        
        if not atp_text:
            return None
        
        # Criar arquivo temporário .atp
        temp_atp = acp_path.with_suffix('.atp')
        with open(temp_atp, 'w', encoding='windows-1252', errors='ignore') as f:
            f.write(atp_text)
        
        print(f"🚀 Executando simulação ATP: {acp_path.name}")
        
        try:
            # Executar ATP
            result = subprocess.run(
                [self.atpdraw_path, str(temp_atp)],
                cwd=acp_path.parent,
                capture_output=True,
                text=True,
                timeout=120  # 2 minutos timeout
            )
            
            # Procurar arquivo .lis gerado
            lis_path = acp_path.with_suffix('.lis')

            # Verificar código de retorno do processo ATP
            if result.returncode != 0:
                print(f"❌ ATP retornou código {result.returncode}. Considerando falha na simulação.")
                print(f"   Stdout: {result.stdout[:200]}")
                print(f"   Stderr: {result.stderr[:200]}")
                # Se um .lis vazio foi gerado, remover
                if lis_path.exists():
                    try:
                        lis_size = lis_path.stat().st_size
                    except Exception:
                        lis_size = 0
                    if lis_size <= 0:
                        try:
                            lis_path.unlink(missing_ok=True)
                            print(f"🗑️  Arquivo .lis vazio removido: {lis_path}")
                        except Exception as e:
                            print(f"⚠️ Não foi possível remover .lis vazio: {e}")
                return None
            
            if lis_path.exists():
                # Validar tamanho do .lis (> 0 bytes) antes de mover
                try:
                    lis_size = lis_path.stat().st_size
                except Exception:
                    lis_size = 0
                
                if lis_size <= 0:
                    print("⚠️ .lis gerado, porém vazio (0 bytes). Considerando falha na simulação.")
                    print(f"   Stdout: {result.stdout[:200]}")
                    print(f"   Stderr: {result.stderr[:200]}")
                    # Remover arquivo vazio para evitar acúmulo
                    try:
                        lis_path.unlink(missing_ok=True)
                        print(f"🗑️  Arquivo .lis vazio removido: {lis_path}")
                    except Exception as e:
                        print(f"⚠️ Não foi possível remover .lis vazio: {e}")
                    return None
                
                # Mover para output_dir se especificado
                if output_dir:
                    output_dir = Path(output_dir)
                    output_dir.mkdir(parents=True, exist_ok=True)
                    
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    new_lis = output_dir / f"{acp_path.stem}_{timestamp}.lis"
                    
                    shutil.move(lis_path, new_lis)
                    lis_path = new_lis
                
                print(f"✅ Simulação concluída: {lis_path}")
                
                # Limpar arquivos temporários
                temp_atp.unlink(missing_ok=True)
                
                return lis_path
            else:
                print(f"⚠️ Simulação executada mas .lis não foi gerado")
                print(f"   Stdout: {result.stdout[:200]}")
                print(f"   Stderr: {result.stderr[:200]}")
                return None
        
        except subprocess.TimeoutExpired:
            print("❌ Timeout: simulação demorou mais de 5 minutos")
            return None
        except Exception as e:
            print(f"❌ Erro ao executar ATP: {e}")
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
