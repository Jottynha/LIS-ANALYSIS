"""
Parser e modificador de arquivos .acp do ATPDraw.
Permite extrair, modificar e executar simulações ATP.
"""

import zipfile
import os
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

        # Suporte Windows: procurar .exe e runATP.bat
        if os.name == 'nt':
            possible_paths = [
                'runATP.bat',
                'tpbig.exe',
                'atpmingw.exe',
                *possible_paths
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
            # Determinar extensão e diretório de execução
            cmd: List[str]
            ext = Path(self.atpdraw_path).suffix.lower() if self.atpdraw_path else ''
            # Política: manter CWD no solver para achar 'startup' e outros artefatos
            solver_executable = shutil.which(self.atpdraw_path) or self.atpdraw_path
            solver_path = Path(solver_executable)
            run_cwd = solver_path.parent if solver_path.exists() else acp_path.parent

            # Sanitizar nome do deck e copiar para pasta do solver
            sanitized_name = re.sub(r'[=\s]+', '_', temp_atp.name)
            deck_in_solver = run_cwd / sanitized_name
            try:
                shutil.copy2(temp_atp, deck_in_solver)
            except Exception:
                # Fallback: usar caminho original
                deck_in_solver = temp_atp

            # Heurística simples para copiar includes para o CWD do solver
            try:
                include_pat = re.compile(r'\b(INCLUDE|\$INCLUDE|\.INC)\b', re.IGNORECASE)
                for line in atp_text.splitlines():
                    if include_pat.search(line):
                        # Tenta extrair caminho entre aspas, caso contrário último token
                        m = re.search(r'"([^"]+)"|\'([^\']+)\'', line)
                        candidate = None
                        if m:
                            candidate = m.group(1) or m.group(2)
                        else:
                            parts = line.strip().split()
                            candidate = parts[-1] if parts else None
                        if candidate:
                            inc_path = (acp_path.parent / candidate).resolve()
                            if inc_path.exists() and inc_path.is_file():
                                target = run_cwd / inc_path.name
                                if not target.exists():
                                    try:
                                        shutil.copy2(inc_path, target)
                                    except Exception:
                                        pass
            except Exception:
                pass

            # Diretório de saída efetivo: se não especificado, usar pasta /ACP do projeto
            effective_output_dir = Path(output_dir) if output_dir else self._default_output_dir(acp_path)
            # Determinar diretórios que serão monitorados para novos arquivos
            search_dirs: List[Path] = [run_cwd]
            if acp_path.parent not in search_dirs:
                search_dirs.append(acp_path.parent)
            # Se for .bat, também verificar pasta do script (pode ser diferente)
            if ext in ['.bat', '.cmd']:
                script_dir = Path(self.atpdraw_path).parent
                if script_dir.exists() and script_dir not in search_dirs:
                    search_dirs.append(script_dir)

            # Listar arquivos antes para detectar novos gerados em cada diretório
            before_files = {str(d): set(p.name for p in d.glob('*')) for d in search_dirs if d.exists()}

            # Montar comando com suporte a .bat/.cmd (Windows ou Wine)
            # Usar somente o nome do deck se ele estiver no CWD
            deck_arg = deck_in_solver.name if deck_in_solver.parent == run_cwd else str(deck_in_solver)
            if ext in ['.bat', '.cmd']:
                script_path = Path(self.atpdraw_path)
                if os.name == 'nt':
                    cmd = ['cmd', '/c', str(script_path), deck_arg]
                else:
                    if shutil.which('wine'):
                        cmd = ['wine', 'cmd', '/c', str(script_path), deck_arg]
                    else:
                        print("❌ Não é possível executar .bat neste sistema (Wine não encontrado).")
                        print("💡 Use tpbig/atpmingw nativo ou instale o Wine para usar scripts .bat.")
                        return None
            else:
                cmd = [self.atpdraw_path, deck_arg]

            # Executar ATP com controle de timeout robusto
            result_stdout = ''
            result_stderr = ''
            result_returncode = None
            try:
                # Preferir Popen para poder encerrar árvore de processos em timeout
                if os.name == 'nt':
                    proc = subprocess.Popen(cmd, cwd=run_cwd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, stdin=subprocess.PIPE, text=True)
                else:
                    # Em POSIX, criar novo grupo para matar filhos em cascata
                    import os as _os, signal as _signal  # locais para evitar shadow
                    proc = subprocess.Popen(cmd, cwd=run_cwd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, stdin=subprocess.PIPE, text=True, preexec_fn=_os.setsid)
                # Timeout padrão 300s (5 min). Pode ser sobrescrito pela variável de ambiente ATP_TIMEOUT
                timeout_sec = 300
                try:
                    env_timeout = os.environ.get('ATP_TIMEOUT')
                    if env_timeout:
                        timeout_sec = int(env_timeout)
                except Exception:
                    pass
                # Enviar 'go' para avançar caso haja PAUSE no deck; se não necessário, é ignorado
                result_stdout, result_stderr = proc.communicate(input='go\n', timeout=timeout_sec)
                result_returncode = proc.returncode
            except subprocess.TimeoutExpired:
                # Tentar encerrar processo e filhos
                if os.name == 'nt':
                    try:
                        subprocess.run(['taskkill', '/PID', str(proc.pid), '/T', '/F'], stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
                    except Exception:
                        pass
                else:
                    try:
                        import os as _os, signal as _signal
                        _os.killpg(_os.getpgid(proc.pid), _signal.SIGKILL)
                    except Exception:
                        pass
                result_returncode = -9
                result_stdout = (result_stdout or '')
                result_stderr = (result_stderr or '') + f"\n[timeout] Processo excedeu 300s e foi terminado."
            except Exception as e:
                # Falha genérica ao executar
                try:
                    if proc and proc.poll() is None:
                        proc.kill()
                except Exception:
                    pass
                result_returncode = -1
                result_stderr = f"Falha ao executar ATP: {e}"

            # Listar arquivos depois (em todos os diretórios monitorados)
            after_files = {str(d): set(p.name for p in d.glob('*')) for d in search_dirs if d.exists()}
            new_files_per_dir = {}
            for d in before_files:
                before_set = before_files.get(d, set())
                after_set = after_files.get(d, set())
                new_files_per_dir[d] = sorted(after_set - before_set)
            # Agregar lista geral
            new_files = []
            for v in new_files_per_dir.values():
                new_files.extend(v)
            new_files = sorted(set(new_files))
            
            # Procurar arquivo .lis gerado (case-insensitive, .lis ou .LIS)
            lis_path = None
            candidates: List[Path] = [
                acp_path.with_suffix('.lis'),
                acp_path.with_suffix('.LIS')
            ]
            for c in candidates:
                if c.exists():
                    lis_path = c
                    break
            # Se não encontrado diretamente, procurar por mesmo stem nos diretórios monitorados
            if lis_path is None:
                try:
                    extra_search = list(search_dirs)  # já inclui run_cwd e pasta do .acp
                    sanitized_stem = Path(deck_in_solver.name).stem if deck_in_solver else acp_path.stem
                    for d in extra_search:
                        for p in list(d.glob('*.lis')) + list(d.glob('*.LIS')):
                            if p.stem.lower() in (acp_path.stem.lower(), sanitized_stem.lower()):
                                lis_path = p
                                break
                        if lis_path:
                            break
                except Exception:
                    pass

            # Verificar código de retorno do processo ATP
            log_entry = None
            logs_dir = effective_output_dir / 'logs'
            logs_dir.mkdir(parents=True, exist_ok=True)
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            log_path = logs_dir / f"{acp_path.stem}_{timestamp}.log"

            if result_returncode != 0:
                print(f"❌ ATP retornou código {result_returncode}. Considerando falha na simulação.")
                print(f"   Stdout: {result_stdout[:200]}")
                print(f"   Stderr: {result_stderr[:200]}")
                # Se um .lis foi gerado e NÃO está vazio, vamos aproveitar o resultado mesmo com timeout/erro
                effective_lis: Optional[Path] = None
                lis_size = 0
                if lis_path and lis_path.exists():
                    try:
                        lis_size = lis_path.stat().st_size
                    except Exception:
                        lis_size = 0
                    if lis_size > 0:
                        effective_lis = lis_path
                if effective_lis is not None:
                    # Mover sempre para o diretório de saída efetivo
                    effective_output_dir.mkdir(parents=True, exist_ok=True)
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    new_lis = effective_output_dir / f"{acp_path.stem}_{timestamp}.lis"
                    try:
                        shutil.move(effective_lis, new_lis)
                        effective_lis = new_lis
                    except Exception:
                        # Se mover falhar, manter original
                        pass
                    # Detectar arquivo .dbg correspondente e mover também
                    dbg_path = None
                    for dir_str, files in new_files_per_dir.items():
                        for fname in files:
                            if fname.lower().endswith('.dbg') and Path(fname).stem.lower() == acp_path.stem.lower():
                                candidate = Path(dir_str) / fname
                                if candidate.exists():
                                    dbg_path = candidate
                                    break
                        if dbg_path:
                            break
                    moved_dbg_path = None
                    if dbg_path:
                        try:
                            moved_dbg_path = effective_output_dir / f"{acp_path.stem}_{timestamp}.dbg"
                            shutil.move(dbg_path, moved_dbg_path)
                            dbg_path = moved_dbg_path
                        except Exception:
                            pass
                    # Log como "timeout_with_lis" e retornar caminho
                    try:
                        lines = [
                            "Status: timeout_with_lis",
                            f"Return code: {result_returncode}",
                            f"CWD: {run_cwd}",
                            f"Command: {' '.join(cmd)}",
                            f"New files: {', '.join(new_files) if new_files else '(none)'}",
                            "New files per directory:",
                        ]
                        for d, files in new_files_per_dir.items():
                            lines.append(f"  {d}: {', '.join(files) if files else '(none)'}")
                        lines.extend([
                            f"LIS: {effective_lis} (gerado apesar do erro/timeout; pode estar incompleto)",
                            f"DBG: {dbg_path if dbg_path else '(none)'}",
                            "---- STDOUT ----",
                            result_stdout or '(vazio)',
                            "---- STDERR ----",
                            result_stderr or '(vazio)'
                        ])
                        log_path.write_text('\n'.join(lines), encoding='utf-8')
                        print(f"📝 Log salvo em {log_path}")
                    except Exception as e:
                        print(f"⚠️ Falha ao salvar log: {e}")
                    # Limpeza automática de temporários (*.tmp, *.bin) gerados durante a simulação
                    removed_temp: List[Path] = []
                    for dir_str, files in new_files_per_dir.items():
                        base_dir = Path(dir_str)
                        for fname in files:
                            lower = fname.lower()
                            if lower.endswith('.tmp') or lower.endswith('.bin'):
                                fpath = base_dir / fname
                                try:
                                    fpath.unlink(missing_ok=True)
                                    removed_temp.append(fpath)
                                except Exception:
                                    pass
                    if removed_temp:
                        print(f"🧹 Temporários removidos: {', '.join(p.name for p in removed_temp)}")
                        try:
                            with log_path.open('a', encoding='utf-8') as lf:
                                lf.write('\nRemoved temps: ' + ', '.join(p.name for p in removed_temp) + '\n')
                        except Exception:
                            pass
                    # Limpar temporário .atp e retornar .lis aproveitado
                    temp_atp.unlink(missing_ok=True)
                    return effective_lis
                # Caso contrário: se existir .lis vazio, remover e registrar erro
                if lis_path and lis_size <= 0:
                    try:
                        lis_path.unlink(missing_ok=True)
                        print(f"🗑️  Arquivo .lis vazio removido: {lis_path}")
                    except Exception as e:
                        print(f"⚠️ Não foi possível remover .lis vazio: {e}")
                # Log erro padrão
                log_entry = {
                    'status': 'error',
                    'returncode': result_returncode,
                    'command': ' '.join(cmd),
                    'cwd': str(run_cwd),
                    'new_files': new_files,
                    'new_files_per_dir': new_files_per_dir,
                    'stdout': result_stdout,
                    'stderr': result_stderr
                }
                try:
                    lines = [
                        f"Status: {log_entry['status']}",
                        f"Return code: {log_entry['returncode']}",
                        f"CWD: {log_entry['cwd']}",
                        f"Command: {log_entry['command']}",
                        f"New files: {', '.join(log_entry['new_files']) if log_entry['new_files'] else '(none)'}",
                        "New files per directory:",
                    ]
                    for d, files in log_entry['new_files_per_dir'].items():
                        lines.append(f"  {d}: {', '.join(files) if files else '(none)'}")
                    lines.extend([
                        "---- STDOUT ----",
                        log_entry['stdout'] or '(vazio)',
                        "---- STDERR ----",
                        log_entry['stderr'] or '(vazio)'
                    ])
                    log_path.write_text('\n'.join(lines), encoding='utf-8')
                    print(f"📝 Log salvo em {log_path}")
                except Exception as e:
                    print(f"⚠️ Falha ao salvar log: {e}")
                return None
            
            if lis_path and lis_path.exists():
                # Validar tamanho do .lis (> 0 bytes) antes de mover
                try:
                    lis_size = lis_path.stat().st_size
                except Exception:
                    lis_size = 0
                
                if lis_size <= 0:
                    print("⚠️ .lis gerado, porém vazio (0 bytes). Considerando falha na simulação.")
                    print(f"   Stdout: {result_stdout[:200]}")
                    print(f"   Stderr: {result_stderr[:200]}")
                    # Remover arquivo vazio para evitar acúmulo
                    try:
                        lis_path.unlink(missing_ok=True)
                        print(f"🗑️  Arquivo .lis vazio removido: {lis_path}")
                    except Exception as e:
                        print(f"⚠️ Não foi possível remover .lis vazio: {e}")
                    # Log arquivo vazio
                    try:
                        lines = [
                            "Status: empty_lis",
                            f"Return code: {result_returncode}",
                            f"CWD: {run_cwd}",
                            f"Command: {' '.join(cmd)}",
                            f"New files: {', '.join(new_files) if new_files else '(none)'}",
                            "New files per directory:",
                        ]
                        for d, files in new_files_per_dir.items():
                            lines.append(f"  {d}: {', '.join(files) if files else '(none)'}")
                        lines.extend([
                            "---- STDOUT ----",
                            result_stdout or '(vazio)',
                            "---- STDERR ----",
                            result_stderr or '(vazio)'
                        ])
                        log_path.write_text('\n'.join(lines), encoding='utf-8')
                        print(f"📝 Log salvo em {log_path}")
                    except Exception as e:
                        print(f"⚠️ Falha ao salvar log: {e}")
                    return None
                
                # Mover sempre para o diretório de saída efetivo (ACP por padrão)
                effective_output_dir.mkdir(parents=True, exist_ok=True)
                
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                new_lis = effective_output_dir / f"{acp_path.stem}_{timestamp}.lis"
                
                shutil.move(lis_path, new_lis)
                lis_path = new_lis

                # Detectar arquivo .dbg correspondente e mover também
                dbg_path = None
                for dir_str, files in new_files_per_dir.items():
                    for fname in files:
                        if fname.lower().endswith('.dbg') and Path(fname).stem.lower() == acp_path.stem.lower():
                            candidate = Path(dir_str) / fname
                            if candidate.exists():
                                dbg_path = candidate
                                break
                    if dbg_path:
                        break
                moved_dbg_path = None
                if dbg_path:
                    try:
                        moved_dbg_path = effective_output_dir / f"{acp_path.stem}_{timestamp}.dbg"
                        shutil.move(dbg_path, moved_dbg_path)
                        dbg_path = moved_dbg_path
                    except Exception:
                        pass
                
                print(f"✅ Simulação concluída: {lis_path}")
                # Log sucesso
                try:
                    lines = [
                        "Status: success",
                        f"Return code: {result_returncode}",
                        f"CWD: {run_cwd}",
                        f"Command: {' '.join(cmd)}",
                        f"New files: {', '.join(new_files) if new_files else '(none)'}",
                        "New files per directory:",
                    ]
                    for d, files in new_files_per_dir.items():
                        lines.append(f"  {d}: {', '.join(files) if files else '(none)'}")
                    lines.extend([
                        f"LIS: {lis_path}",
                        f"DBG: {dbg_path if dbg_path else '(none)'}",
                        "---- STDOUT ----",
                        result_stdout or '(vazio)',
                        "---- STDERR ----",
                        result_stderr or '(vazio)'
                    ])
                    log_path.write_text('\n'.join(lines), encoding='utf-8')
                    print(f"📝 Log salvo em {log_path}")
                except Exception as e:
                    print(f"⚠️ Falha ao salvar log: {e}")

                # Limpeza automática de temporários (*.tmp, *.bin) gerados durante a simulação
                removed_temp: List[Path] = []
                for dir_str, files in new_files_per_dir.items():
                    base_dir = Path(dir_str)
                    for fname in files:
                        lower = fname.lower()
                        # Mantém .lis e .dbg; remove .tmp e .bin (scratch)
                        if lower.endswith('.tmp') or lower.endswith('.bin'):
                            fpath = base_dir / fname
                            try:
                                fpath.unlink(missing_ok=True)
                                removed_temp.append(fpath)
                            except Exception:
                                pass
                if removed_temp:
                    print(f"🧹 Temporários removidos: {', '.join(p.name for p in removed_temp)}")
                    # Acrescentar informação ao log (append)
                    try:
                        with log_path.open('a', encoding='utf-8') as lf:
                            lf.write('\nRemoved temps: ' + ', '.join(p.name for p in removed_temp) + '\n')
                    except Exception:
                        pass
                
                # Limpar arquivos temporários
                temp_atp.unlink(missing_ok=True)
                
                return lis_path
            else:
                print(f"⚠️ Simulação executada mas .lis não foi gerado")
                print(f"   Stdout: {result_stdout[:200]}")
                print(f"   Stderr: {result_stderr[:200]}")
                # Log ausência de .lis
                try:
                    log_path.write_text('\n'.join([
                        "Status: no_lis",
                        f"Return code: {result_returncode}",
                        f"CWD: {run_cwd}",
                        f"Command: {' '.join(cmd)}",
                        f"New files: {', '.join(new_files) if new_files else '(none)'}",
                        "---- STDOUT ----",
                        result_stdout or '(vazio)',
                        "---- STDERR ----",
                        result_stderr or '(vazio)'
                    ]), encoding='utf-8')
                    print(f"📝 Log salvo em {log_path}")
                except Exception as e:
                    print(f"⚠️ Falha ao salvar log: {e}")
                return None
        
        except subprocess.TimeoutExpired:
            print("❌ Timeout: simulação excedeu o tempo limite e foi interrompida")
            return None
        except Exception as e:
            print(f"❌ Erro ao executar ATP: {e}")
            return None
        finally:
            # Limpar arquivos temporários
            temp_atp.unlink(missing_ok=True)

    def _default_output_dir(self, acp_path: Path) -> Path:
        """Resolve diretório padrão de saída para .lis/.dbg.
        Preferir a pasta 'ACP' do projeto; se o arquivo estiver dentro dela, usar a própria.
        """
        try:
            # Se o .acp já está em uma pasta chamada ACP, usar essa pasta
            if acp_path and acp_path.parent.name.lower() == 'acp':
                return acp_path.parent
        except Exception:
            pass
        # Caso contrário, usar a pasta ACP ao lado deste script
        project_root = Path(__file__).parent
        return project_root / 'ACP'


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
