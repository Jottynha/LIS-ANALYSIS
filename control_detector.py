"""
Detector e parser de parametros de controle em nomes de arquivos ATP.
Identifica RPI, RF e outros parametros e permite modificacao dinamica.
"""

import re
from pathlib import Path
from typing import Dict, List, Optional, Tuple
from dataclasses import dataclass

@dataclass
class ControlParameter:
    """Representa um parâmetro de controle detectado"""
    name: str  # 'RPI', 'RF', etc
    value: float  # Valor numérico
    unit: str  # 'Ω', 'Ω', etc
    position_in_name: Tuple[int, int]  # Posição no nome do arquivo
    pattern_matched: str  # Padrão que foi encontrado

@dataclass
class FileControlInfo:
    """Informações de controle extraídas de um arquivo"""
    original_path: Path
    base_name: str  # Nome sem extensão
    has_control: bool  # False se "SemControle"
    parameters: List[ControlParameter]
    file_type: str  # 'CONVENCIONAL', 'OTIMIZADA', etc
    
    def __str__(self):
        if not self.has_control:
            return f"{self.original_path.name} [SEM CONTROLE]"
        
        params_str = ", ".join([f"{p.name}={p.value}{p.unit}" for p in self.parameters])
        return f"{self.original_path.name} [{params_str}]"


class ControlDetector:
    """Detector de parâmetros de controle em nomes de arquivos"""
    
    # Padrões de regex para detectar parâmetros
    PATTERNS = {
        'RPI': [
            r'RPI\s*=\s*(\d+(?:\.\d+)?)',  # RPI=100
            r'RPI(\d+)',  # RPI100
            r'Rpi\s*=\s*(\d+(?:\.\d+)?)',  # Rpi=100
        ],
        'RF': [
            r'RF\s*=\s*(\d+(?:\.\d+)?)',  # RF=30
            r'RF(\d+)',  # RF30
            r'Rf\s*=\s*(\d+(?:\.\d+)?)',  # Rf=30
        ],
        'RCRIT': [
            r'RCRIT\s*=\s*(\d+(?:\.\d+)?)',  # RCRIT=50
            r'Rcrit\s*=\s*(\d+(?:\.\d+)?)',
        ],
        'TCRIT': [
            r'TCRIT\s*=\s*(\d+(?:\.\d+)?)',  # TCRIT=0.01
            r'Tcrit\s*=\s*(\d+(?:\.\d+)?)',
        ],
    }
    
    # Unidades padrão para cada parâmetro
    UNITS = {
        'RPI': 'Ω',
        'RF': 'Ω',
        'RCRIT': 'Ω',
        'TCRIT': 's',
    }
    
    # Descrições amigáveis
    DESCRIPTIONS = {
        'RPI': 'Resistência de Pré-Inserção',
        'RF': 'Resistor de Falta',
        'RCRIT': 'Resistência Crítica',
        'TCRIT': 'Tempo Crítico',
    }
    
    @staticmethod
    def detect_from_file(file_path: Path) -> FileControlInfo:
        """
        Detecta parâmetros de controle a partir do nome do arquivo.
        
        Args:
            file_path: Caminho do arquivo .lis
            
        Returns:
            FileControlInfo com todos os parâmetros detectados
        """
        file_path = Path(file_path)
        base_name = file_path.stem
        
        # Verificar se é "Sem Controle"
        has_control = not bool(re.search(r'sem\s*controle', base_name, re.IGNORECASE))
        
        # Detectar tipo (CONVENCIONAL, OTIMIZADA, etc)
        file_type = 'UNKNOWN'
        if 'convenc' in base_name.lower():
            file_type = 'CONVENCIONAL'
        elif 'otimizada' in base_name.lower():
            file_type = 'OTIMIZADA'
        elif 'hibrida' in base_name.lower():
            file_type = 'HÍBRIDA'
        
        parameters = []
        
        if has_control:
            # Tentar detectar cada tipo de parâmetro
            for param_name, patterns in ControlDetector.PATTERNS.items():
                for pattern in patterns:
                    match = re.search(pattern, base_name, re.IGNORECASE)
                    if match:
                        value = float(match.group(1))
                        unit = ControlDetector.UNITS.get(param_name, '')
                        
                        param = ControlParameter(
                            name=param_name,
                            value=value,
                            unit=unit,
                            position_in_name=match.span(),
                            pattern_matched=match.group(0)
                        )
                        parameters.append(param)
                        break  # Pegar apenas primeira ocorrência
        
        return FileControlInfo(
            original_path=file_path,
            base_name=base_name,
            has_control=has_control,
            parameters=parameters,
            file_type=file_type
        )
    
    @staticmethod
    def detect_from_files(file_paths: List[Path]) -> List[FileControlInfo]:
        """Detecta parâmetros de múltiplos arquivos"""
        return [ControlDetector.detect_from_file(fp) for fp in file_paths]
    
    @staticmethod
    def generate_new_filename(info: FileControlInfo, new_params: Dict[str, float]) -> str:
        """
        Gera novo nome de arquivo com parâmetros modificados.
        
        Args:
            info: Informações originais do arquivo
            new_params: Dicionário com novos valores {param_name: new_value}
            
        Returns:
            Novo nome do arquivo
        """
        new_name = info.base_name
        
        # Substituir cada parâmetro
        for param in info.parameters:
            if param.name in new_params:
                new_value = new_params[param.name]
                # Substituir valor no nome
                new_pattern = param.pattern_matched.replace(
                    str(int(param.value)), 
                    str(int(new_value))
                )
                new_name = new_name.replace(param.pattern_matched, new_pattern)
        
        return new_name + info.original_path.suffix
    
    @staticmethod
    def get_parameter_description(param_name: str) -> str:
        """Retorna descrição amigável do parâmetro"""
        return ControlDetector.DESCRIPTIONS.get(param_name, param_name)
    
    @staticmethod
    def suggest_values(param_name: str, current_value: float) -> List[float]:
        """
        Sugere valores típicos para um parâmetro.
        
        Args:
            param_name: Nome do parâmetro
            current_value: Valor atual
            
        Returns:
            Lista de valores sugeridos
        """
        if param_name == 'RPI':
            # Sugerir valores de 100 a 1000 em incrementos de 100
            return [100, 200, 300, 400, 500, 600, 700, 800, 900, 1000]
        
        elif param_name == 'RF':
            # Valores típicos de RF
            return [10, 20, 30, 40, 50, 60, 70, 80, 90, 100]
        
        elif param_name == 'RCRIT':
            return [10, 25, 50, 75, 100, 150, 200]
        
        elif param_name == 'TCRIT':
            return [0.001, 0.005, 0.01, 0.02, 0.05, 0.1]
        
        return [current_value]


def analyze_workspace_files(folder: Path, extensions: List[str] = ['.lis']) -> Dict[str, List[FileControlInfo]]:
    """
    Analisa todos os arquivos de uma pasta e organiza por tipo de controle.
    
    Args:
        folder: Pasta para analisar
        extensions: Extensões de arquivo a considerar
        
    Returns:
        Dict organizado: {'COM_CONTROLE': [...], 'SEM_CONTROLE': [...]}
    """
    folder = Path(folder)
    
    files = []
    for ext in extensions:
        files.extend(folder.rglob(f'*{ext}'))
        files.extend(folder.rglob(f'*{ext.upper()}'))
    
    infos = ControlDetector.detect_from_files(files)
    
    result = {
        'COM_CONTROLE': [],
        'SEM_CONTROLE': [],
        'POR_TIPO': {}
    }
    
    for info in infos:
        if info.has_control:
            result['COM_CONTROLE'].append(info)
        else:
            result['SEM_CONTROLE'].append(info)
        
        # Organizar por tipo
        if info.file_type not in result['POR_TIPO']:
            result['POR_TIPO'][info.file_type] = []
        result['POR_TIPO'][info.file_type].append(info)
    
    return result


# ==================== EXEMPLO DE USO ====================

if __name__ == "__main__":
    print("="*80)
    print("🔍 TESTE DO DETECTOR DE PARÂMETROS DE CONTROLE")
    print("="*80)
    
    # Testes com nomes de arquivos
    test_files = [
        "Caso0_ReEnergizacao_Convenc_RPI=100 e RF=30.LIS",
        "Caso0_ReEnergizacao_Convenc_SemControle.LIS",
        "Caso0_ReEnergizacao_OTIMIZADA_RPI=500 e RF=30.lis",
        # TODO: adicionar testes com .atp quando a integracao ATP for reescrita
        "Simulacao_RCRIT=50_TCRIT=0.01.lis"
    ]
    
    for filename in test_files:
        print(f"\n📄 Arquivo: {filename}")
        print("-" * 80)
        
        info = ControlDetector.detect_from_file(Path(filename))
        
        print(f"   Tipo: {info.file_type}")
        print(f"   Tem controle: {'✅ Sim' if info.has_control else '❌ Não'}")
        
        if info.parameters:
            print(f"   Parâmetros detectados: {len(info.parameters)}")
            for param in info.parameters:
                desc = ControlDetector.get_parameter_description(param.name)
                print(f"      • {param.name} ({desc}): {param.value} {param.unit}")
                
                # Sugerir novos valores
                suggestions = ControlDetector.suggest_values(param.name, param.value)
                print(f"        Sugestões: {suggestions[:5]}")
            
            # Testar geração de novo nome
            new_params = {}
            for param in info.parameters:
                if param.name == 'RPI':
                    new_params['RPI'] = 250.0
                elif param.name == 'RF':
                    new_params['RF'] = 45.0
            
            if new_params:
                new_filename = ControlDetector.generate_new_filename(info, new_params)
                print(f"\n   📝 Novo nome com parâmetros modificados:")
                print(f"      {filename} → {new_filename}")
    
    # Analisar workspace real
    print("\n" + "="*80)
    print("📁 ANÁLISE DO WORKSPACE")
    print("="*80)
    
    workspace = Path("/home/joao/Projetos/6º Periodo/Eletrônica/LIS-ANALYSIS")
    
    if workspace.exists():
        analysis = analyze_workspace_files(workspace)
        
        print(f"\n✅ Arquivos COM controle: {len(analysis['COM_CONTROLE'])}")
        for info in analysis['COM_CONTROLE'][:5]:  # Mostrar até 5
            params_str = ", ".join([f"{p.name}={p.value}" for p in info.parameters])
            print(f"   • {info.original_path.name} [{params_str}]")
        
        print(f"\n❌ Arquivos SEM controle: {len(analysis['SEM_CONTROLE'])}")
        for info in analysis['SEM_CONTROLE']:
            print(f"   • {info.original_path.name}")
        
        print(f"\n📊 Por tipo:")
        for tipo, infos in analysis['POR_TIPO'].items():
            if tipo != 'UNKNOWN':
                print(f"   • {tipo}: {len(infos)} arquivo(s)")
