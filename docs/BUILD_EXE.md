# BUILD_EXE — Como gerar e executar o executável (Linux e Windows)

Este guia mostra, de forma direta e simples, como gerar e testar o executável do aplicativo `LIS-ANALYSIS`.

Importante:
- Você só precisa seguir os passos da sua plataforma (Linux ou Windows).
- Para distribuir o programa a usuários finais, gere o executável na mesma plataforma alvo (PyInstaller não recompila cross‑platform).

Requisitos (para quem vai COMPILAR o exe)
- Python 3.9+ instalado na máquina onde você vai rodar o script de build.
- No Linux: instale `python3-tk` se for usar a interface gráfica (Tkinter).
- Não é necessário Python para quem só VAI RODAR o `.exe` gerado.

Estrutura útil no repositório
- `scripts/build_linux.sh` — cria um ambiente isolado e gera o executável no Linux.
- `scripts/build_windows.bat` — gera o executável no Windows.
- `scripts/run_gui.bat` — helper que executa o `.exe` com console aberto (útil para ver erros ao testar no Windows).
- `run.py` — launcher local simples para executar a aplicação sem configurar `PYTHONPATH` manualmente.
- Resultado da build: `dist/LIS-ANALYSIS/` contendo o binário `LIS-ANALYSIS` (Linux) ou `LIS-ANALYSIS.exe` (Windows).

Execução local simplificada (sem empacotar)

```bash
python run.py --gui
```

Como compilar (passo a passo)

Windows (PowerShell ou Prompt de Comando)
1. Abra PowerShell no diretório do projeto.
2. Execute:

```bat
scripts\build_windows.bat
```


3. O executável estará em:

```bat
dist\LIS-ANALYSIS\LIS-ANALYSIS.exe
```

Linux (WSL ou máquina Linux)
1. Abra um terminal e vá para a pasta do projeto.
2. Execute:

```bash
./scripts/build_linux.sh
```

3. Após terminar, o executável estará em:

```bash
dist/LIS-ANALYSIS/LIS-ANALYSIS
```

Como TESTAR o executável

- No Windows: entre na pasta `dist\LIS-ANALYSIS` e:
  - Dê duplo-clique no `LIS-ANALYSIS.exe` para abrir a GUI (deveria funcionar).
  - Se nada abrir ao duplo-clique, execute `run_gui.bat` (na mesma pasta) para abrir um console que mostra mensagens e erros:

```bat
run_gui.bat
```

- No Linux: abra um terminal, entre em `dist/LIS-ANALYSIS` e rode:

```bash
./LIS-ANALYSIS --gui
```

Notas rápidas e solução de problemas
- Build deve ser feita na mesma plataforma do alvo (Windows build no Windows, Linux build no Linux).
- Se a GUI não abrir ao dar duplo-clique, use `run_gui.bat` para ver erros.
- Caso falte alguma dependência ao compilar, a mensagem no terminal indicará o pacote que falta; normalmente é só instalar com `pip`.
- `--onedir` (padrão aqui) gera uma pasta com o executável e arquivos auxiliares — mais fácil para depurar.
- `--onefile` gera um único `.exe` mas pode ter inicialização mais lenta; use só depois que a versão `--onedir` funcionar.

Higiene do Git (artefatos gerados)
- Para verificar artefatos rastreados por engano sem alterar nada:

```bash
./scripts/cleanup_git_artifacts.sh --dry-run
```

- Para desrastrear artefatos gerados (sem apagar arquivos locais):

```bash
./scripts/cleanup_git_artifacts.sh --apply
```

Se quiser que eu gere o executável para você e envie o artefato (via CI), diga qual opção prefere: gerar no Windows (exe) ou no Linux (binário). 

---
Arquivo de referência: `scripts/build_linux.sh`, `scripts/build_windows.bat`, `scripts/run_gui.bat`.
