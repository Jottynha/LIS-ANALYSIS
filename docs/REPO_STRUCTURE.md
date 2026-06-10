# Estrutura do repositorio

Este projeto foi reorganizado para separar responsabilidades por pasta:

- `src/`: codigo-fonte da aplicacao (`lis_analysis`).
- `data/`: insumos de referencia e amostras opcionais.
- `docs/`: documentacao do projeto e guias de build.
- `scripts/`: automacoes (build e limpeza de artefatos Git).
- `tests/`: testes automatizados.

## Estrutura atual

```text
src/lis_analysis/...
data/samples/ACP/...
data/samples/Listas ATP/...
docs/BUILD_EXE.md
scripts/*.sh|*.bat
tests/*
```

## Politica de dados

- Amostras grandes e opcionais ficam em `data/samples/`.
- Se o volume crescer, o recomendado e mover `data/samples/` para um repositorio separado e manter aqui apenas um subconjunto minimo para testes.
- Artefatos gerados em execucao (logs, xlsx, png e outputs de simulacao) nao devem ser versionados.

## Execucao local

- Comando simples (recomendado):

```bash
python run.py --gui
```
