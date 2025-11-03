# 🚀 Instruções para Configurar o ATP no LIS-ANALYSIS

## Problema Identificado

O erro `Permission denied: '/home/pedro/ATPDraw/Atpdraw.exe'` indica que:
1. O sistema está tentando executar um arquivo `.exe` do Windows no Linux
2. O arquivo não tem permissão de execução ou não pode ser executado diretamente

## Soluções Disponíveis

### ✅ Opção 1: Usar ATP Nativo para Linux (RECOMENDADO)

O ATP possui versões nativas para Linux que funcionam melhor:

```bash
# Instalar ATP para Linux (se disponível no seu repositório)
sudo apt update
sudo apt install atp

# OU baixar e instalar manualmente
# Visite: https://www.atpdraw.net/ ou fontes oficiais do ATP
```

**Vantagens:**
- Melhor performance
- Sem necessidade de Wine
- Mais estável

### ✅ Opção 2: Usar Wine para Executar Atpdraw.exe

Se você só tem a versão Windows do ATPDraw, pode usar Wine:

#### 1. Instalar Wine (com suporte 32-bit)

O Atpdraw.exe é um programa 32-bit, então precisa do wine32:

```bash
# Ubuntu/Debian - Habilitar arquitetura 32-bit
sudo dpkg --add-architecture i386

# Atualizar repositórios
sudo apt update

# Instalar Wine (Ubuntu 22.04+)
sudo apt install --install-recommends wine64 wine32:i386

# OU, se o comando acima não funcionar:
sudo apt install wine wine64 libwine:i386

# Verificar instalação
wine --version
```

**Nota**: Em versões modernas do Ubuntu, o `wine32` foi substituído por `wine32:i386` ou `libwine:i386`.

#### 2. Verificar Permissões do Atpdraw.exe

```bash
# Verificar se o arquivo existe
ls -la /home/pedro/ATPDraw/Atpdraw.exe

# Dar permissão de execução (se necessário)
chmod +x /home/pedro/ATPDraw/Atpdraw.exe
```

#### 3. Testar Execução Manual

```bash
# Tentar executar com Wine
dpkg --add-architecture i386 && apt-get update &&
apt-get install wine32"```

### ✅ Opção 3: Configurar Manualmente na Interface

Se você tiver o ATP instalado em outro local:

1. Abra a interface gráfica: `python3 main.py --gui`
2. Vá até a seção **"🎯 Controle Inteligente de Parâmetros"**
3. No campo **"Executável ATP:"**, clique em **"Escolher…"**
4. Selecione o executável correto:
   - Para Linux: `/usr/local/bin/tpbig` ou `/usr/bin/atp`
   - Para Wine: `/home/pedro/.wine/drive_c/ATP/Atpdraw.exe`

## Como Usar o Botão "🚀 Rodar ATP"

Após configurar corretamente:

1. **Selecione arquivos .acp** na lista de arquivos
2. Configure o **"Executável ATP"** (se ainda não configurou)
3. Clique no botão **"🚀 Rodar ATP"**
4. O sistema irá:
   - Extrair o conteúdo ATP do arquivo .acp
   - Executar a simulação
   - Gerar o arquivo .lis resultante
   - Salvar na pasta de saída configurada

## Verificação da Configuração

Execute o script de teste para verificar se tudo está configurado:

```bash
cd /home/pedro/vscode_ubuntu/LIS-ANALYSIS
python3 test_atp_runner.py
```

## Executáveis ATP Suportados

O sistema procura automaticamente por:

### Linux Nativo
- `/usr/local/bin/tpbig`
- `/usr/bin/tpbig`
- `/opt/atp/tpbig`
- `/usr/local/bin/atpmingw`
- `/usr/bin/atp`

### Wine (Windows via Linux)
- `~/ATPDraw/Atpdraw.exe` (com Wine)
- `~/.wine/drive_c/ATP/Atpdraw.exe` (com Wine)
- `/opt/atpdraw/Atpdraw.exe` (com Wine)

## Solução de Problemas

### Erro: "Package 'wine32' has no installation candidate"

Em versões modernas do Ubuntu, o pacote mudou de nome:

```bash
# Solução:
sudo dpkg --add-architecture i386
sudo apt update

# Tente uma destas opções:
# Opção 1 (recomendada):
sudo apt install --install-recommends wine64 wine32:i386

# Opção 2:
sudo apt install wine libwine:i386

# Opção 3 (se as anteriores falharem):
sudo apt install wine-stable winehq-stable
```

### Erro: "wine32 is missing" ou "not supported on this system"

Após instalar o Wine, se ainda aparecer esse erro:

```bash
# Verificar o que está faltando:
wine --version

# Reinstalar com dependências 32-bit:
sudo apt install --install-recommends wine64 wine32:i386

# Depois teste novamente:
wine /home/pedro/ATPDraw/Atpdraw.exe
```

### Erro: "Permission denied"
```bash
# Dar permissão de execução
chmod +x /caminho/para/executavel
```

### Erro: "Executável não encontrado"
- Verifique se o ATP está instalado
- Configure manualmente o caminho na interface
- Use o script de teste para verificar

### Erro: "Wine não encontrado"
```bash
# Instalar Wine
sudo apt install wine wine64
```

### Simulação não gera .lis
- Verifique se o arquivo .acp é válido
- Confira a saída do console para erros
- Teste executar o ATP manualmente

## Recursos Adicionais

- **Documentação ATP**: https://www.atpdraw.net/
- **Wine**: https://www.winehq.org/
- **Suporte**: Verifique os logs em `Simulation_Result/`

## Exemplo de Uso Completo

```bash
# 1. Instalar Wine (se necessário)
sudo apt install wine wine64

# 2. Dar permissão ao executável
chmod +x /home/pedro/ATPDraw/Atpdraw.exe

# 3. Executar interface
python3 main.py --gui

# 4. Configurar:
#    - Selecionar pasta com arquivos .acp
#    - Configurar "Executável ATP" para: wine /home/pedro/ATPDraw/Atpdraw.exe
#    - Selecionar arquivos .acp desejados
#    - Clicar em "🚀 Rodar ATP"
```

---

**Nota**: Para melhor experiência, recomenda-se usar a versão nativa do ATP para Linux ao invés da versão Windows via Wine.
