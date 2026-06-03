# Como gerar o executável (.exe)

Esse guia mostra como empacotar o Reiniciador de Workflow - iiLex (RPA) num **`.exe` único**
pra distribuir pros outros usuários do escritório sem precisar instalar
Python.

## Passo a passo

### 1. Instale as dependências do projeto
```cmd
py -m pip install -r requirements.txt
```

### 2. Instale o PyInstaller
```cmd
py -m pip install pyinstaller
```

### 3. Rode o build
Duas formas equivalentes:

**Opção A — clique duplo:** dá dois cliques em `build.bat`.

**Opção B — manual:**
```cmd
py -m PyInstaller iilex.spec --clean
```

### 4. Pegue o executável
O arquivo final fica em:
```
dist/Reiniciador de Workflow - iiLex (RPA).exe
```

Esse `.exe` é **autocontido** — basta copiar pra outra máquina Windows
e executar. ~80–100 MB.

---

## Distribuindo para outros usuários

Copie pra um pendrive / rede compartilhada / OneDrive:
```
Reiniciador de Workflow - iiLex (RPA).exe
```

O usuário só dá dois cliques. As configurações dele são salvas em:
- `%USERPROFILE%\.iilex_login.dat`  ← credenciais (DPAPI)
- `%USERPROFILE%\.iilex_config.json` ← preferências
- `%USERPROFILE%\.iilex_checkpoint.json` ← retomar (criado/apagado automaticamente)

---

## Solução de problemas

### "Windows Defender bloqueou o app"
PyInstaller gera executáveis sem assinatura digital, e o SmartScreen
pode bloquear na primeira execução. Solução:
1. Clique em **"Mais informações"** na janela do SmartScreen
2. Clique em **"Executar assim mesmo"**

Pra resolver de vez, seria preciso comprar um **certificado de
assinatura de código** (~R$ 1.000 / ano).

### "Falha ao baixar o ChromeDriver"
O executável tenta baixar o ChromeDriver na primeira execução. Se a
rede do escritório bloquear, baixe manualmente:
- https://googlechromelabs.github.io/chrome-for-testing/
- Coloque `chromedriver.exe` na mesma pasta do `.exe`

### "Antivirus detectou ameaça"
Falso positivo comum com PyInstaller. Adicione exceção:
```
Configurações → Vírus e proteção → Gerenciar configurações →
Adicionar ou remover exclusões → Reiniciador de Workflow - iiLex (RPA).exe
```

### Executável muito grande (~150 MB)
Já estamos excluindo libs grandes (tkinter, numpy, etc.). Pra reduzir
mais, considere usar UPX (instale `upx-ucl` e ele será aplicado
automaticamente pelo spec).

---

## Versionamento

Para atualizar a versão futuramente:
1. Edite o código
2. Rode `build.bat` de novo
3. Distribua o novo `.exe` substituindo o antigo
4. As configurações dos usuários são preservadas (ficam no `%USERPROFILE%`)
