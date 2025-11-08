# Eletronic Invoice Launching System

## :bulb: About

- An electronic invoice launching system for the company I work for.

- Where a .xlsx file ensures data integrity and keeps the client list updated.

- It uses Pandas for data manipulation and Selenium for browser automation, with event waits to ensure smooth project flow.

- The system is active and has been used since 2021 every end of month.

## :wrench: Setup

Este repositório não comita o ambiente virtual (`.venv`) nem o binário do GeckoDriver.

Requisitos mínimos: Python 3.8+.

1) Criar um ambiente virtual e instalar dependências

Windows (PowerShell):
```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python automate.py
```

Linux / macOS (bash):
```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python automate.py
```

2) Variáveis de ambiente

O projeto lê as variáveis `LOGIN`, `PASSWORD` e `PATH_URL` a partir de um arquivo `.env` (usando `python-dotenv`). Crie um arquivo `.env` na raiz com o conteúdo:

```
LOGIN=seu_login
PASSWORD=sua_senha
PATH_URL=https://exemplo.com
```

3) GeckoDriver

O projeto espera o binário `geckodriver` presente na raiz (ou ajuste o caminho em `automate.py`). No Linux garanta que o arquivo `geckodriver` seja executável:

```bash
chmod +x geckodriver
```

5) Engine do Excel

O script usa `pandas.read_excel` para carregar `boleto.xlsx`. Para garantir compatibilidade com arquivos .xlsx modernos, `openpyxl` foi adicionada às dependências e está listada em `requirements.txt`.

4) Notas

- Se faltar alguma dependência durante a execução, execute `pip install <pacote>` e atualize `requirements.txt` com `pip freeze > requirements.txt` se desejar fixar versões.
- Para rodar uma verificação rápida de sintaxe: `python -m py_compile automate.py`.

