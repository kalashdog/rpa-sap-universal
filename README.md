# RPA SAP Universal - Multi-Plantas

Bem-vindo ao repositório do **RPA SAP Universal**, um robô de automação altamente estruturado em Python, construído para orquestrar e extrair relatórios (Jobs) do sistema ERP SAP.

Esta ferramenta não apenas clica em botões do `SAP Logon` (GUI), mas faz um gerenciamento resiliente das operações em dezenas de transações (`MB51`, `LT23`, `VL06I`, entre outras) via **PythonCOM** (`win32com.client`), evitando falhas de interação ou vazamentos de memória na infraestrutura do Windows.

## A Arquitetura

1. **Agnóstico a Filiais/Plantas:** Em vez de duplicações massivas de código (ex: `script_anchieta`, `script_taubate`), o robô possui funções centrais (Handlers) enxutas no `transactions/request.py`. Toda a lógica de variação por Planta (como depósito "LGNUM", tipo "LGTYP" ou até seleção condicional de tela "radio") reside no **`config/sapscripts_config.json`**.
   *(Você adiciona plantas sem tocar no interpretador Python!)*
2. **Watchdog de Infraestrutura (`watchdog.py`)**: Ele monitora os processos dependentes (OneDrive para tráfego em nuvem silencioso), além de proteger contra vazamento de memória recriando as sessões limpas automaticamente.
3. **O Robusto "Extrator SP02" e o truque do XXL (`export.py` / `requests.py`)**: O RPA aguarda a concorrência assíncrona do SAP na impressora virtual (SP02) e extrai com timeout seguro. Ele consegue exportar listas dinâmicas em `XXL` de forma "Invisível".
4. **Integração Real-Time com Power Automate**: Um progress bar autêntico! O orquestrador tem a flag de erros e acusa seu pipeline vivo via HTTPS, atualizando as 3 fases de andamento para o Dashboard Web Front-end de gestão em tempo real.

## Como Executar
**Pré-requisitos:**
* Ambiente Windows (`*.exe` via OneDrive Local ou C:\).
* SAP Logon aberto (ou auto recriável via configuração do Watchdog) e credenciais conectadas ou engatilhadas de single-sign-on (SSO).
* O instalador `requirements.txt` atualizado (`pip install -r requirements.txt`).

1. Para inicializar a máquina de trabalho para uma **Planta Específica** (ex: Taubaté), rode o script global declarando seu ID:

```powershell
python main.py --plant "02-Taubate"
```

## Como adicionar Novas Plantas? 

Basta popular a sessão correspondente no `sapscripts_config.json`.
Veja um exemplo adicionando São Carlos à extração de Fifo:
```json
        "04-SaoCarlos": {
          "lgnum": "SCA",
          "variant": "/LT23SCA",
          "local_extract": "002 - Quebra de Fifo\\LT23",
          "name_file": "LT23-{date:%d-%m-%Y}.txt"
        }
```

O orquestrador varrerá as parametrizações `plant_params`, acionará as transações corretas, fará os inputs via dicionário e extrairá de forma cíclica sem intervenção humana.

---
**Autor:** Vinicius Lima