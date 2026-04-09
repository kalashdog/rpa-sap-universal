# Hub Sesé • RPA Logística Universal (SAP)

![Python](https://img.shields.io/badge/Python-3.11-blue?logo=python&logoColor=white)
![CustomTkinter](https://img.shields.io/badge/GUI-CustomTkinter-darkblue)
![SAP](https://img.shields.io/badge/ERP-SAP_Logon-orange?logo=sap&logoColor=white)
![CI/CD](https://img.shields.io/badge/CI%2FCD-OneDrive_Offline-success)

Bem-vindo ao **Hub Sesé RPA**, uma solução de automação corporativa de alto desempenho construída em Python. 
Esta ferramenta orquestra, extrai e distribui dezenas de relatórios (Jobs) do sistema ERP SAP para múltiplas plantas simultaneamente, contando com uma interface gráfica moderna e um sistema de distribuição contínua (CI/CD) 100% offline.

![Demonstração da Interface](.assets/demo_interface.gif)

## A Nova Arquitetura

O robô evoluiu de um script de terminal para um ecossistema de software completo, sustentado por 3 pilares:

1. **GUI Moderna (Modo Insónia):** Desenvolvido com `CustomTkinter`, o robô possui uma interface limpa em Dark Mode, barra de progresso em tempo real, integração nativa com o cofre de senhas do Windows (Keyring) e injeção de kernel (`SetThreadExecutionState`) para impedir que o PC entre em suspensão durante extrações longas.
2. **Agnóstico a Filiais/Plantas (Data-Driven):** Toda a lógica de variação por Planta (como depósito `LGNUM` ou parâmetros de transação) reside puramente no arquivo **`config/sapscripts_config.json`**. Zero necessidade de duplicar código em Python para escalar para novas filiais.
3. **Distribuição Contínua (Bypass de Firewall):** A implementação de atualizações é feita via um `Launcher` inteligente. Ele consulta um Oráculo JSON hospedado no OneDrive corporativo, faz o bypass de firewalls restritivos e atualiza o executável em *background* (`AppData`) sem que o operador perceba, garantindo versão única em todas as fábricas.

![Arquitetura de Deploy](.assets/deploy_architecture.png)

## Como Executar (Para Operadores)

O operador final não precisa de Python instalado. O processo foi reduzido a **fricção zero**:
1. O utilizador clica no ícone **RPA Sesé** na sua Área de Trabalho.
2. O Launcher verifica silenciosamente a pasta `.rpa_update` no SharePoint/OneDrive local.
3. Se houver versão nova, atualiza e abre; caso contrário, abre imediatamente o painel.

## Como Lançar Novas Versões (Para Desenvolvedores)

O pipeline de compilação foi automatizado para 1 clique. Esqueça o `auto-py-to-exe` e configurações manuais.

1. Faça suas alterações no código (ex: `gui.py` ou `extract.py`).
2. Execute o script de build:
   ```cmd
   lançar_versao.bat

   ```
3. O assistente de linha de comando perguntará o número da nova versão e o que mudou. Ele ativará o ambiente virtual isolado, compilará o robô otimizado (~25MB), injetará no OneDrive da empresa, limpará os arquivos temporários e atualizará o Oráculo (`update_info.json`). 

## Como adicionar Novas Plantas ou Transações? 

Basta popular a sessão correspondente no `sapscripts_config.json`.
Veja um exemplo prático adicionando a planta "São Carlos" à extração de Filas (Fifo):

```json
"04-SaoCarlos": {
    "lgnum": "SCA",
    "variant": "/LT23SCA",
    "local_extract": "002 - Quebra de Fifo\\LT23",
    "name_file": "LT23-{date:%d-%m-%Y}.txt"
}
```

O core dinâmico do orquestrador varrerá as parametrizações, acionará as transações corretas, fará os inputs e extrairá de forma cíclica via **PythonCOM** sem intervenção humana.

---
**Autor:** Vinicius Lima