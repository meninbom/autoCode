![Python](https://img.shields.io/badge/python-3.10+-blue)
--- 
# Gerador de Links SGD

## Visão Geral

**Gerador de Links SGD** é uma aplicação desktop em Python desenvolvida para automatizar o processo de login no sistema SGD, aplicar filtros por data e responsável, baixar relatórios em Excel e gerar links personalizados com base nos dados do relatório.

A aplicação possui uma interface gráfica moderna construída com **customtkinter**, opera em **modo escuro** e é totalmente responsiva. O arquivo Excel baixado é processado diretamente — com remoção de colunas, geração de links e formatação — tudo em um fluxo de trabalho de **arquivo único**.

Ideal para usuários que desejam agilizar tarefas repetitivas no sistema SGD, garantindo consistência nos dados e economia de tempo por meio da automação.

## Funcionalidades

### Automação e Navegação

- Login automático no sistema SGD.
- Navegação até a página de programações.
- Aplicação de filtros por data.

### Processamento de Relatórios Excel

- Download de relatórios via `requests`, utilizando os cookies do Selenium para autenticação.
- Filtro por "Responsável" selecionado.
- Remoção de colunas desnecessárias (ex: "Data de entrada", "Unidade" etc).
- Geração de coluna de **links** com base nos códigos da coluna "Número".
- Ajuste automático da largura das colunas e aplicação de bordas com `openpyxl`.
- **Todas as alterações são feitas diretamente no arquivo baixado**, sem criar um novo.

### Interface Gráfica

- Construída com `customtkinter`, em **modo escuro**.
- Menu suspenso para selecionar o "Responsável".
- Barra de progresso e mensagens de status durante o processamento.

### Gerenciamento de Arquivos

- Renomeia o arquivo baixado no formato `YYYY-MM-DD.xlsx`, adicionando sufixos (_1, _2, etc.) se necessário.
- Caminho do último arquivo processado salvo em `config.json`.

### Persistência de Configurações

- Entradas do usuário (usuário, senha, diretório, responsável, datas) são salvas automaticamente em `config.json`.
- As configurações são carregadas automaticamente ao iniciar.

### Logs e Feedback

- Registro detalhado de todas as etapas com o módulo `logging` do Python.
- Atualizações de progresso exibidas na interface com callbacks.

## Requisitos

- Python 3.8 ou superior
- Navegador compatível (recomenda-se Google Chrome para uso com Selenium)

## Dependências

As bibliotecas utilizadas são:

- `customtkinter`
- `selenium`
- `webdriver-manager`
- `pandas`
- `openpyxl`
- `requests`
- `tkinter` (já incluída no Python)

Instale as dependências com:

```bash
pip install customtkinter selenium webdriver-manager pandas openpyxl requests
```

## Instalação

### Clonar o Repositório

```bash
git clone <url-do-repositório>
cd gerador-de-links-sgd
```

### Instalar Dependências

```bash
pip install -r requirements.txt
```

Se o arquivo `requirements.txt` não estiver presente, use o comando listado na seção de dependências.

### Executar a Aplicação

```bash
python main.py
```

Substitua `main.py` caso o nome do arquivo principal seja diferente.

## Executável

Já está disponível um executável `.exe` na página do GitHub, pronto para uso. Basta baixá-lo, executá-lo e começar a usar. O arquivo `config.json` será criado automaticamente na primeira execução, se não existir.

---

