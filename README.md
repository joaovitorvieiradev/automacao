# Automação em Python 
# 🤖 Automação - Mapa de Alunos

Este repositório contém um projeto de **automação em Python** desenvolvido para simplificar um processo repetitivo de extração, transformação e inserção de dados em uma planilha chamada **MAPA_DE_ALUNOS**.

---

## 🎯 Objetivo

Automatizar a rotina de:

1. 📥 Coletar dados de um site da empresa automaticamente.
2. 🔄 Transformar esses dados em uma base estruturada em Excel.
3. 📊 Criar uma tabela dinâmica.
4. 📌 Colar essa tabela em uma célula específica da planilha final.

Tudo isso com o uso de **Python** e bibliotecas especializadas em automação, manipulação de arquivos e análise de dados.

---

## ⚠️Todos os dados presentes no arquivo são fictícios e foram criados para fins de teste.⚠️

---

## 🧠 Como Funciona

O projeto está dividido em **subpastas** para manter o código organizado por etapas do processo:

### 📁 `automacao/`

- `extrair_dados.py`: Código que usa Selenium para acessar o site da empresa, baixar os dados e salvá-los como HTML.
- `filtros_selecionados.txt`: Arquivo que armazena filtros ou parâmetros usados durante a automação.

### 📁 `tratamento_de_planilhas/`

- `Arquivo_HTML/`: Pasta onde o HTML baixado é salvo.
- `BaseDeDados.xlsx`: Arquivo gerado a partir do HTML com os dados limpos.
- `MAPA_DE_ALUNOS.XLSX`: Planilha final que recebe a tabela dinâmica.
- `transformar_em_planilha.py`: Script que transforma o HTML em uma planilha estruturada.
- `tratamento_de_BaseDeDados`: Código responsável por aplicar filtros e criar a tabela dinâmica.

### 📄 `interface.py`

Interface principal que conecta e organiza a execução de todo o processo de forma automatizada.

---

## 📂 Estrutura do Projeto

ALTOMAÇÃO_MAPA_DE_ALUNOS/
│
├── automacao/
│ │
│ ├── extrair_dados.py
│ │
│ └── filtros_selecionados.txt
│
├── tratamento_de_planilhas/
│ │
│ ├── Arquivo_HTML/
│ │
│ ├── BaseDeDados.xlsx
│ │
│ ├── MAPA_DE_ALUNOS.XLSX
│ │
│ ├── transformar_em_planilha.py
│ │
│ └── tratamento_de_BaseDeDados
│
├── interface.py
│
└── README.md

---

## 🛠️ Tecnologias e Bibliotecas Utilizadas

- 🧭 **Selenium** – Automação de navegador para extração de dados.
- 📂 **os** – Manipulação de arquivos e diretórios.
- 📊 **pandas** – Análise e transformação de dados.
- 📈 **openpyxl** – Manipulação de arquivos Excel (.xlsx).

Instale as dependências com:

```bash
pip install selenium pandas openpyxl
