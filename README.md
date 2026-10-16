
# Projeto 01 — Análise de Vendas

## Objetivo
Desenvolver um projeto de análise de dados utilizando Python,
Pandas e Excel para investigar o desempenho das vendas.

## Tecnologias
- Python
- Pandas
- Openpyxl
- Matplotlib
- Git e GitHub

## Funcionalidades
- Importação de dados de uma planilha Excel.
- Cálculo do valor total das vendas.
- Agrupamento por vendedor, produto e período.
- Geração de tabelas-resumo.
- Criação de gráficos no Excel.

## Estrutura do projeto
```text
projeto_01_analise_vendas/
├── src/
│   └── analise_vendas.py
├── dados/
├── resultados/
├── .gitignore
└── README.md
```

## Como executar

1. Crie um ambiente virtual.
2. Instale as dependências:
   `pip install pandas openpyxl matplotlib`
3. Coloque o arquivo Excel na pasta `dados/`.
4. Execute:
   `python src/analise_vendas.py`

## Observação
Os resultados dependem da qualidade dos dados de entrada.
As duplicações identificadas na planilha original ainda
precisam ser investigadas antes de validar o faturamento.

## Status
Em desenvolvimento.

projeto_01_analise_vendas/
├── README.md
├── .gitignore
├── dados/
├── src/
│   └── analise_vendas.py
└── resultados/