# 📊 Gerador de Relatório de Vendas em Excel

Script de automação que lê um arquivo **CSV de vendas**, **limpa os dados** e gera um **relatório Excel formatado**, com resumos e gráficos, em poucos segundos.

Feito com **Python, pandas e openpyxl**.

## 🎯 Problema que resolve

Montar relatório de vendas à mão no Excel é lento e sujeito a erro. Este projeto automatiza o processo: basta apontar o CSV e o relatório sai pronto.

## ✨ Funcionalidades

- Leitura e validação das colunas obrigatórias do CSV
- Limpeza automática dos dados:
  - remove linhas duplicadas
  - descarta linhas com data, quantidade ou preço inválidos
  - padroniza nomes de vendedores (`"  bruno "` → `"Bruno"`)
  - calcula o total de cada venda
- Relatório Excel com 5 abas:
  - **Resumo**: faturamento total, nº de vendas e ticket médio
  - **Por Produto**: quantidade e faturamento + gráfico
  - **Por Vendedor**: ranking de vendedores
  - **Mensal**: faturamento por mês + gráfico
  - **Dados Limpos**: base final tratada
- Formatação profissional (cabeçalho colorido, moeda em R$, colunas ajustadas, painel congelado)
- Interface de linha de comando (CLI)
- 5 testes automatizados

## 🛠️ Tecnologias

| Tecnologia | Uso |
|---|---|
| Python 3.10+ | Linguagem |
| pandas | Limpeza e agregação dos dados |
| openpyxl | Formatação do Excel e gráficos |
| argparse | Interface de linha de comando |
| pytest | Testes |

## 🚀 Como rodar

```bash
# 1. Clonar o repositório
git clone https://github.com/nathangmmelo/relatorio-vendas-excel.git
cd relatorio-vendas-excel

# 2. Criar e ativar o ambiente virtual
python -m venv .venv
source .venv/bin/activate        # Windows: .venv\Scripts\activate

# 3. Instalar as dependências
pip install -r requirements.txt

# 4. Rodar com o CSV de exemplo que já vem no projeto
python -m src.relatorio data/vendas_exemplo.csv
```

Saída esperada:

```
Linhas lidas: 303 | Linhas válidas: 301 | Descartadas: 2
Relatório gerado em: output/relatorio_vendas.xlsx
```

### Usando seus próprios dados

Seu CSV precisa ter estas colunas:

```csv
data,produto,vendedor,quantidade,preco_unitario
2026-01-15,Extintor ABC 4kg,Ana,2,120.0
```

```bash
python -m src.relatorio meus_dados.csv -o output/meu_relatorio.xlsx
```

Para gerar um novo CSV de exemplo: `python src/gerar_dados.py`

## 🧪 Testes

```bash
python -m pytest -v
```

## 📁 Estrutura

```
relatorio-vendas-excel/
├── data/
│   └── vendas_exemplo.csv    # dados fictícios (com sujeiras de propósito)
├── src/
│   ├── relatorio.py          # leitura, limpeza, resumos e geração do Excel
│   └── gerar_dados.py        # gerador do CSV de exemplo
├── tests/
│   └── test_relatorio.py
├── requirements.txt
└── README.md
```

## 🔮 Próximos passos

- Filtro por período (`--inicio` / `--fim`)
- Envio do relatório por e-mail
- Suporte a leitura direta de arquivos `.xlsx`

## 👤 Autor

**Nathan Gabriel Martins de Melo**
[GitHub](https://github.com/nathangmmelo)
