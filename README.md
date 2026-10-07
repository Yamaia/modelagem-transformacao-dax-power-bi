# 🧮 Modelagem e Transformação de Dados com DAX no Power BI

Projeto de modelagem e transformação de dados no Power BI: tabelas calculadas, colunas calculadas e medidas em **DAX** sobre um modelo de vendas.

Desenvolvido como desafio de projeto da [DIO](https://www.dio.me/).

## 🎯 Objetivo

Estruturar um modelo de dados, transformar e enriquecer as informações com DAX e entregar indicadores prontos para análise.

## 🗂️ Estrutura do repositório

```
├── dados/
│   └── base-vendas.xlsx
├── relatorio/
│   └── modelagem-dax.pbix
├── dax/
│   └── medidas.dax              # Todas as expressões DAX do projeto
├── imagens/
│   ├── modelo.png
│   └── relatorio.png
└── README.md
```

## 🧱 Modelagem

- Importação e tratamento das tabelas no Power Query
- Tabelas fato e dimensão, com relacionamentos 1:N e filtro em direção única
- Tabela Calendário marcada como tabela de datas
- Colunas de chave ocultas e tabela de medidas dedicada

## 🔧 Transformações com DAX

### Tabela calculada: Calendário
```DAX
dim_Calendario =
ADDCOLUMNS(
    CALENDAR(MIN(fato_Vendas[Data]), MAX(fato_Vendas[Data])),
    "Ano", YEAR([Date]),
    "Mês", MONTH([Date]),
    "Nome do Mês", FORMAT([Date], "MMMM"),
    "Trimestre", "T" & QUARTER([Date]),
    "Dia da Semana", FORMAT([Date], "dddd")
)
```

### Colunas calculadas
```DAX
Valor Total = fato_Vendas[Quantidade] * fato_Vendas[Preco_Unitario]

Faixa de Venda =
SWITCH(
    TRUE(),
    fato_Vendas[Valor Total] >= 5000, "Alta",
    fato_Vendas[Valor Total] >= 1000, "Média",
    "Baixa"
)

Categoria Padronizada = UPPER(TRIM(dim_Produto[Categoria]))
```

### Medidas básicas
```DAX
Total Vendas = SUM(fato_Vendas[Valor Total])

Quantidade Vendida = SUM(fato_Vendas[Quantidade])

Lucro = SUM(fato_Vendas[Valor Total]) - SUM(fato_Vendas[Custo])

Margem % = DIVIDE([Lucro], [Total Vendas])

Ticket Médio = DIVIDE([Total Vendas], DISTINCTCOUNT(fato_Vendas[ID_Venda]))
```

### Medidas com CALCULATE e filtros
```DAX
Vendas Brasil = CALCULATE([Total Vendas], dim_Localidade[País] = "Brasil")

Vendas Sem Filtro de Categoria =
CALCULATE([Total Vendas], ALL(dim_Produto[Categoria]))

% do Total = DIVIDE([Total Vendas], CALCULATE([Total Vendas], ALL(dim_Produto)))
```

### Inteligência de tempo
```DAX
Vendas Ano Anterior =
CALCULATE([Total Vendas], SAMEPERIODLASTYEAR(dim_Calendario[Date]))

Variação % YoY =
DIVIDE([Total Vendas] - [Vendas Ano Anterior], [Vendas Ano Anterior])

Vendas Acumuladas no Ano = TOTALYTD([Total Vendas], dim_Calendario[Date])

Vendas Mês Anterior =
CALCULATE([Total Vendas], DATEADD(dim_Calendario[Date], -1, MONTH))
```

### Ranking e iteradores
```DAX
Ranking Produtos =
RANKX(ALL(dim_Produto[Produto]), [Total Vendas], , DESC, DENSE)

Receita (Iterador) =
SUMX(fato_Vendas, fato_Vendas[Quantidade] * fato_Vendas[Preco_Unitario])
```

> Ajuste nomes de tabelas e colunas conforme a base usada no desafio.

## 📊 Relatório

- Cartões: Total de Vendas, Lucro, Margem %, Variação YoY
- Gráfico de linhas: vendas atuais vs. ano anterior
- Gráfico de barras: ranking de produtos
- Segmentadores: Ano, Categoria, País

## 🖼️ Prints

![Modelo](imagens/modelo.png)
![Relatório](imagens/relatorio.png)

## 🌐 Relatório publicado

🔗 [Acessar no Power BI Service](COLE_AQUI_O_LINK)

## 💡 Aprendizados

- Diferença entre colunas calculadas e medidas
- Contexto de linha e contexto de filtro
- Uso de `CALCULATE`, `ALL` e iteradores (`SUMX`, `RANKX`)
- Inteligência de tempo com tabela Calendário
- Organização do modelo e das medidas

## 👤 Autor

**Yamaia** — desenvolvedor focado em backend, cloud e arquitetura de soluções.

[GitHub](https://github.com/SEU_USUARIO) · [LinkedIn](https://linkedin.com/in/SEU_PERFIL)
