# Price History — histórico de preços de um fornecedor

Ferramenta autónoma que lê o CSV de preços de um fornecedor, gera um Excel formatado,
**deteta mudanças de preço face à versão anterior**, guarda backups e regista tudo num log.

Complementa a app [`catalogo`](../catalogo/): a app não guarda histórico; esta ferramenta
guarda. Por omissão lê o **último CSV carregado na app** para o fornecedor indicado
(`catalogo/uploads/<chave>_*.csv`), por isso basta correr depois de cada upload.

## Utilização

```bash
cd price_history
pip install -r requirements.txt
python update_prices.py --supplier <chave>
```

`<chave>` é o identificador do fornecedor na app (o prefixo dos ficheiros em `catalogo/uploads/`).

Saída:

```
✓ CSV processado com 4731 produtos (fornecedor_precos.csv)
✓ Mudanças detectadas: 5 (novos 1, removidos 1, preços 3)
✓ Backup criado: backup_20260911_115250.csv
✓ Excel criado: prices_current.xlsx
✓ Log atualizado: update_log.txt
```

Outras opções:

```bash
python update_prices.py --csv "C:\Downloads\precos.csv"   # CSV em qualquer lado
python update_prices.py --out D:\precos                   # outra pasta de output
```

(ou variáveis de ambiente `CSV_SOURCE` / `SUPPLIER` / `OUTPUT_DIR`).

## Formato do CSV

| Coluna | Obrigatória | Descrição |
|---|---|---|
| `id` | sim | identificador único e estável do produto (chave de comparação) |
| `product_name` | sim | nome do produto |
| `price` | sim | preço |
| `brand`, `ean`, `sku`, `product_type`, `product_url` | não | informação extra |
| `stock_amount`, `supply_status` | não | stock / disponibilidade (numéricos) |

Os nomes das colunas não distinguem maiúsculas de minúsculas.

## Estrutura

```
price_history/
├── update_prices.py      script
├── requirements.txt      pandas, openpyxl
└── output/               gerado automaticamente
    ├── prices_current.csv    versão atual (base da próxima comparação)
    ├── prices_current.xlsx   Excel formatado
    ├── update_log.txt        histórico de execuções
    └── backups/backup_AAAAMMDD_HHMMSS.csv
```

## O Excel

| Aba | Conteúdo |
|---|---|
| **Produtos** | todos os produtos, ordenados por categoria → marca → nome; preços em `€`, cabeçalho congelado, filtros |
| **Análise** | total, em stock / indisponíveis, preço mín / médio / mediana / máx, top categorias e marcas |
| **Mudanças** | contagens de novos / removidos / preços alterados / subidas / descidas (só com versão anterior) |
| **Preços alterados** | id, produto, preço anterior, atual, variação % |
| **Novos** / **Removidos** | produtos que entraram / saíram do feed |

## Erros e códigos de saída

| Situação | Comportamento | rc |
|---|---|---|
| CSV não encontrado | mensagem com o caminho procurado | 1 |
| CSV sem as colunas obrigatórias | mensagem com colunas em falta | 2 |
| Sem versão anterior | comparação ignorada (1.ª execução) | 0 |
| Excel falha | aviso; CSV atual e log ficam na mesma | 0 |

Requisitos: Python 3.8+, `pandas`, `openpyxl`.
