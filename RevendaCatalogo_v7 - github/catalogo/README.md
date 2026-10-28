# Catálogo Comercial

Aplicação Flask que consolida as tabelas de preços de vários fornecedores num único catálogo,
com o preço mais baixo de cada produto e margens de venda aplicadas.

## Instalação rápida (Windows)

1. Certifica-te que tens **Python 3.10+** instalado
2. Faz duplo clique em **`INICIAR.bat`**
3. O browser abre automaticamente em `http://localhost:5050`

## Instalação manual

```bash
python -m venv .venv
.venv\Scripts\activate      # Windows
source .venv/bin/activate   # macOS / Linux

pip install -r requirements.txt
python app.py
```

Abre o browser em **http://localhost:5050**.

## Primeiros passos

1. **Definições** — define o nome da empresa e carrega o teu logótipo (opcional).
2. **Fornecedores** — adiciona cada fornecedor:
   - **Auto-sync**: indica o URL do ficheiro de preços (CSV ou Excel); é descarregado ao
     arrancar e depois a cada 24 h.
   - **Manual**: carrega o ficheiro à mão sempre que tiveres uma versão nova.
3. **Margens** — define a margem de Revenda e a margem VIP.
4. **Catálogo** — consulta o resultado e exporta as tabelas em Excel.

A configuração dos fornecedores fica em `suppliers.json` (criado pela app e ignorado pelo git,
porque os URLs costumam conter chaves de acesso).

## Estrutura

```
catalogo/
├── app.py               ← backend Flask + motor de consolidação
├── requirements.txt     ← dependências
├── requirements-dev.txt ← dependências dos testes (pytest)
├── INICIAR.bat          ← lançador Windows
├── margins.json         ← margens Revenda / VIP
├── templates/index.html ← interface
├── static/              ← logótipo carregado pela app
├── tests/               ← testes do emparelhamento de produtos
├── suppliers.json       ← (gerado) fornecedores configurados
├── settings.json        ← (gerado) nome da empresa, auto-sync
└── uploads/             ← (gerado) último ficheiro de cada fornecedor
```

## Como funciona

| Passo | Função | O que faz |
|---|---|---|
| 1 | `download_provider()` / `/upload` | obtém o ficheiro de cada fornecedor |
| 2 | `map_source()` | deteta as colunas e normaliza para um esquema comum |
| 3 | `_prefilter()` | descarta acessórios, recondicionados, marcas fora da lista, sem preço ou sem stock |
| 4 | `_dedup()` | junta o mesmo produto entre fornecedores e fica com o mais barato |
| 5 | `audit_catalog()` | verifica o resultado e calcula o preço dos outros fornecedores |
| 6 | `get_final()` | aplica as margens; `/generate/<tipo>` exporta o Excel |

As marcas aceites estão em `BRANDS` e as palavras que excluem acessórios em `EXCLUDE_WORDS`,
no topo de `app.py` — ajusta-as ao teu catálogo.

## Formatos suportados

- `.csv` (qualquer separador) ou `.xlsx`
- Codificações: UTF-8, Latin-1, CP1252

### Colunas detetadas automaticamente

| Campo | Palavras-chave procuradas |
|---|---|
| Descrição | DESCRIPTION, NOME, NAME, PRODUTO, TITULO, NOMBRE, PRODUCT |
| Marca | BRAND, MARCA, FABRICANTE, MANUFACTURER |
| Part Number | PARTNUMBER, PART, REF, SKU, CODIGO |
| EAN | EAN, BARCODE, CODBAR |
| Preço | PRICE, PRECO, PRECIO, PVP, CUSTO, COST, … |
| Stock | STOCK, QUANTITY, QTY, CANTIDAD, … |

Se a deteção falhar, as colunas podem ser escolhidas à mão (rota `/set_mapping`).

## Testes

```bash
pip install -r requirements-dev.txt
pytest
```

## Segurança

O servidor escuta apenas em `127.0.0.1` e **não tem autenticação**. Não o exponhas na rede
nem na Internet sem antes acrescentar uma camada de autenticação (proxy com login + HTTPS, ou VPN).
