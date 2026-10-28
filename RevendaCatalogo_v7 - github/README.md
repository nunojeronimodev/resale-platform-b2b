# Catálogo Comercial — agregador de preços de fornecedores

Aplicação web local (Flask + pandas) que junta as tabelas de preços de vários fornecedores
de tecnologia (smartphones, tablets e smartwatches), elimina duplicados entre fornecedores,
escolhe o preço mais baixo de cada produto, aplica margens de venda e exporta tabelas em Excel.

Projeto de exemplo — pensado para ser adaptado a qualquer negócio de revenda.

## O que inclui

| Pasta | Descrição |
|---|---|
| [`catalogo/`](catalogo/) | A aplicação principal (interface web + motor de consolidação) |
| [`price_history/`](price_history/) | Ferramenta de linha de comandos que guarda o histórico de preços de um fornecedor e deteta mudanças |

## Arranque rápido

```bash
cd catalogo
python -m venv .venv
.venv\Scripts\activate        # Windows
source .venv/bin/activate     # macOS / Linux
pip install -r requirements.txt
python app.py
```

Abre **http://localhost:5050**, vai a **Fornecedores** e adiciona as tuas fontes de dados
(URL para sincronização automática ou upload manual de CSV/Excel).

No Windows basta fazer duplo clique em `catalogo/INICIAR.bat`.

## Funcionalidades principais

- Importação de CSV / Excel com deteção automática de colunas (descrição, marca, EAN, PN, preço, stock)
- Sincronização automática por URL (a cada 24 h) ou upload manual
- Normalização de descrições (tradução ES/EN → PT de cores e termos técnicos)
- Deduplicação entre fornecedores por EAN, part number e comparação difusa com guardas
  (modelo, capacidade, cor, conectividade, tamanho)
- Auditoria automática do catálogo (acessórios, recondicionados, duplicados, preço mínimo)
- Margens configuráveis (Revenda / VIP) e exportação para Excel por categoria
- Nome da empresa e logótipo configuráveis na própria aplicação

## Licença

[MIT](LICENSE)
