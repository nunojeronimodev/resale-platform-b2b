# Testes do emparelhamento de produtos entre fornecedores (same_product e auxiliares).
# Cada caso vem de uma regra documentada em app.py — se um falhar, o catálogo passa a
# fundir produtos diferentes ou a duplicar o mesmo produto.
import pandas as pd
import pytest

import app
from app import capacity_match, category, clean_ean, safe_price, same_product


# ── produtos diferentes: nunca podem juntar ───────────────────────────────────
@pytest.mark.parametrize("a, b", [
    # cor
    ("GALAXY S25 256GB PRETO", "GALAXY S25 256GB AZUL"),
    ("IPHONE 16 128GB AZUL ESCURO", "IPHONE 16 128GB AZUL GELO"),       # duas tonalidades de azul
    # ano / geração
    ("IPAD AIR 11 2024 128GB", "IPAD AIR 11 2025 128GB"),
    # tier
    ("SAMSUNG GALAXY S25 256GB PRETO", "SAMSUNG GALAXY S25 ULTRA 256GB PRETO"),
    ("SAMSUNG GALAXY S25+ 256GB PRETO", "SAMSUNG GALAXY S25 256GB PRETO"),
    # nome comprido: uma palavra a mais ainda dá ≥ 90% no fuzzy — só a guarda de tier separa
    ("XIAOMI REDMI NOTE 14 PRO 256GB PRETO", "XIAOMI REDMI NOTE 14 PRO+ 256GB PRETO"),
    # número de modelo / versão
    ("SAMSUNG GALAXY A17 128GB PRETO", "SAMSUNG GALAXY A27 128GB PRETO"),
    ("SAMSUNG GALAXY S24 128GB PRETO", "SAMSUNG GALAXY S25 128GB PRETO"),
    ("XIAOMI REDMI NOTE 14 128GB PRETO", "XIAOMI REDMI NOTE 15 128GB PRETO"),
    # capacidade
    ("IPHONE 16 128GB PRETO", "IPHONE 16 256GB PRETO"),
    # conectividade explícita
    ("SAMSUNG GALAXY A16 4G 128GB PRETO", "SAMSUNG GALAXY A16 5G 128GB PRETO"),
    # tamanho de relógio
    ("APPLE WATCH SERIES 10 GPS 42MM PRETO", "APPLE WATCH SERIES 10 GPS 46MM PRETO"),
    # tamanho de iPad
    ("IPAD AIR 2026 11 WIFI 128GB", "IPAD AIR 2026 13 WIFI 128GB"),
])
def test_produtos_diferentes_nao_juntam(a, b):
    assert not same_product(a, b)
    assert not same_product(b, a)


# ── o mesmo produto escrito por fornecedores diferentes: tem de juntar ───────
@pytest.mark.parametrize("a, b", [
    ("IPHONE 16 128GB PRETO", "IPHONE 16 128GB PRETO"),
    # cor base ⊂ tonalidade
    ("IPHONE 16 128GB AZUL", "IPHONE 16 128GB AZUL ESCURO"),
    # alias de tonalidade: "Sky Blue" = "Light Blue"
    ("SAMSUNG GALAXY S25 128GB AZUL CÉU", "SAMSUNG GALAXY S25 128GB AZUL CLARO"),
    # RAM+armazenamento vs só armazenamento (caso XCover 7)
    ("SAMSUNG GALAXY XCOVER 7 6GB 128GB PRETO", "SAMSUNG GALAXY XCOVER 7 128GB PRETO"),
    # "5G" decorativo (há fornecedores que o escrevem sempre) num modelo sem variante 4G
    ("SAMSUNG GALAXY S25 5G 128GB PRETO", "SAMSUNG GALAXY S25 128GB PRETO"),
    # LTE ≡ 4G
    ("SAMSUNG GALAXY A16 LTE 128GB PRETO", "SAMSUNG GALAXY A16 4G 128GB PRETO"),
    # fornecedor que omite "GALAXY"
    ("SAMSUNG A36 128GB PRETO", "SAMSUNG GALAXY A36 128GB PRETO"),
    # código de modelo Samsung no nome
    ("SAMSUNG GALAXY S25 128GB S931 DOURADO", "SAMSUNG GALAXY S25 128GB DOURADO"),
    # iPad: geração/chip vs ano
    ('IPAD AIR 13" 8TH WI-FI M4 128GB', "IPAD AIR 2026 13 WIFI 128GB"),
    # descrição em espanhol
    ("MOVIL SAMSUNG GALAXY A16 128GB NEGRO", "SAMSUNG GALAXY A16 128GB PRETO"),
])
def test_mesmo_produto_junta(a, b):
    assert same_product(a, b)
    assert same_product(b, a)


# ── regras que dependem do que os feeds mostram (_prefilter) ─────────────────
def test_marcador_5g_so_distingue_quando_o_fornecedor_tem_as_duas_variantes():
    a, b = "SAMSUNG GALAXY A17 128GB PRETO", "SAMSUNG GALAXY A17 5G 128GB PRETO"
    assert same_product(a, b)                      # sem dados: "5G" tratado como decoração
    app._AMBIG_CONN_KEYS.add(app._fuzzy_key(app._prep(a)))
    assert not same_product(a, b)                  # um fornecedor lista "A17" e "A17 5G": SKUs diferentes


def test_cor_base_nao_junta_quando_o_modelo_tem_varias_tonalidades():
    a, b = "SAMSUNG GALAXY S25 FE 128GB AZUL", "SAMSUNG GALAXY S25 FE 128GB AZUL ESCURO"
    assert same_product(a, b)
    app._MULTI_SHADE_KEYS.add((app._fuzzy_key(app._prep(a)), "AZUL"))
    app._colors_mergeable.cache_clear()
    assert not same_product(a, b)                  # "Azul" genérico passa a ser ambíguo


# ── auxiliares ────────────────────────────────────────────────────────────────
def test_capacity_match():
    assert capacity_match(frozenset({"128GB"}), frozenset({"128GB"}))
    assert capacity_match(frozenset({"6GB", "128GB"}), frozenset({"128GB"}))
    assert capacity_match(frozenset({"1TB"}), frozenset({"12GB", "1TB"}))
    assert not capacity_match(frozenset({"128GB"}), frozenset({"256GB"}))
    assert not capacity_match(frozenset({"8GB", "128GB"}), frozenset({"12GB", "128GB"}))
    assert not capacity_match(frozenset(), frozenset({"128GB"}))


@pytest.mark.parametrize("bruto, esperado", [
    ("8806095420123", "8806095420123"),
    ("0190199370388", "190199370388"),     # EAN-13 com zero = UPC-12
    ("190199370388", "190199370388"),
    (8806095420123.0, "8806095420123"),    # lido do Excel como float
    ("'8806095420123", "8806095420123"),   # apóstrofo de texto do Excel
    (float("nan"), ""),
])
def test_clean_ean(bruto, esperado):
    assert clean_ean(bruto) == esperado


def test_safe_price_formatos_dos_fornecedores():
    df = pd.DataFrame({"p": ["2.421,80", "11,93", "10012.2", "44.46", "1 299,00 €", "abc"]})
    assert safe_price(df, "p").tolist() == [2421.80, 11.93, 10012.2, 44.46, 1299.00, 0.0]


@pytest.mark.parametrize("desc, tipo, esperado", [
    ("APPLE IPHONE 16 128GB PRETO", "", "SMARTPHONES"),
    ("SAMSUNG GALAXY TAB S10 FE 128GB", "", "TABLETS"),
    ("HONOR PAD 9 128GB", "", "TABLETS"),
    ("APPLE WATCH SERIES 10 42MM", "", "SMARTWATCHES"),
    ("RELÓGIO GPS COM LOCALIZADOR", "", "SMARTWATCHES"),   # relógio ganha a OUTROS
    ("LOCALIZADOR BLUETOOTH", "", None),                   # OUTROS fica fora do catálogo
    ("NOTHING PHONE 3A 128GB", "TELEFONE", "SMARTPHONES"),  # tipo declarado pelo fornecedor
    ("QUBO X-230C 4G BLACK", "", "SMARTPHONES"),           # só pela marca
    ("GARMIN VIVOACTIVE 5", "", "SMARTWATCHES"),           # "VIVO" não é a marca Vivo
])
def test_category(desc, tipo, esperado):
    assert category(desc, tipo) == esperado
