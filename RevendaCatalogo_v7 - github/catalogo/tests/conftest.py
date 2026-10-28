# Configuração partilhada dos testes.
# Importar app.py arranca a thread de auto_sync() e carrega os ficheiros de uploads/ —
# WERKZEUG_RUN_MAIN=false desliga as duas coisas, para os testes não irem à rede.
import os
import sys

os.environ["WERKZEUG_RUN_MAIN"] = "false"
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import pytest

import app as _app


@pytest.fixture(autouse=True)
def _estado_limpo():
    # _AMBIG_CONN_KEYS e _MULTI_SHADE_KEYS são preenchidos por _prefilter() a partir dos
    # feeds reais e mudam o resultado de same_product(). Cada teste começa sem eles.
    def limpar():
        _app._AMBIG_CONN_KEYS.clear()
        _app._MULTI_SHADE_KEYS.clear()
        _app._colors_mergeable.cache_clear()
    limpar()
    yield
    limpar()
