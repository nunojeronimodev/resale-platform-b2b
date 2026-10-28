# ─────────────────────────────────────────────────────────────────────────────
# Catálogo — agregador de preços de fornecedores (Excel / CSV)
# Fluxo principal:
#   1. download_provider()  — descarrega ficheiros dos fornecedores
#   2. map_source()         — normaliza cada ficheiro para um esquema comum
#   3. _prefilter()/_dedup() — filtros + deduplicação por EAN/PN + fuzzy, escolhe o menor preço
#   4. get_final()          — aplica margens, devolve o catálogo final
#   5. /preview, /generate  — rotas que expõem o catálogo ao frontend
# ─────────────────────────────────────────────────────────────────────────────
from flask import Flask, request, jsonify, render_template, send_file
import pandas as pd
import re, os, io, threading, traceback, json, functools, unicodedata, itertools, time as _time
import gzip as _gzip
import requests as req_lib
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 32 * 1024 * 1024
app.config['TEMPLATES_AUTO_RELOAD'] = True      # relê index.html em cada pedido (sem reiniciar o servidor)
app.config['SEND_FILE_MAX_AGE_DEFAULT'] = 0    # static/ sem cache no browser

# Compressão gzip das respostas de texto: o /preview leva o catálogo inteiro (~325 KB de
# JSON) e comprime para menos de um décimo. Não toca em ficheiros servidos em streaming
# (Excel do /generate) nem em respostas já codificadas.
_GZIP_TYPES = ("application/json", "text/html", "text/css", "application/javascript", "text/plain")
_GZIP_MIN_BYTES = 1024

@app.after_request
def _gzip_response(resp):
    if (resp.direct_passthrough or resp.status_code >= 300
            or "gzip" not in request.headers.get("Accept-Encoding", "").lower()
            or resp.headers.get("Content-Encoding")
            or not resp.content_type.split(";")[0].strip().lower().startswith(_GZIP_TYPES)):
        return resp
    data = resp.get_data()
    if len(data) < _GZIP_MIN_BYTES:
        return resp
    resp.set_data(_gzip.compress(data, 6))
    resp.headers["Content-Encoding"] = "gzip"
    resp.headers["Vary"] = "Accept-Encoding"
    resp.headers["Content-Length"] = len(resp.get_data())
    return resp

BASE_DIR      = os.path.dirname(__file__)
UPLOAD_FOLDER = os.path.join(BASE_DIR, 'uploads')
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

SUPPLIERS_JSON = os.path.join(BASE_DIR, 'suppliers.json')

# Fornecedores — adicionados pelo UI (página Fornecedores) e guardados em suppliers.json.
# Na primeira execução a lista está vazia.
def _load_suppliers():
    if os.path.exists(SUPPLIERS_JSON):
        try:
            with open(SUPPLIERS_JSON, 'r', encoding='utf-8') as fh:
                return json.load(fh)
        except Exception:
            pass
    return {}

def _save_suppliers():
    with open(SUPPLIERS_JSON, 'w', encoding='utf-8') as fh:
        json.dump(suppliers_config, fh, ensure_ascii=False, indent=2)

# Marcas reconhecidas — usadas por detect_brand() e pelos regex de categorização
BRANDS = [
    "APPLE","SAMSUNG","GOOGLE","XIAOMI","HONOR","ONEPLUS",
    "OPPO","REALME","MOTOROLA","HUAWEI","NOKIA","CUBOT","GIGASET",
    "ZTE","QUBO","SPC","DAEWOO","SUNSTECH","TCL","ULEFONE","VIVO",
    "LENOVO","AMAZFIT","GARMIN","FOREVER","LEOTEC","PANASONIC","ALCATEL",
    "BLACKVIEW","SONY","NOTHING",
]

EXCLUDE_WORDS    = ["CAPA","CASE","CABO","CABLE","CARREGADOR","CHARGER","POWERBANK",
                    "SUPORTE","PELICULA","PROTECTOR","HEADSET","AURICULAR","EARBUD",
                    # termos ingleses — sem isto uma capa "Back Cover ... Galaxy S24"
                    # era classificada como SMARTPHONE por conter "GALAXY"
                    "COVER","TEMPERED GLASS","SCREEN PROTECTOR","EARBUDS","HEADPHONE",
                    "HEADPHONES","STYLUS","DOCKING STATION","MEMORY CARD","MICROSD",
                    "POWER BANK","CAR HOLDER","WIRELESS CHARGER",
                    # Canetas: nem todos os fornecedores declaram tipo, por isso uma "Apple Pencil" ou uma
                    # "Samsung S Pen" era apanhada pela regra da marca e entrava como telemóvel.
                    "PENCIL","LAPIS","S PEN","SPEN","TECLADO","KEYBOARD"]
EXCLUDE_PREFIXES = ["SOPORTE","FUNDA","MARCO","LÁPIZ"]

# Tipos declarados pelo próprio fornecedor que nunca entram no catálogo.
# Alguns feeds trazem uma coluna product_type (phone/tablet/Smartwatches/Cases/...), o que é
# mais fiável do que adivinhar pelo nome. "Activity trackers" fica de fora desta lista
# para o category() continuar a decidir (Xiaomi Smart Band → SMARTWATCHES).
ACCESSORY_TYPES = frozenset([
    # feeds em inglês
    "CASES","SCREENPROTECTORS","CHARGERS","CABLES","EARBUDS","HEADPHONES",
    "BLUETOOTH SPEAKERS","POWERBANKS","HOLDERS","STYLUS PENS","GADGETS",
    "DOCKING STATIONS","MEMORY CARDS","BATTERIES",
    # feeds em português — comparado sem acentos
    "CAPAS","PROTETORES DE ECRA","AURICULARES","AUSCULTADORES","CARREGADORES",
    "CABOS","COLUNAS BLUETOOTH","SUPORTES","CANETAS STYLUS","ESTACOES DE ANCORAGEM",
    "CARTOES DE MEMORIA","BATERIAS",
])
TYPE_COLS = ("PRODUCT_TYPE","PRODUCTTYPE","PRODUCT TYPE","TIPO","TIPOLOGIA","FAMILIA","FAMILY")

# Tipo declarado pelo fornecedor → categoria do catálogo. Usado por category() como
# recurso quando a descrição não traz nenhuma marca/palavra reconhecida: um feed pode dizer
# "TELEFONE" num Nothing Phone ou num Sony Xperia, que de outra forma se perderiam.
# Comparado sem acentos, tal como ACCESSORY_TYPES.
TYPE_TO_CATEGORY = {
    "TELEFONE": "SMARTPHONES", "TELEFONES": "SMARTPHONES", "TELEMOVEL": "SMARTPHONES",
    "TELEMOVEIS": "SMARTPHONES", "SMARTPHONE": "SMARTPHONES", "SMARTPHONES": "SMARTPHONES",
    "PHONES": "SMARTPHONES", "MOBILE PHONES": "SMARTPHONES",
    "TABLET": "TABLETS", "TABLETS": "TABLETS",
    "SMARTWATCH": "SMARTWATCHES", "SMARTWATCHES": "SMARTWATCHES",
    "SMARTBAND": "SMARTWATCHES", "SMARTBANDS": "SMARTWATCHES",
    "WEARABLES": "SMARTWATCHES", "RELOGIOS": "SMARTWATCHES",
}

# Recondicionados / usados: o catálogo só leva produtos novos.
# Há feeds que os marcam no nome ("... 128GB Yellow Refurbished") e na marca ("Apple-Refurbished");
# as variantes PT/ES ficam aqui para o caso de outro fornecedor passar a vendê-los.
# NOTA: "GRADE" não entra de propósito — apanharia as capas "Military Grade".
_REFURB_RE = re.compile(
    r'\b(?:REFURBISHED|REFURB|RECONDICIONAD[OA]|REACONDICIONAD[OA]|RENEWED|'
    r'SEMINOVO|SEMI-NOVO|PRE-OWNED|SECOND\s+HAND)\b', re.IGNORECASE
)

# ── tradução ES → PT ──────────────────────────────────────────────────────────
# Alguns fornecedores dão as descrições em espanhol. translate_es_pt() converte-as para PT
# antes de qualquer comparação ou exibição.
# Ordenado por comprimento desc na função — frases compostas são substituídas primeiro
ES_PT = {
    # tipo de produto
    "MÓVIL":"TELEMÓVEL","MOVIL":"TELEMÓVEL",
    "TELÉFONO":"TELEMÓVEL","TELEFONO":"TELEMÓVEL",
    "RELOJ":"RELÓGIO",

    # saúde / fitness
    "SATURACIÓN DE OXÍGENO":"SATURAÇÃO DE OXIGÉNIO",
    "SATURACION DE OXIGENO":"SATURAÇÃO DE OXIGÉNIO",
    "FRECUENCIA CARDÍACA":"FREQUÊNCIA CARDÍACA",
    "FRECUENCIA CARDIACA":"FREQUÊNCIA CARDÍACA",
    "RITMO CARDIACO":"RITMO CARDÍACO",
    "HUELLA DACTILAR":"IMPRESSÃO DIGITAL",
    "RECONOCIMIENTO FACIAL":"RECONHECIMENTO FACIAL",
    "PERSONAS MAYORES":"PESSOAS IDOSAS",
    "PERSONAS MAYOR":"PESSOAS IDOSAS",
    "TERCERA EDAD":"TERCEIRA IDADE",
    "SATURACIÓN":"SATURAÇÃO","SATURACION":"SATURAÇÃO",
    "FRECUENCIA":"FREQUÊNCIA",
    "OXÍGENO":"OXIGÉNIO","OXIGENO":"OXIGÉNIO",
    "CARDIACA":"CARDÍACA",
    "PULSÓMETRO":"MONITOR CARDÍACO","PULSOMETRO":"MONITOR CARDÍACO",
    "MONITORIZACIÓN":"MONITORIZAÇÃO","MONITORIZACION":"MONITORIZAÇÃO",
    "MONITOREO":"MONITORIZAÇÃO",
    "PASOS":"PASSOS","PASO":"PASSO",
    "PODÓMETRO":"PEDÓMETRO","PODOMETRO":"PEDÓMETRO",
    "CALORÍAS":"CALORIAS",
    "DISTANCIA":"DISTÂNCIA",
    "SUEÑO":"SONO","SALUD":"SAÚDE",
    "PERSONAS":"PESSOAS","PERSONA":"PESSOA",
    "MAYORES":"IDOSOS","MAYOR":"IDOSO",
    "SENIOR":"SÉNIOR",
    "NIÑOS":"CRIANÇAS","NIÑO":"CRIANÇA",

    # câmara
    "CÁMARA PRINCIPAL":"CÂMARA PRINCIPAL","CAMARA PRINCIPAL":"CÂMARA PRINCIPAL",
    "ULTRA GRAN ANGULAR":"ULTRA GRANDE ANGULAR",
    "GRAN ANGULAR":"GRANDE ANGULAR",
    "MODO NOCHE":"MODO NOITE",
    "MEGAPÍXELES":"MEGAPÍXEIS","MEGAPIXELES":"MEGAPÍXEIS",
    "MPÍXELES":"MPÍXEIS","MPIXELES":"MPÍXEIS",
    "PÍXELES":"PÍXEIS","PIXELES":"PÍXEIS",
    "CÁMARA":"CÂMARA","CAMARA":"CÂMARA",
    "TRASERA":"TRASEIRA","DELANTERA":"FRONTAL",
    "TELEOBJETIVO":"TELEOBJETIVA",
    "NOCHE":"NOITE",

    # ecrã
    "PANTALLA":"ECRÃ","TACTIL":"TÁCTIL","TACIL":"TÁCTIL",
    "RESOLUCIÓN":"RESOLUÇÃO","RESOLUCION":"RESOLUÇÃO",
    "PULGADAS":"POLEGADAS","PULGADA":"POLEGADA",
    "BRILLO":"BRILHO",

    # hardware / memória
    "ALMACENAMIENTO":"ARMAZENAMENTO","MEMORIA":"MEMÓRIA",
    "PROCESADOR":"PROCESSADOR","NÚCLEOS":"NÚCLEOS","NUCLEOS":"NÚCLEOS",
    "VELOCIDAD":"VELOCIDADE","VERSIÓN":"VERSÃO","VERSION":"VERSÃO",
    "EXPANDIBLE":"EXPANSÍVEL","TARJETAS":"CARTÕES","TARJETA":"CARTÃO",
    "GENERACIÓN":"GERAÇÃO","GENERACION":"GERAÇÃO",

    # conectividade
    "NOTIFICACIONES":"NOTIFICAÇÕES","NOTIFICACION":"NOTIFICAÇÃO",
    "LLAMADAS":"CHAMADAS","LLAMADA":"CHAMADA",
    "VIDEOLLAMADAS":"VIDEOCHAMADAS","VIDEOLLAMADA":"VIDEOCHAMADA",
    "MENSAJES":"MENSAGENS","MENSAJE":"MENSAGEM",
    "ALTAVOCES":"ALTIFALANTES","ALTAVOZ":"ALTIFALANTE",
    "MICRÓFONO":"MICROFONE","MICROFONO":"MICROFONE",
    "LOCALIZACIÓN":"LOCALIZAÇÃO","LOCALIZACION":"LOCALIZAÇÃO",
    "SEGUIMIENTO":"RASTREAMENTO","RASTREO":"RASTREAMENTO",
    "COMPATIBLE":"COMPATÍVEL","COMPATIBLES":"COMPATÍVEIS",
    "INCLUYE":"INCLUI","INCLUIDOS":"INCLUÍDOS","INCLUIDO":"INCLUÍDO",
    "LLAVERO":"PORTA-CHAVES","PILA":"PILHA",
    "CARGADOR":"CARREGADOR","CABLE":"CABO",
    "PUERTOS":"PORTAS","PUERTO":"PORTA",
    "RECONOCIMIENTO":"RECONHECIMENTO",
    "DESBLOQUEO":"DESBLOQUEIO","HUELLA":"IMPRESSÃO DIGITAL",
    "ALARMAS":"ALARMES","ALARMA":"ALARME",
    "RECORDATORIOS":"LEMBRETES","RECORDATORIO":"LEMBRETE",
    "CALENDARIO":"CALENDÁRIO",

    # bateria / autonomia
    "BATERÍA":"BATERIA",
    "RECARGA":"CARREGAMENTO","CARGA":"CARREGAMENTO",
    "INALÁMBRICA":"SEM FIOS","INALAMBRICA":"SEM FIOS",
    "RÁPIDA":"RÁPIDO","RAPIDA":"RÁPIDO",
    "LARGA DURACIÓN":"LONGA DURAÇÃO","LARGA DURACION":"LONGA DURAÇÃO",
    "DURACIÓN":"DURAÇÃO","DURACION":"DURAÇÃO",
    "AUTONOMÍA":"AUTONOMIA",

    # materiais / pulseira
    "ACERO INOXIDABLE":"AÇO INOXIDÁVEL",
    "ACERO":"AÇO","SILICONA":"SILICONE","ALUMINIO":"ALUMÍNIO",
    "CUERO":"COURO","PIEL":"PELE","TELA":"TECIDO","MALLA":"MALHA",
    "CORREA":"BRACELETE","PULSERA":"BRACELETE",
    "ESFERA":"MOSTRADOR",
    # equivalentes em inglês — mesmos termos PT que as descrições em espanhol produzem
    "STAINLESS STEEL":"AÇO INOXIDÁVEL","STEEL":"AÇO",
    "RUBBER STRAP":"BRACELETE BORRACHA","SILICONE STRAP":"BRACELETE SILICONE",
    "LEATHER STRAP":"BRACELETE COURO","FABRIC STRAP":"BRACELETE TECIDO",
    "NYLON STRAP":"BRACELETE NYLON","MILANESE STRAP":"BRACELETE MILANESA",
    "METAL STRAP":"BRACELETE METAL","TITANIUM STRAP":"BRACELETE TITÂNIO",
    "BRACELETE STRAP":"BRACELETE",
    # "BAND" fica de fora de propósito: partiria nomes de produto (Xiaomi Smart Band 10)
    "STRAP":"BRACELETE",
    "RUBBER":"BORRACHA","LEATHER":"COURO","FABRIC":"TECIDO","WOVEN":"TECIDO",
    "MILANESE":"MILANESA",

    # cores compostas (antes das simples)
    "NEGRO OSCURO":"PRETO ESCURO","NEGRO CLARO":"PRETO CLARO",
    "NEGRO BRILLANTE":"PRETO BRILHANTE","NEGRO CARBON":"PRETO CARBONO",
    "GRIS OSCURO":"CINZENTO ESCURO","GRIS CLARO":"CINZENTO CLARO",
    "DEEP BLUE":"AZUL ESCURO","DARK BLUE":"AZUL ESCURO","DEEP AZUL":"AZUL ESCURO",
    "AZUL OSCURO":"AZUL ESCURO","VERDE OSCURO":"VERDE ESCURO",
    "GRIS ESPACIAL":"CINZENTO SIDERAL","AZUL ESPACIAL":"AZUL SIDERAL",
    "SAGE GREEN":"VERDE SÁLVIA","VERDE SALVIA":"VERDE SÁLVIA","AZUL ZAFIRO":"AZUL SAFIRA",
    "SKY BLUE":"AZUL CÉU","SILVER BLUE":"AZUL PRATEADO","SILVERBLUE":"AZUL PRATEADO",
    "SKY AZUL":"AZUL CÉU",
    "AZUL CIELO":"AZUL CÉU","AZUL MARINO":"AZUL MARINHO",
    "AZUL TORMENTA":"AZUL TEMPESTADE",   # (AZUL HIELO → AZUL CLARO, ver secção Icyblue)
    "AZUL AGUAMARINA":"AZUL ÁGUA-MARINHA","AZUL OCEANO":"AZUL OCEANO",
    "VERDE BOSQUE":"VERDE FLORESTA","VERDE LIMA":"VERDE LIMA",
    "BLANCO MARFIL":"BRANCO MARFIM","BLANCO LUNA":"BRANCO LUA",
    "BLANCO NUBE":"BRANCO NUVEM","BLANCO CHAMPAN":"BRANCO CHAMPANHE",
    "ROJO INTENSO":"VERMELHO INTENSO","VERDE ESMERALDA":"VERDE ESMERALDA",
    # cores simples
    "NEGRO":"PRETO","NEGRA":"PRETA","NEGROS":"PRETOS","NEGRAS":"PRETAS",
    "BLANCO":"BRANCO","BLANCA":"BRANCA","BLANCOS":"BRANCOS","BLANCAS":"BRANCAS",
    "PLATEADA":"PRATEADA","PLATEADOS":"PRATEADOS","PLATEADAS":"PRATEADAS",
    "ROJO":"VERMELHO","ROJA":"VERMELHA","ROJOS":"VERMELHOS","ROJAS":"VERMELHAS",
    "GRIS":"CINZENTO","DORADO":"DOURADO","PLATEADO":"PRATEADO",
    "PLATA":"PRATEADO","ORO":"DOURADO","MORADO":"ROXO",
    "AMARILLO":"AMARELO","NARANJA":"LARANJA",
    "TITANIO":"TITANIUM",
    "OSCURO":"ESCURO","OSCURA":"ESCURA",
    "AZÚL":"AZUL",
    "MARRÓN":"CASTANHO","MARRON":"CASTANHO","CAFÉ":"CASTANHO",
    "BEIGE":"BEGE","ARENA":"AREIA","CREMA":"CREME","PERLA":"PÉROLA",
    "BURDEOS":"BORDÔ","VINO":"VINHO","GRAFITO":"PRETO GRAFITE",
    "ANTRACITA":"ANTRACITE","BRONCE":"BRONZE","CIAN":"CIANO",
    "LILA":"LILÁS",
    "PIZARRA":"ARDÓSIA","ESTRELLA":"ESTELAR",
    "ESPACIAL":"SIDERAL","SAGE":"VERDE SÁLVIA","SALVIA":"VERDE SÁLVIA","ZAFIRO":"SAFIRA",
    "CIELO":"CÉU","MARINO":"MARINHO","MARFIL":"MARFIM","INDIGO":"ÍNDIGO",
    "DESIERTO":"DESERTO","HIELO":"GELO","TORMENTA":"TEMPESTADE",
    "NUBE":"NUVEM","LUNA":"LUA","BOSQUE":"FLORESTA",
    "CARBON":"CARBONO","AGUAMARINA":"ÁGUA-MARINHA",
    "CHAMPÁN":"CHAMPANHE","CHAMPA":"CHAMPANHE",
    "FRAMBUESA":"FRAMBOESA","BRILLANTE":"BRILHANTE",
    "MANDARINA":"TANGERINA","TURQUESA":"TURQUESA",

    # desporto / jogos
    "ENTRENAMIENTO":"TREINO","ENTRENAMIENTOS":"TREINOS",
    "DEPORTES":"DESPORTOS","DEPORTE":"DESPORTO",
    "DEPORTIVO":"DESPORTIVO","DEPORTIVA":"DESPORTIVA",
    "ACTIVIDAD":"ACTIVIDADE","JUEGOS":"JOGOS","JUEGO":"JOGO",

    # medidas
    "KILÓMETROS":"QUILÓMETROS","KILOMETROS":"QUILÓMETROS",
    "KILÓMETRO":"QUILÓMETRO","KILOMETRO":"QUILÓMETRO",
    "GRAMOS":"GRAMAS","GRAMO":"GRAMA","TAMAÑO":"TAMANHO",

    # geral
    "CON":"COM","SIN":"SEM","Y":"E","GRAN":"GRANDE",
    "COLOR":"COR","SERIE":"SÉRIE","EDICIÓN":"EDIÇÃO","EDICION":"EDIÇÃO",
    "PLEGABLE":"DOBRÁVEL","PLEGADO":"DOBRADO",
    "REACONDICIONADO":"RECONDICIONADO",
    "NUEVO":"NOVO","NUEVA":"NOVA","NUEVOS":"NOVOS",
    "LIBRE":"DESBLOQUEADO",
    "CALIDAD":"QUALIDADE","RENDIMIENTO":"RENDIMENTO",
    "CONTROL":"CONTROLO","AUDIO":"ÁUDIO","SONIDO":"SOM",
    "SOPORTE":"SUPORTE",
    # termos técnicos adicionais
    "FUNCIONES":"FUNÇÕES","FUNCIÓN":"FUNÇÃO","FUNCION":"FUNÇÃO",
    "RESISTENTE AL AGUA":"RESISTENTE À ÁGUA",
    "AGUA":"ÁGUA","RESISTENCIA":"RESISTÊNCIA",
    "IMPERMEABLE":"IMPERMEÁVEL","ASISTENTE":"ASSISTENTE",
    "CUERPO":"CORPO","TECNOLOGÍA":"TECNOLOGIA",
    "MODO AVION":"MODO AVIÃO","AVION":"AVIÃO","AHORRO":"POUPANÇA",

    # ── inglês → português (cores) ──
    "MIDNIGHT BLACK":"PRETO MEIA-NOITE","MATTE BLACK":"PRETO MATE",
    "PHANTOM BLACK":"PRETO FANTASMA","AWESOME BLACK":"PRETO",
    "MIDNIGHT GREEN":"VERDE MEIA-NOITE","MIDNIGHT BLUE":"AZUL MEIA-NOITE",
    "ROSE GOLD":"OURO ROSA","SPACE GRAY":"CINZENTO SIDERAL",
    "SPACE GREY":"CINZENTO SIDERAL",
    "ALPINE GREEN":"VERDE ALPINO","SIERRA BLUE":"AZUL SIERRA",
    "GRAPHITE":"PRETO GRAFITE","GRAFITE":"PRETO GRAFITE","TITANIUM":"TITANIUM","NATURAL TITANIUM":"TITANIUM NATURAL",
    "BLACK TITANIUM":"TITANIUM PRETO","WHITE TITANIUM":"TITANIUM BRANCO",
    "DESERT TITANIUM":"TITANIUM DESERTO","BLUE TITANIUM":"TITANIUM AZUL",
    # cores simples inglês
    "BLACK":"PRETO","WHITE":"BRANCO","BLUE":"AZUL","GREEN":"VERDE",
    "RED":"VERMELHO","YELLOW":"AMARELO","ORANGE":"LARANJA",
    "PURPLE":"ROXO","PINK":"ROSA","GOLD":"DOURADO","SILVER":"PRATEADO",
    "GRAY":"CINZENTO","GREY":"CINZENTO","BROWN":"CASTANHO",
    "VIOLET":"VIOLETA","LAVENDER":"ROXO LAVANDA","LAVANDA":"ROXO LAVANDA","CREAM":"BRANCO CREME","CREME":"BRANCO CREME",
    "SILVER SHADOW":"PRATEADO SHADOW","PLATA SOMBRA":"PRATEADO SHADOW",
    "CORAL":"CORAL","MINT":"VERDE MENTA","MENTA":"VERDE MENTA",
    "LILAC":"LILÁS","SAND":"AREIA","OLIVE":"VERDE AZEITONA",
    "BURGUNDY":"BORDÔ","INDIGO":"ÍNDIGO","PEACH":"PÊSSEGO",
    "COSMIC ORANGE":"LARANJA CÓSMICO","COSMIC":"CÓSMICO","COSMICO":"CÓSMICO",
    "DESERT":"DESERTO","NATURAL":"NATURAL",
    # Sinónimos entre fornecedores → UM nome canónico, senão o mesmo iPhone aparecia
    # em duplicado ("TEAL", "Verde Azulado" e "Green" em três fornecedores).
    # REGRA: o nome canónico de uma tonalidade inclui SEMPRE a cor base ("AZUL ULTRAMARINO",
    # não "ULTRAMARINO"). É isso que permite ao guard de cores em same_product() perceber
    # que "AZUL" num fornecedor e "AZUL ULTRAMARINO" noutro são o mesmo telemóvel.
    "TEAL":"VERDE AZULADO","AZUL-ESVERDEADO":"VERDE AZULADO","AZUL ESVERDEADO":"VERDE AZULADO",
    "VERDE-AZULADO":"VERDE AZULADO",
    "ULTRAMARINE":"AZUL ULTRAMARINO","ULTRAMARINO":"AZUL ULTRAMARINO",
    "AZUL ULTRAMAR":"AZUL ULTRAMARINO","ULTRAMAR":"AZUL ULTRAMARINO",
    "MIDNIGHT":"PRETO MEIA-NOITE","MEDIANOCHE":"PRETO MEIA-NOITE","MEIA-NOITE":"PRETO MEIA-NOITE",
    "STARLIGHT":"BRANCO ESTELAR","LUZ ESTELAR":"BRANCO ESTELAR","ESTELAR":"BRANCO ESTELAR",
    "LUZ DAS ESTRELAS":"BRANCO ESTELAR","BLANCO ESTRELLA":"BRANCO ESTELAR",
    # Samsung "Icyblue": "ICY BLUE", "Azul Hielo" ou "Azul Claro", consoante o fornecedor → um só nome
    "ICY BLUE":"AZUL CLARO","ICE BLUE":"AZUL CLARO","AZUL HIELO":"AZUL CLARO","AZUL GELO":"AZUL CLARO",
    "ICY":"CLARO",
    # qualificadores ingleses que ficavam por traduzir ("LIGHT DOURADO", "CLOUD BRANCO")
    "LIGHT GOLD":"DOURADO CLARO","LIGHT GRAY":"CINZENTO CLARO","LIGHT GREY":"CINZENTO CLARO",
    "LIGHT PINK":"ROSA CLARO","LIGHT GREEN":"VERDE CLARO","LIGHT BLUE":"AZUL CLARO",
    "CLOUD WHITE":"BRANCO NUVEM","CLOUD BLUSH":"ROSA NUVEM","SOFT PINK":"ROSA SUAVE",
    "FROST BLACK":"PRETO GELO","SKY TEAL":"VERDE AZULADO CÉU","SPACE BLUE":"AZUL SIDERAL",
    "SPACE BLACK":"PRETO SIDERAL","CAJA":"CAIXA","GRIS PIEDRA":"CINZENTO PEDRA","PIEDRA":"PEDRA",
    "NEGRO MEDIANOCHE":"PRETO MEIA-NOITE","FUNDA":"CAPA","CORREIA":"BRACELETE",
    "NATUREL":"NATURAL","TITANIO NATURAL":"TITANIUM NATURAL",
    # holandês que sobra em alguns feeds (braceletes dos relógios)
    "LEISTEEN":"ARDÓSIA","RUBBEREN":"DE BORRACHA","GELE":"AMARELA","PAARSE":"ROXA","WITTE":"BRANCA",
    "ZILVEREN":"PRATEADA","ZWARTE":"PRETA","ZWART":"PRETO","BLAUWE":"AZUL","GROENE":"VERDE","RODE":"VERMELHA",
    "GRIJZE":"CINZENTA","LEREN":"DE PELE","NYLON":"DE NYLON","STALEN":"DE AÇO","MILANESE":"MILANESA",
    # restos de espanhol / inglês em cores e materiais
    "VIDRIO":"VIDRO","TITÁN":"TITÂNIO","TITAN GRAY":"CINZENTO TITÂNIO","GRIS TITÁN":"CINZENTO TITÂNIO",
    "AZABACHE":"AZEVICHE","VERDE OLIVA":"VERDE AZEITONA","OLIVA":"AZEITONA","LAVANDER":"ROXO LAVANDA",
    "DESING":"DESIGN","SAPPHIRE":"SAFIRA","AZUL SAPPHIRE":"AZUL SAFIRA","PISTACHIO":"VERDE PISTÁCIO",
    "BLUEBERRY":"AZUL MIRTILO","OBSIDIAN":"PRETO OBSIDIANA","OBSIDIANA":"PRETO OBSIDIANA","HIBISCUS":"HIBISCO","FOG":"NÉVOA",
    "STELLAR BLUE":"AZUL ESTELAR","STELLAR AZUL":"AZUL ESTELAR","CELESTIAL PURPLE":"ROXO CELESTIAL",
    "CELESTIAL ROXO":"ROXO CELESTIAL","LOTUS PURPLE":"ROXO LÓTUS","LOTUS ROXO":"ROXO LÓTUS",
    "VIVID ORANGE":"LARANJA VIVO","VIVID LARANJA":"LARANJA VIVO","GLACIER BLUE":"AZUL GLACIAR",
    # "Glacier" sozinho e "Glacial" (fornecedores diferentes) são a mesma cor da Apple; sem isto a
    # linha com "Glacier" ficava sem cor nenhuma reconhecida e fundia-se com outra cor qualquer.
    "GLACIER":"AZUL GLACIAR","GLACIAL":"GLACIAR",
    "HUAWEI CORREIA":"HUAWEI BAND","SMART CORREIA":"SMART BAND","HONOR CORREIA":"HONOR BAND",
    "GALAXY FIT CORREIA":"GALAXY FIT",
    "NAVY":"AZUL MARINHO","AZUL NAVY":"AZUL MARINHO","PHANTOM NAVY":"AZUL MARINHO","AWESOME NAVY":"AZUL MARINHO",
    "DARK NIGHT GREY":"CINZENTO ESCURO","DARK NIGHT GRAY":"CINZENTO ESCURO",
    # alguns feeds em português usam termos diferentes dos restantes
    "PRATA":"PRATEADO","PÚRPURA":"ROXO","PURPURA":"ROXO","OURO":"DOURADO",
    "COR DE ROSA":"ROSA","COR-DE-ROSA":"ROSA",
    "GRIS GRAFITO":"PRETO GRAFITE","CINZENTO GRAFITE":"PRETO GRAFITE",
    # termos técnicos inglês
    "EE":"ENTERPRISE EDITION",   # abreviatura de alguns fornecedores; outros escrevem por extenso
    "DUAL SIM":"DUAL SIM","SINGLE SIM":"SINGLE SIM",
    "REFURBISHED":"RECONDICIONADO","UNLOCKED":"DESBLOQUEADO",
    "STORAGE":"ARMAZENAMENTO","MEMORY":"MEMÓRIA",
    "BATTERY":"BATERIA","SCREEN":"ECRÃ","DISPLAY":"ECRÃ",
    "CAMERA":"CÂMARA","PROCESSOR":"PROCESSADOR","NETWORK":"REDE",
    "WIRELESS":"SEM FIOS","CHARGING":"CARREGAMENTO",
    "FAST CHARGING":"CARREGAMENTO RÁPIDO",
    "WATER RESISTANT":"RESISTENTE À ÁGUA","WATERPROOF":"IMPERMEÁVEL",
    "FINGERPRINT":"IMPRESSÃO DIGITAL","FACE UNLOCK":"DESBLOQUEIO FACIAL",
}

# ── normalização e tradução ───────────────────────────────────────────────────
# norm() — converte para maiúsculas e colapsa espaços (base de todas as comparações)
def norm(txt):
    if pd.isna(txt): return ""
    return re.sub(r"\s+", " ", str(txt).upper().strip())

# translate_es_pt() — aplica ES_PT numa ÚNICA passagem (alternância longest-first).
# Cada trecho do texto original é substituído no máximo uma vez e o resultado nunca é
# re-analisado — evita cadeias tipo "STARLIGHT"→"BRANCO ESTELAR"→"BRANCO BRANCO ESTELAR"
# que aconteciam quando o output de uma entrada continha a chave de outra.
# LRU cache porque a mesma descrição aparece várias vezes. O tamanho tem de cobrir todas
# as descrições de todos os fornecedores (~6 mil e a crescer): com 4096 a cache enchia e
# a mesma descrição era retraduzida a cada passagem do pipeline.
_ES_PT_UPPER = {k.upper(): v for k, v in ES_PT.items()}
# Idempotência: traduzir texto já traduzido não pode alterá-lo. Se um valor contém outra
# chave ("PRETO MEIA-NOITE" contém "MEIA-NOITE"), regista-se o valor como chave de si
# próprio — sendo mais comprido, ganha na alternância e o texto fica intacto.
for _v in list(_ES_PT_UPPER.values()):
    _ES_PT_UPPER.setdefault(_v.upper(), _v)
_ES_PT_RE = re.compile(
    r'\b(?:' + '|'.join(re.escape(k) for k in sorted(_ES_PT_UPPER, key=len, reverse=True)) + r')\b',
    re.IGNORECASE
)

@functools.lru_cache(maxsize=65536)
def translate_es_pt(txt):
    # Espaços colapsados primeiro: há feeds com espaços duplos ("Black Silicone  Strap")
    # que impediriam o match das entradas compostas do dicionário.
    txt = re.sub(r"\s+", " ", txt)
    return _ES_PT_RE.sub(lambda m: _ES_PT_UPPER.get(m.group(0).upper(), m.group(0)), txt)

_IPHONE_PREFIX_RE    = re.compile(r'^(?:SMARTPHONE\s+)?(?:APPLE\s+)?(IPHONE\b.*)', re.IGNORECASE)
_TELEMOVEL_PREFIX_RE = re.compile(r'^(?:(?:TELEMÓVEL|SMARTPHONE)\s+)+', re.IGNORECASE)
# "SAMSUNG A36" → "SAMSUNG GALAXY A36"  (some suppliers omit GALAXY)
_SAMSUNG_GALAXY_RE   = re.compile(r'\bSAMSUNG\s+(?!GALAXY\b)(\S)', re.IGNORECASE)
_VANITY_COLORS_RE   = re.compile(
    r'\b(PLAYA|CISNE|AIRE|MOCHA|JADE|NEBULA|NEBULOSA|PHANTOM|PRISMA|PRISM|'
    r'NIEBLA|DUNA|PÁRAMO|PARAMO|SOMBRA)\b', re.IGNORECASE
)
_ORPHAN_PREP_RE     = re.compile(r'\s+\b(DEL|DE LA|DE LOS|DE LAS)\s*$', re.IGNORECASE)
_SEP_SLASH_RE       = re.compile(r'\s*/\s+')

# clean_desc() — versão final para EXIBIÇÃO: remove códigos internos de fornecedor,
# prefixos redundantes ("SMARTPHONE SAMSUNG" → "SAMSUNG"), normaliza "SAMSUNG A36"
# → "SAMSUNG GALAXY A36", apaga cores fantasia (NEBULA, PHANTOM, etc.) sem tradução PT.
# Parênteses das braceletes (alguns feeds): "(Preto Rubberen Correia)" → "(BRACELETE DE BORRACHA PRETA)",
# "(Roxa Correia de Borracha)" → "(BRACELETE DE BORRACHA ROXA)". Cor vai para o fim, no feminino.
_STRAP_FEM = {"PRETO":"PRETA","BRANCO":"BRANCA","ROXO":"ROXA","AMARELO":"AMARELA","VERMELHO":"VERMELHA",
              "PRATEADO":"PRATEADA","DOURADO":"DOURADA","CINZENTO":"CINZENTA","CASTANHO":"CASTANHA","ESCURO":"ESCURA","CLARO":"CLARA"}
def _fix_strap(m):
    before = m.group(1).split()
    after  = m.group(2).split()
    size   = [w for w in after if re.fullmatch(r'[SML](?:/[SML])?', w)]
    rest   = [w for w in after if w not in size]
    colors = [_STRAP_FEM.get(w, w) for w in before + rest if w in _COLOR_WORDS or w in _STRAP_FEM]
    others = [w for w in before + rest if not (w in _COLOR_WORDS or w in _STRAP_FEM or w == 'EM')]
    return '(' + ' '.join(w for w in ['BRACELETE'] + others + colors + size if w) + ')'

def clean_desc(txt):
    t = norm(translate_es_pt(txt))
    t = _MCODE_RE.sub('', t)
    if 'SAMSUNG' in t:
        t = _SAMSUNG_CODE_RE.sub('', t)   # "GALAXY S25 128GB S931 PRETO" → sem o código interno
    t = re.sub(r'\bOEM\b', '', t, flags=re.IGNORECASE)
    t = re.sub(r'\s+', ' ', _SEP_SLASH_RE.sub(' ', t)).strip()
    m = _IPHONE_PREFIX_RE.match(t)
    if m:
        t = m.group(1).strip()
    else:
        t = _TELEMOVEL_PREFIX_RE.sub('', t).strip()
    if "TELEMÓVEL" in t and "TELEFONE" in t:
        t = re.sub(r'\bTELEFONE\b', '', t)
    t = _SAMSUNG_GALAXY_RE.sub(r'SAMSUNG GALAXY \1', t)
    if 'IPAD' in t:   # "IPAD 2025 11TH WIFI A16 BIONIC 128GB" / "IPAD 11TH GENERATION 11"" → "IPAD 2025 11 WIFI 128GB"
        t = _ipad_norm(t)
    t = _VANITY_COLORS_RE.sub('', t)
    t = _ORPHAN_PREP_RE.sub('', t)
    # "NEGRO MEDIANOCHE" → "PRETO PRETO MEIA-NOITE"; "(Correia Correia)" em alguns feeds
    t = re.sub(r'\b([A-ZÀ-Ú0-9-]+) \1\b', r'\1', t)
    t = re.sub(r'\(\s*(?:BRACELETE|CORREIA)\s*\)', '', t)
    # alguns feeds: "(Preto Rubberen Correia)", "(Preto Milanesa Correia M)" → "(BRACELETE DE BORRACHA PRETA)"
    t = re.sub(r'\(\s*((?:[A-ZÀ-Ú-]+ )*?)BRACELETE\b([^)]*)\)', _fix_strap, t)
    t = re.sub(r'\s*[-/,]+\s*$', '', t)                      # "BRANCO ESTELAR-" → "BRANCO ESTELAR"
    t = re.sub(r'\(\s+', '(', t); t = re.sub(r'\s+\)', ')', t)
    return re.sub(r'\s+', ' ', t).strip()

_WATCH_CASE_MAT_RE  = re.compile(
    r'\b(AÇO INOXIDÁVEL|ACERO INOXIDABLE|TITANIUM|TITANIO|'
    r'ALUMÍNIO|ALUMINIUM|ALUMINIO|CERÂMICA|CERAMICA|CERAMIC|'
    r'INOXIDÁVEL|INOXIDABLE|AÇO)\b', re.IGNORECASE
)
_WATCH_BRAND_RE     = re.compile(
    r'\b(' + '|'.join(re.escape(b) for b in sorted(BRANDS, key=len, reverse=True)) + r')\b',
    re.IGNORECASE
)
_WATCH_GENERIC_STARTS = {'RELÓGIO','RELOGIO','SMARTWATCH','WATCH','RELOJ','BRACELETE','SMARTBAND'}

def clean_watch_desc(desc):
    d = _WATCH_CASE_MAT_RE.sub('', desc)
    d = re.sub(r'\bCAIXA(?: DE)?\b', '', d, flags=re.IGNORECASE)   # "CAIXA DE TITÂNIO NATURAL" → "NATURAL"
    d = re.sub(r'\s{2,}', ' ', d).strip()
    first = d.split()[0].upper() if d else ''
    if first in _WATCH_GENERIC_STARTS:
        m = _WATCH_BRAND_RE.search(d)
        d = d[m.start():] if m else re.sub(r'^\S+\s+', '', d)
    return d.strip()

# ── categorização / exclusão / marca ─────────────────────────────────────────
# Acessórios / gadgets que não são telemóveis — verificados ANTES da lista de marcas
_OUTROS_KW = frozenset([
    "LOCALIZADOR","SELFIE STICK","POWER STRIP","POWERSTRIP",
    "POWER BANK","POWERBANK",
    "REGLETA","TRIPOD","ASPIRADOR","ROBOT VACUUM","ALTIFALANTE",
    "ALTAVOZ","SMART SPEAKER","WEBCAM",
])

# category() — classifica cada produto; devolve None para descartar.
# Ordem importa: TABLETS > SMARTWATCHES > OUTROS > SMARTPHONES.
# "Relógio GPS" tem keyword WATCH → SMARTWATCHES antes de OUTROS ver LOCALIZADOR.
# Produtos sem categoria reconhecida são filtrados em _prefilter().
def category(txt, tipo=""):
    # traduzido primeiro: há feeds que escrevem "Huawei Correia 11 Pro" (= Huawei Band 11 Pro)
    t = norm(translate_es_pt(str(txt)))
    if any(x in t for x in ["IPAD","TABLET","GALAXY TAB","TAB S","TAB A","TAB E",
                              "LENOVO TAB","IDEATAB","KINGKONG TAB","TCL TAB","IDEA TAB",
                              "LEGION TAB","NXTPAPER 1","MAGICPAD","MATEPAD","REDMI PAD",
                              "XIAOMI PAD","POCO PAD","ONEPLUS PAD","OPPO PAD"])             or _PAD_RE.search(t):
        return "TABLETS"
    if any(x in t for x in ["WATCH","RELOJ","RELÓGIO","SMARTBAND","GALAXY FIT",
                              "GALAXY RING","MI BAND","HONOR BAND","HUAWEI BAND",
                              "SMART BAND","AMAZFIT","GARMIN","VIVOSMART","VIVOACTIVE",
                              "FOREVER VELORA","FOREVER LUCCA","PIXEL BAND"]):
        return "SMARTWATCHES"
    # OUTROS (localizadores, colunas, aspiradores...) fica FORA do catálogo: só entram
    # Smartphones, Tablets e Smartwatches. Verificado depois de SMARTWATCHES para
    # "Relógio com Localizador" continuar a ser um relógio. Tem prioridade sobre o tipo
    # declarado pelo fornecedor — um localizador nunca entra como telemóvel.
    if any(x in t for x in _OUTROS_KW):
        return None
    if any(x in t for x in ["IPHONE","GALAXY","PIXEL","SMARTPHONE","KINGKONG",
                              "REDMI","POCO","TELEMÓVEL","TELÉFONO","TELEFONO","MOVIL","MÓVIL"]):
        return "SMARTPHONES"
    # O tipo declarado pelo fornecedor vale mais do que adivinhar pela marca: um feed pode dizer
    # "TABLET" num Honor Pad, que a regra da marca classificaria como telemóvel por ser HONOR.
    declarado = TYPE_TO_CATEGORY.get(_strip_accents(str(tipo).strip().upper()))
    if declarado:
        return declarado
    # Sem tipo declarado (nem todos os fornecedores têm essa coluna): marca reconhecida como
    # fabricante de telemóveis, por palavra inteira para "VIVO" não apanhar "VIVOACTIVE".
    # Sem isto um "QUBO X-230C_4G 48MB 4G BLACK" caía fora do catálogo.
    return "SMARTPHONES" if _PHONE_BRAND_RE.search(t) else None

# Marcas cujos produtos, na ausência de outra pista, são telemóveis. Exclui as marcas
# só de wearables/acessórios (Amazfit, Garmin, Forever, Leotec) — essas são apanhadas
# pelas palavras-chave de SMARTWATCHES ou ficam de fora de propósito.
_PHONE_ONLY_BRANDS = [b for b in BRANDS if b not in
                      ("AMAZFIT", "GARMIN", "FOREVER", "LEOTEC")]
# "Honor Pad", "Ulefone Armor Pad" e afins: qualquer PAD isolado é tablet. Os acessórios
# com "pad" no nome (gamepad, mousepad, charging pad) são palavras únicas e não batem.
_PAD_RE = re.compile(r'\bPAD\b')
_PHONE_BRAND_RE = re.compile(
    r'\b(' + '|'.join(re.escape(b) for b in sorted(_PHONE_ONLY_BRANDS, key=len, reverse=True)) + r')\b')

_EXCLUDE_RE  = re.compile(
    r'\b(' + '|'.join(re.escape(w) for w in EXCLUDE_WORDS) + r')\b', re.IGNORECASE
)
_EXCLUDE_PFX = tuple(p.upper() for p in EXCLUDE_PREFIXES)

def ignored(txt):
    t = norm(txt)
    return bool(_EXCLUDE_RE.search(t)) or t.startswith(_EXCLUDE_PFX)

_BRAND_BY_MODEL = {
    "IPHONE":"APPLE","IPAD":"APPLE","MACBOOK":"APPLE","AIRPODS":"APPLE",
    "GALAXY":"SAMSUNG","REDMI":"XIAOMI","POCO":"XIAOMI",
    "MOTO ":"MOTOROLA","RAZR":"MOTOROLA",
}

def detect_brand(text):
    t = norm(text)
    for model, brand in _BRAND_BY_MODEL.items():
        if model in t: return brand
    m = _BRANDS_WORD_RE.search(t)      # por palavra: "VIVOSMART" (Garmin) não é "VIVO"
    return m.group(0) if m else "OUTRAS"

_BRANDS_WORD_RE = re.compile(r'\b(' + '|'.join(re.escape(b) for b in sorted(BRANDS, key=len, reverse=True)) + r')\b')

# ── leitura de ficheiros ──────────────────────────────────────────────────────
HEADER_KEYWORDS = {
    "DESCRIPTION","DESCRIP","NOME","NAME","PRODUTO","TITULO","TITLE","NOMBRE","PRODUCT",
    "PRICE","PRECO","PRECIO","PVP","CUSTO","COST","VALOR",
    "BRAND","MARCA","EAN","BARCODE","SKU","REF","PARTNUMBER","PART",
    "QUANTITY","QTY","STOCK","CATEGORY","CATEGORIA",
}

def find_header_row(path):
    try:
        raw = pd.read_excel(path, header=None, nrows=15)
        for skip in range(min(10, len(raw))):
            row_vals = raw.iloc[skip].dropna().astype(str).str.strip().tolist()
            text_vals = [v for v in row_vals if not re.match(r'^[\d€$.,\s]+$', v)]
            if len(text_vals) < 2: continue
            cols_upper = {re.sub(r'[^A-Z0-9]', '', v.upper()) for v in text_vals}
            hits = sum(1 for k in HEADER_KEYWORDS
                       if any(k == cu or (len(k) > 4 and k in cu) for cu in cols_upper))
            if hits >= 2:
                print(f"[HEADER] row {skip}, hits={hits}")
                return skip
    except Exception as e:
        print(f"[HEADER] scan error: {e}")
    return 0

def read_file(path):
    if path.lower().endswith(".xlsx"):
        df = pd.read_excel(path, header=find_header_row(path))
        df.columns = [str(c).strip() for c in df.columns]
        return df
    for enc in ["utf-8", "latin1", "cp1252"]:
        try:
            df = pd.read_csv(path, sep=None, engine="python", encoding=enc)
            df.columns = [str(c).strip() for c in df.columns]
            return df
        except Exception:
            pass
    return pd.read_csv(path)

def find_col(df, *keywords):
    def cl(s): return re.sub(r'[^A-Z0-9]', '', str(s).upper())
    for kw in keywords:
        kw_c = cl(kw)
        for col in df.columns:
            if kw.upper() in str(col).upper(): return col
            if kw_c and kw_c in cl(col):       return col
    return None

EAN_RE = re.compile(r"^'?\d{8,14}(\.0)?$")
PN_RE  = re.compile(r"^[A-Z0-9]{2,}[/\-][A-Z0-9]+$", re.IGNORECASE)

def guess_col(df, role, skip=None):
    skip = set(skip or [])
    candidates = []
    for col in df.columns:
        if col in skip: continue
        vals = df[col].dropna().astype(str).str.strip().str.replace(r"\.0$","",regex=True)
        if vals.empty: continue
        n = len(vals)
        if role == "ean":
            score = vals.apply(lambda v: bool(EAN_RE.match(v.replace("'","")))).sum() / n
            if score > 0.5: candidates.append((score, col))
        elif role == "pn":
            score = vals.apply(lambda v: bool(PN_RE.match(v))).sum() / n
            if score > 0.3: candidates.append((score, col))
        elif role == "desc":
            avg = vals.str.len().mean()
            hits = vals.str.upper().str.contains(
                r"IPHONE|GALAXY|SMARTPHONE|TABLET|WATCH|HUAWEI|XIAOMI|SAMSUNG|APPLE|MOTOROLA",
                regex=True, na=False).mean()
            if avg > 15 or hits > 0.3: candidates.append((avg + hits * 20, col))
        elif role == "brand":
            uniq = vals.str.upper().nunique()
            avg  = vals.str.len().mean()
            if uniq < 40 and avg < 20: candidates.append((-uniq, col))
    return candidates and max(candidates)[1] or None

PRICE_KWS = [
    "PRICE","PRECO","PRECIO","PVP","CUSTO","COST","VALOR","NET",
    "TARIF","TARIFA","IMPORTE","AMOUNT","UNITPRICE","UNIT_PRICE",
    "NETTO","NETO","LISTE","LIST","RETAIL","VENDA","SALE",
    "PRICEINCL","PRICEWITH","PRICEEXCL","PVPINCL","PVPIVA",
    "NETPRICE","NETOPRICE","SALESPRICE",
]

def safe_price(df, col):
    if col is None: return pd.Series([0.0] * len(df))
    s = df[col].astype(str).str.strip()
    s = s.str.replace(r"[€$£\s]", "", regex=True)
    s = s.str.replace(r"\.(?=\d{3}(?:[,\s]|$))", "", regex=True)
    s = s.str.replace(",", ".", regex=False)
    return pd.to_numeric(s, errors="coerce").fillna(0.0)

# ── mapeamento de colunas ─────────────────────────────────────────────────────
DESC_COLS  = ("DESCRIPTION","DESCRIP","NOME","NAME","PRODUTO","TITULO","TITLE","NOMBRE","PRODUCT")
BRAND_COLS = ("BRAND","MARCA","FABRICANTE","MANUFACTURER")
PN_COLS    = ("PARTNUMBER","PART","REF","SKU","CODIGO","CODE")
EAN_COLS   = ("EAN","BARCODE","CODBAR")
STOCK_COLS = ("STOCK","QUANTITY","QTY","CANTIDAD","DISPONIVEL","DISPONIBLE","EXISTENCIAS")

def _resolve_cols(df, ov, src):
    desc_col  = ov.get("desc")  or find_col(df, *DESC_COLS)
    brand_col = ov.get("brand") or find_col(df, *BRAND_COLS)
    pn_col    = ov.get("pn")    or find_col(df, *PN_COLS)
    ean_col   = ov.get("ean")   or find_col(df, *EAN_COLS)
    price_col = ov.get("price") or find_col(df, *PRICE_KWS)
    stock_col = ov.get("stock") or find_col(df, *STOCK_COLS)

    known = {c for c in [desc_col, brand_col, pn_col, ean_col, price_col, stock_col] if c}
    if not ean_col:   ean_col   = guess_col(df, "ean",   skip=known); known.add(ean_col)
    if not pn_col:    pn_col    = guess_col(df, "pn",    skip=known); known.add(pn_col)
    if not desc_col:  desc_col  = guess_col(df, "desc",  skip=known); known.add(desc_col)
    if not brand_col: brand_col = guess_col(df, "brand", skip=known); known.add(brand_col)

    if not price_col:
        skip_kws = {"QUANT","STOCK","QTY","ARRIV","DATE","AVAIL","END","PROMO"}
        for col in df.columns:
            if col in known or any(x in str(col).upper() for x in skip_kws): continue
            try:
                samp = pd.to_numeric(
                    df[col].astype(str).str.replace(r"[€$£,\s]","",regex=True)
                                      .str.replace(",",".",regex=False),
                    errors="coerce").dropna()
                if len(samp) > 5 and samp.mean() > 0.5:
                    price_col = col; break
            except Exception:
                pass

    print(f"[{src}] desc={desc_col} brand={brand_col} pn={pn_col} ean={ean_col} price={price_col} stock={stock_col}")
    return desc_col, brand_col, pn_col, ean_col, price_col, stock_col

# map_source() — normaliza qualquer ficheiro de fornecedor para o esquema comum:
# EAN, DESCRICAO, MARCA, PN, PRECO, FONTE, STOCK, STOCK_CHEGADA, DATA_CHEGADA.
# _resolve_cols() detecta as colunas automaticamente (por nome + heurística),
# com possibilidade de override manual pelo utilizador via /set_mapping.
def map_source(df, src, overrides=None):
    ov = overrides or {}
    desc_col, brand_col, pn_col, ean_col, price_col, stock_col = _resolve_cols(df, ov, src.upper())
    out = pd.DataFrame()
    out["EAN"]       = df[ean_col].values   if ean_col   else ""
    out["DESCRICAO"] = df[desc_col].values  if desc_col  else ""
    out["MARCA"]     = df[brand_col].values if brand_col else ""
    out["PN"]        = df[pn_col].values    if pn_col    else ""
    out["PRECO"]     = safe_price(df, price_col)
    out["FONTE"]     = src.upper()
    out["STOCK"]     = pd.to_numeric(df[stock_col], errors="coerce") if stock_col else float("nan")

    # Tipo declarado pelo fornecedor (ex.: coluna product_type). _prefilter() usa-o
    # para descartar acessórios antes de os classificar pelo nome.
    type_col = ov.get("type") or find_col(df, *TYPE_COLS)
    out["TIPO_FONTE"] = (df[type_col].astype(str).str.strip().str.upper().values
                         if type_col else "")

    # Campos de chegada (stock a chegar / data prevista) — opcionais, nem todos os fornecedores os têm
    arriving_col = next((c for c in df.columns if "arriving" in str(c).lower()), None)
    date_col     = next((c for c in df.columns if "expected" in str(c).lower()), None)
    out["STOCK_CHEGADA"] = pd.to_numeric(df[arriving_col], errors="coerce").fillna(0).astype(int) if arriving_col else 0

    def _fmt_date(v):
        if pd.isna(v) or str(v).strip() in ("", "nan", "NaT", "None", "0"): return ""
        try:
            s = str(int(float(str(v).strip())))
            if len(s) == 8 and s.isdigit():  # formato YYYYMMDD
                return f"{s[6:8]}/{s[4:6]}/{s[0:4]}"
            ts = pd.Timestamp(v)
            if ts.year < 2000: return ""
            return ts.strftime("%d/%m/%Y")
        except: return str(v).strip()
    out["DATA_CHEGADA"] = df[date_col].apply(_fmt_date) if date_col else ""

    return out

# ── dedup e consolidação ──────────────────────────────────────────────────────
CAPACITY_RE = re.compile(r'\b\d+\s*(?:GB|TB)\b', re.IGNORECASE)

# capacity_tokens() — extrai tokens de capacidade (ex: {"128GB","8GB"}) para
# comparação exata antes do fuzzy: 128GB ≠ 256GB mesmo que tudo o resto coincida.
def capacity_tokens(txt):
    return frozenset(t.upper().replace(" ", "") for t in CAPACITY_RE.findall(txt))

def _cap_gb(tok):
    n = int(re.sub(r'\D', '', tok))
    return n * 1024 if tok.endswith("TB") else n

# capacity_match() — mesma capacidade? Igualdade exata, com uma tolerância: se um lado
# só indica o armazenamento ("128GB") e o outro indica RAM+armazenamento
# ("6GB 128GB"), compara-se só o maior valor (o armazenamento é
# sempre ≥ RAM). Sem isto o mesmo XCover 7 aparecia duas vezes.
def capacity_match(ca, cb):
    if ca == cb:
        return True
    if not ca or not cb:
        return False
    if len(ca) == 1 or len(cb) == 1:
        return max(map(_cap_gb, ca)) == max(map(_cap_gb, cb))
    return False

# fuzzy_ratio() — coeficiente de Dice ao nível de palavras (0-100).
# Limiar de 90% em same_product() para absorver diferenças menores de formatação.
def fuzzy_ratio(a, b):
    wa, wb = set(a.lower().split()), set(b.lower().split())
    if not wa or not wb: return 0
    return int(100 * 2 * len(wa & wb) / (len(wa) + len(wb)))

# Tiers that must match exactly — ULTRA ≠ PLUS ≠ base model
_TIERS = frozenset(["ULTRA","PLUS","PRO","MAX","FE","EDGE","LITE","FOLD","FLIP","NOTE"])
_TIER_RES = {x: re.compile(r'\b' + x + r'\b') for x in _TIERS}

# Model number guard: A17 ≠ A27, S24 ≠ S25, P30 ≠ P40, etc.
_MODEL_NUM_RE = re.compile(r'\b([A-Z])(\d{2,3})\b')
# Version number guard: NOTE 14 ≠ NOTE 15, RUGKING 3 ≠ RUGKING 4, etc.
# Matches standalone 1-3 digit numbers — not inside GB/TB/G/MHz (no word boundary follows)
_VERSION_NUM_RE = re.compile(r'\b\d{1,3}\b')
# Connectivity guard: 5G ≠ 4G, but "5G" vs silent = same product (suppliers are inconsistent)
_CONN_RE = re.compile(r'\b(5G|4G|LTE|CELL|CELLULAR)\b')

# Colors: different color word sets = different products (PRETO ≠ BRANCO, etc.)
_COLOR_WORDS = frozenset([
    "PRETO","PRETA","BRANCO","BRANCA","AZUL","VERDE","VERMELHO","VERMELHA",
    "AMARELO","AMARELA","LARANJA","ROXO","ROXA","ROSA","CINZENTO","CINZENTA",
    "DOURADO","DOURADA","PRATEADO","PRATEADA","GRAFITE","VIOLETA","LAVANDA",
    "CORAL","MENTA","TURQUESA","LILAS","AREIA","BEGE","PEROLA","CASTANHO",
    "CASTANHA","BORDO","VINHO","ANTRACITE","BRONZE","CIANO","ARDOSIA",
    "CHAMPANHE","FRAMBOESA","TANGERINA","BRILHANTE","COSMICO","ULTRAMARINO",
    "DESERTO","NATURAL","TITANIUM","ESTELAR","SIDERAL","SALVIA","SAFIRA",
    "INDIGO","PESSEGO","MEIA-NOITE","MARINHO","GELO","FLORESTA","LIMA",
    "ESMERALDA","AZUL-ESVERDEADO","AZULADO","ESCURO","ESCURA","CLARO","CLARA",
    # tonalidades que o dicionário produz — TODAS têm de estar aqui, senão o guard
    # de cores fica cego (conjunto vazio) e funde cores diferentes (CREAM ⊂ GRAPHITE)
    "CREME","SHADOW","CEU","LUA","MARFIM","NUVEM","CARBONO","MATE","FANTASMA",
    "AZEITONA","TEMPESTADE","OCEANO","ALPINO","SIERRA","AGUA-MARINHA","SOMBRA",
    "NEVOA","PETROLEO","GRAFENO","CARVAO","CHUMBO","SUAVE","PEDRA","FROST",
    "AZEVICHE","MIRTILO","OBSIDIANA","HIBISCO","PISTACIO","CELESTIAL","LOTUS","GLACIAR","VIVO",
])

# Cores base — há fornecedores que descrevem quase sempre só com estas ("Blue", "Green", "Orange"),
# enquanto outros usam o nome comercial ("Deep Blue"→AZUL ESCURO, "Sage"→VERDE SÁLVIA).
_BASE_COLORS = frozenset([
    "PRETO","BRANCO","AZUL","VERDE","VERMELHO","AMARELO","LARANJA","ROXO","ROSA",
    "CINZENTO","DOURADO","PRATEADO","CASTANHO","BEGE","VIOLETA",
])

# _colors_compatible() — decide se dois conjuntos de cores DIFERENTES ainda descrevem o
# mesmo produto: um lado tem só a cor base e o outro tem uma tonalidade dessa cor.
#   {AZUL} ⊂ {AZUL, ESCURO}        → mesmo ("Blue" num fornecedor = "Azul Escuro" noutro)
#   {LARANJA} ⊂ {LARANJA, COSMICO} → mesmo
#   {AZUL} vs {PRETO}              → diferente
#   {AZUL, ESCURO} vs {AZUL, GELO} → diferente (duas tonalidades distintas de azul)
def _colors_compatible(ca, cb):
    small, big = (ca, cb) if len(ca) < len(cb) else (cb, ca)
    return len(small) == 1 and next(iter(small)) in _BASE_COLORS and small < big

# _colors_mergeable() — regra completa, usada pelo dedup E pela auditoria (têm de concordar):
# cores iguais → sim; cor base ⊂ tonalidade → sim, EXCETO se para este modelo algum
# fornecedor lista várias tonalidades dessa cor base (ver _MULTI_SHADE_KEYS) — aí só o EAN
# ou a cor exata decidem. Ambos os conjuntos devem vir não-vazios.
# Tonalidades que os fornecedores nomeiam de forma diferente para o mesmo produto:
# "Sky Blue" (AZUL CÉU) = "Light Blue"/"Azul Claro" (mesmo código de cor Samsung);
# "Violeta" = "Púrpura" (ROXO) — confirmado por EAN comum entre fornecedores.
# Aplicado só na comparação — a descrição mantém o nome do fornecedor vencedor.
# BORDÔ ≡ VERMELHO e GLACIAR ≡ AZUL: no iPhone 18 Pro há fornecedores que chamam "Vermelho"
# ao Burgundy (outros: Bordô) e "Azul" ao Glacier (outros: "Glacier", "Azul Glacial").
# Só afeta a COMPARAÇÃO — a tabela mantém o nome do fornecedor que ganhar o preço.
_SHADE_ALIAS = {"CEU": "CLARO", "VIOLETA": "ROXO", "BORDO": "VERMELHO", "GLACIAR": "AZUL"}
@functools.lru_cache(maxsize=16384)
def _canon_colors(cs):
    return frozenset(_SHADE_ALIAS.get(w, w) for w in cs)

# Cache: os mesmos pares de cores repetem-se milhares de vezes numa consolidação.
# Depende de _MULTI_SHADE_KEYS, que _prefilter() recalcula a cada importação —
# por isso é lá limpa (_clear_match_caches).
@functools.lru_cache(maxsize=65536)
def _colors_mergeable(ca, cb, model_key):
    ca, cb = _canon_colors(ca), _canon_colors(cb)
    if ca == cb:
        return True
    if not _colors_compatible(ca, cb):
        return False
    base = next(iter(ca if len(ca) < len(cb) else cb))
    return (model_key, base) not in _MULTI_SHADE_KEYS

# _enrich_color() — se 'stored' só tem a cor base e 'candidate' tem a tonalidade dessa
# cor, devolve 'stored' (já traduzido para PT) com a tonalidade no lugar da cor base:
#   stored="APPLE IPHONE 16 128GB BLUE", candidate="IPHONE 16 128GB ULTRAMARINE"
#   → "APPLE IPHONE 16 128GB AZUL ULTRAMARINO"
# Caso contrário devolve 'stored' traduzido, sem outras alterações.
def _strip_accents(s):
    return unicodedata.normalize('NFD', s).encode('ascii', 'ignore').decode()

def _enrich_color(stored, candidate):
    st, ct = norm(translate_es_pt(str(stored))), norm(translate_es_pt(str(candidate)))
    sp, cp = _match_prep(st), _match_prep(ct)
    cs = frozenset(w for w in sp.split() if w in _COLOR_WORDS)
    cc = frozenset(w for w in cp.split() if w in _COLOR_WORDS)
    if not (len(cs) == 1 and cs < cc):
        return st
    base  = next(iter(cs))
    words = cp.split()
    idx   = [i for i, w in enumerate(words) if w in cc]
    if idx[-1] - idx[0] != len(idx) - 1:       # tonalidade não é contígua → não arriscar
        return st
    phrase = " ".join(words[idx[0]:idx[-1] + 1])   # sem acentos (vem de _match_prep)
    # Recupera a frase COM acentos a partir do candidato traduzido, e localiza a cor base
    # no vencedor — ambos por comparação sem acentos, só se o comprimento se mantiver.
    st_a, ct_a = _strip_accents(st), _strip_accents(ct)
    if len(st_a) != len(st) or len(ct_a) != len(ct):
        return st
    mc = re.search(r'\b' + re.escape(phrase) + r'\b', ct_a)
    ms = re.search(r'\b' + base + r'\b', st_a)
    if not (mc and ms):
        return st
    return st[:ms.start()] + ct[mc.start():mc.end()] + st[ms.end():]

# _match_prep() — pré-processamento APENAS para comparação (não para exibição).
# Remove ruído por-fornecedor que baixaria o fuzzy_ratio abaixo de 90%:
#   _MCODE_RE  — códigos internos tipo "SM-S948BLBDEUE" ou "MG8Q4ZD/A"
#   _SCREEN_RE — tamanhos de ecrã "6.1"" que diferem entre fornecedores
#   _MNOISE_RE — prefixos como "SMARTPHONE" ou "TELEMÓVEL" que um
#                fornecedor inclui e o outro não; sem isto o ratio caía para ~75%
_MCODE_RE   = re.compile(r'\b[A-Z]{2,3}-[A-Z0-9]{3,}\b'            # SM-S948BLBDEUE
                         r'|\b[A-Z0-9]{4,}/[A-Z0-9]\b'            # MG8Q4ZD/A, MD274SX/A
                         r'|\b[A-Z0-9]{4,}-\d[A-Z]{3,}[A-Z0-9]*(?:-\d+)?\b'  # TCL: T314D1-3ALCA112, 9566X-2BLCA111-1
                         r'|\(?\bS\d{5}-[A-Z0-9]+-[A-Z0-9]+\)?'     # Gigaset: (S30853-H1230-R102)
                         r'|\bZA[A-Z0-9]{2}\d{4}[A-Z]{0,2}\b')       # Lenovo: ZAFR0389ES, ZAE40074ES
_SCREEN_RE  = re.compile(r'\b\d+[.,]\d+"(?=\s|$|/)|\b\d+[.,]\d+(?=\s)|\b\d+"')  # 6.1" 5G / 6.9/ / 11"
_MNOISE_RE  = re.compile(r'\b(?:OEM|DS|DL|DUAL|SMARTPHONE|TELEMÓVEL|TELEMOVEL|MÓVIL|MOVIL|'
                         r'RUGERIZADO|RUGGED|ROBUSTO|TABLET|WIFI|WI-FI|OCTACORE|OCTA-CORE|GPS)\b', re.IGNORECASE)
# Códigos de modelo Samsung que alguns fornecedores metem no nome ("GALAXY S25 128GB S931 GOLD",
# "TAB S10 FE X520"): letra + 3 dígitos. Sem os remover, _MODEL_NUM_RE via {S25, S931}
# vs {S25} e bloqueava TODOS os Samsung desse fornecedor. Só em descrições Samsung — noutras
# marcas letra+3 dígitos é o próprio modelo (Vivo X300, Realme C100, Moto G100).
_SAMSUNG_CODE_RE = re.compile(r'\b[A-Z]\d{3}\b')

# _AMBIG_CONN_KEYS — chaves de modelo (ver _fuzzy_key) que, DENTRO DE UM MESMO FORNECEDOR,
# aparecem com E sem marcador 4G/5G. Nesses modelos o silêncio distingue variantes
# (um fornecedor lista "Galaxy A17" e "Galaxy A17 5G"; "iPad Air WiFi" e "WiFi + 5G").
# Nos restantes (Galaxy S25, XCover 7…) o "5G" é decoração de quem o escreve sempre.
# Preenchido por _prefilter() a partir dos feeds reais; vazio ⇒ tratamento leniente.
_AMBIG_CONN_KEYS: set = set()

# _MULTI_SHADE_KEYS — pares (modelo, cor base) em que algum fornecedor lista DUAS tonalidades
# dessa cor base (ex.: S25 FE "Azul Claro" e "Azul Escuro"). Aí, um "Azul" genérico de
# outro fornecedor é ambíguo (outro pode chamar "Azul" ao Icyblue) e a regra cor-base ⊂ tonalidade
# NÃO se aplica — só fusão por EAN ou cor exata. Preenchido por _prefilter().
_MULTI_SHADE_KEYS: set = set()

# _fuzzy_key() — identidade do modelo para o fuzzy: sem conectividade, capacidades nem cores
# (cada uma tem o seu guard próprio em same_product).
@functools.lru_cache(maxsize=65536)
def _fuzzy_key(prepped):
    t = _CONN_RE.sub('', prepped)
    t = CAPACITY_RE.sub('', t)
    return " ".join(w for w in t.split() if w not in _COLOR_WORDS)

# iPad: cada fornecedor identifica a geração à sua maneira — "iPad Air 13\" 8th Wi-Fi M4",
# "iPad Air 2026 13 WiFi", "iPad 11th Generation 11\"". Normaliza para
# "IPAD AIR 2026 13": geração/chip → ano (quando não há ano), polegadas ficam como número.
_IPAD_YEAR = [  # (família, token, ano)
    (r'IPAD AIR',  r'\bM4\b',       '2026'), (r'IPAD AIR',  r'\bM3\b',        '2025'),
    (r'IPAD AIR',  r'\bM2\b',       '2024'), (r'IPAD PRO',  r'\bM5\b',        '2025'),
    (r'IPAD PRO',  r'\bM4\b',       '2024'), (r'IPAD MINI', r'\bA17( PRO)?\b', '2024'),
    # iPad "normal" 11.ª geração (A16) tem 11": há quem escreva "iPad 2025 11th" sem tamanho
    (r'IPAD(?! AIR| PRO| MINI)', r'\bA16\b', '2025 11'), (r'IPAD(?! AIR| PRO| MINI)', r'\b11TH\b', '2025 11'),
]
def _ipad_norm(t):
    t = _SEP_SLASH_RE.sub(' ', t)                          # "M5/ 1TB/" → "M5 1TB"
    t = re.sub(r'(\d+)"', r'\1', t)                       # 13" → 13 (o tamanho distingue 11 de 13)
    for fam, tok, repl in _IPAD_YEAR:
        if re.search(fam, t) and re.search(tok, t):
            t = re.sub(tok, repl, t, count=1); break
    t = re.sub(r'\b(\d+TH|GENERATION|BIONIC|M[2-5]|A1[5-9]( PRO)?)\b', ' ', t)
    t = ' '.join(dict.fromkeys(t.split()))                  # "2025 11 2025" → "2025 11"
    # exibição: "IPAD 2025 WIFI 11 128GB" → "IPAD 2025 11 WIFI 128GB"
    return re.sub(r'\b(20\d\d) (WI-?FI(?: CELL)?(?: 5G)?) (11|13)\b', r'\1 \3 \2', t)

def _match_prep(txt):
    t = _MCODE_RE.sub(' ', txt)
    if 'IPAD' in t.upper():
        t = _ipad_norm(t.upper())
    t = _SCREEN_RE.sub(' ', t)
    t = _MNOISE_RE.sub(' ', t)
    if 'SAMSUNG' in t.upper():
        t = _SAMSUNG_CODE_RE.sub(' ', t)
    t = t.replace('/', ' ')   # some feeds use "1TB/ 6.9"/ 5G/" delimiters
    t = re.sub(r'(?<!\S)\+(?!\S)', ' ', t)   # "WIFI + 5G" → "WIFI 5G" (o "+" solto escondia a diferença)
    # "12+256GB" → "12GB 256GB"
    t = re.sub(r'(\d+)\+(\d+\s*(?:GB|TB))', lambda m: m.group(1)+'GB '+m.group(2), t, flags=re.IGNORECASE)
    t = _SAMSUNG_GALAXY_RE.sub(r'SAMSUNG GALAXY \1', t)
    t = re.sub(r'\s+', ' ', t).strip()
    # Accent-insensitive comparison: AZÚL→AZUL, SÁLVIA→SALVIA, etc.
    return unicodedata.normalize('NFD', t).encode('ascii', 'ignore').decode()

# _prep(): a MESMA preparação em todo o lado (same_product, chaves de ambiguidade,
# buckets do dedup, auditoria). Antes, same_product trocava "A11+" por "A11 PLUS" mas as
# chaves _AMBIG_CONN_KEYS/_MULTI_SHADE_KEYS eram calculadas com "A11+" → nunca batiam e
# as variantes WiFi/5G de qualquer modelo com "+" (Tab A11+, S25+) fundiam-se.
@functools.lru_cache(maxsize=65536)
def _prep(txt):
    t = norm(translate_es_pt(str(txt)))
    # "S25+" / "PRO+" → "S25 PLUS" / "PRO PLUS" ("8+128GB" fica intacto: tem dígito a seguir)
    t = re.sub(r'(?<=\w)\+(?=\s|$)', ' PLUS', t)
    return _match_prep(t)

# Conectividade normalizada: LTE ≡ 4G; "WIFI" sozinho ≡ sem marcador (a versão base de um
# tablet é WiFi quer o fornecedor o escreva quer não); "WIFI + 5G" ≡ "5G".
_CONN_ALL_RE = re.compile(r'\b(5G|4G|LTE|CELL|CELLULAR|WIFI)\b')
_CONN_ALIAS  = {"LTE": "4G", "CELL": "5G", "CELLULAR": "5G"}
def _conn_set(prepped):
    c = {_CONN_ALIAS.get(m, m) for m in _CONN_ALL_RE.findall(prepped)}
    c.discard("WIFI")
    return frozenset(c)

# same_product() — decide se duas descrições de fornecedores diferentes representam
# o mesmo produto. Aplica uma série de guardas antes do fuzzy, da mais barata à mais cara:
#   1. Tier guard     — ULTRA ≠ PLUS ≠ PRO (devem coincidir dos dois lados)
#   2. Model number   — A17 ≠ A27, S24 ≠ S25
#   3. Version number — NOTE 14 ≠ NOTE 15
#   4. Color guard    — PRETO ≠ BRANCO quando ambos os lados declaram cor
#   5. Connectivity   — 4G vs 5G explícito = produtos diferentes;
#                       5G num lado / silêncio no outro = mesmo produto (só para
#                       iPhone/Pixel que são sempre 5G e o fornecedor omite o marcador)
#   6. Capacidade + Dice ≥ 90%
# _feat() — todas as características que same_product compara, calculadas UMA vez por
# descrição (cache): same_product corre ~120 mil vezes por consolidação e cada guard
# repetia os mesmos regex sobre o mesmo texto.
_MM_RE = re.compile(r'\b(\d{2})\s?MM\b')
@functools.lru_cache(maxsize=65536)
def _feat(txt):
    m = _prep(txt)
    words = m.split()
    fz = _fuzzy_key(m)
    return (
        m,
        frozenset(x for x, rx in _TIER_RES.items() if rx.search(m)),      # tier
        frozenset(x.group(0) for x in _MODEL_NUM_RE.finditer(m)),         # nº modelo
        frozenset(_MM_RE.findall(m)),                                     # tamanho relógio
        frozenset(x.group(0) for x in _VERSION_NUM_RE.finditer(m)),       # versão
        frozenset(w for w in words if w in _COLOR_WORDS),                 # cores
        _conn_set(m),                                                     # conectividade
        fz,                                                               # chave fuzzy
        frozenset(fz.lower().split()),                                    # palavras da chave
        capacity_tokens(m),                                               # capacidades
    )

def same_product(a, b):
    am, ta, ma, sa, va, ca, conn_a, am_fz, wa, capa = _feat(a)
    bm, tb, mb, sb, vb, cb, conn_b, bm_fz, wb, capb = _feat(b)
    # Tier guard: ULTRA/PLUS/PRO etc. must match exactly on both sides
    if ta != tb:
        return False
    # Model number guard: A17 ≠ A27, S24 ≠ S25, P30 ≠ P40
    if ma and mb and ma != mb:
        return False
    # Size guard (relógios): 40MM ≠ 44MM ≠ 46MM
    if sa and sb and sa != sb:
        return False
    # Version number guard: NOTE 14 ≠ NOTE 15, RUGKING 3 ≠ RUGKING 4
    if va and vb and va != vb:
        return False
    # Color guard: if both descriptions carry color words and the sets differ → different product
    if ca and cb and not _colors_mergeable(ca, cb, am_fz):
        return False
    # Connectivity guard (LTE ≡ 4G, "WIFI" ≡ sem marcador — ver _conn_set)
    if conn_a and conn_b and conn_a != conn_b:
        return False  # explicit conflict: 4G vs 5G → different products
    # Um lado marca 5G e o outro cala-se, com o mesmo modelo → só é SKU diferente quando
    # algum fornecedor distingue variantes desse modelo pelo marcador (ver _AMBIG_CONN_KEYS:
    # "Galaxy A17" vs "A17 5G", "iPad Air WiFi" vs "WiFi + 5G"). Caso contrário o "5G" é
    # decoração do fornecedor que o escreve sempre e o produto é o mesmo (S25, XCover).
    if bool(conn_a) != bool(conn_b) and (am_fz in _AMBIG_CONN_KEYS or bm_fz in _AMBIG_CONN_KEYS):
        return False
    # Fuzzy compara só a identidade do modelo (Dice ao nível de palavras ≥ 90): capacidades,
    # cores e conectividade saem porque cada uma já tem o seu guard.
    if not capacity_match(capa, capb):
        return False
    if not wa or not wb:
        return False
    return 2 * len(wa & wb) * 100 >= 90 * (len(wa) + len(wb))

# _dedup() — deduplicação em duas passagens (get_final: _prefilter → _dedup):
#
# 1ª passagem (best dict, O(n)):
#   Itera do mais barato para o mais caro.
#   Agrupa por EAN → PN → same_product() → nome exacto.
#   Para cada chave guarda sempre o registo mais barato; se aparecer um mais caro
#   com o mesmo EAN/PN mas com marcador "5G" que o mais barato não tem, enriquece
#   a descrição guardada (um fornecedor omite "5G" e outro inclui).
#
# 2ª passagem (by_bucket, O(n × bucket)):
#   Após a 1ª passagem, produtos com EANs diferentes mas mesma descrição (variantes
#   regionais, nomenclaturas cruzadas) ainda estão separados. Agrupa por
#   (categoria, capacidade) para reduzir o bucket a ~10-50 produtos e corre
#   same_product() dentro de cada bucket.
# _prefilter() — tudo o que sai do catálogo ANTES da deduplicação. Separado de _dedup()
# para audit_catalog() poder verificar o resultado contra o mesmo conjunto de candidatos.
def _prefilter(df):
    # Descarta acessórios logo à entrada, usando o tipo declarado pelo fornecedor.
    # Sem isto uma capa "Back Cover ... Galaxy S24" cairia em SMARTPHONES por ter "GALAXY".
    if "TIPO_FONTE" in df.columns:
        tipo = df["TIPO_FONTE"].astype(str).map(_strip_accents)
        df = df[~tipo.isin(ACCESSORY_TYPES)].copy()
    # Recondicionados fora: verifica descrição E marca ("Apple-Refurbished"),
    # para apanhar também os que não trouxerem a palavra no nome do produto.
    refurb = (df["DESCRICAO"].astype(str).str.contains(_REFURB_RE, na=False)
              | df["MARCA"].astype(str).str.contains(_REFURB_RE, na=False))
    if refurb.any():
        print(f"[FILTRO] {int(refurb.sum())} produtos recondicionados excluídos")
        df = df[~refurb].copy()
    # Modelos em que algum fornecedor distingue variantes 4G/5G (ou WiFi/5G) pelo marcador:
    # o mesmo _fuzzy_key aparece nesse fornecedor com E sem marcador. Usado por same_product().
    # Calculado ANTES dos filtros de preço/stock: uma variante esgotada continua a provar
    # que o modelo existe nas duas versões.
    _AMBIG_CONN_KEYS.clear()
    _MULTI_SHADE_KEYS.clear()
    _colors_mergeable.cache_clear()   # depende de _MULTI_SHADE_KEYS, que muda aqui
    for _, grp in df.groupby("FONTE"):
        marked, silent = set(), set()
        shades: dict = {}          # (modelo, cor base) → conjuntos de cor vistos neste fornecedor
        for d in grp["DESCRICAO"].astype(str):
            p = _prep(d)
            key = _fuzzy_key(p)
            (marked if _conn_set(p) else silent).add(key)
            cols = frozenset(w for w in p.split() if w in _COLOR_WORDS)
            for base in cols & _BASE_COLORS:
                shades.setdefault((key, base), set()).add(cols)
        _AMBIG_CONN_KEYS.update(marked & silent)
        _MULTI_SHADE_KEYS.update(k for k, s in shades.items() if len(s) > 1)
    print(f"[DEDUP] {len(_AMBIG_CONN_KEYS)} modelos com variantes 4G/5G distinguidas por marcador; "
          f"{len(_MULTI_SHADE_KEYS)} pares modelo/cor com várias tonalidades")

    df["EAN"]       = df["EAN"].apply(clean_ean)
    df["DESCRICAO"] = df["DESCRICAO"].apply(norm)
    tipo_col = df["TIPO_FONTE"] if "TIPO_FONTE" in df.columns else pd.Series("", index=df.index)
    df["CATEGORIA"] = [category(d, t) for d, t in zip(df["DESCRICAO"], tipo_col)]
    df = df[df["CATEGORIA"].notna()].copy()
    df = df[~df["DESCRICAO"].apply(ignored)].copy()
    # Só entram produtos das marcas listadas em BRANDS — a lista é uma AUTORIZAÇÃO, não só
    # uma etiqueta. Quem não for reconhecido (detect_brand → "OUTRAS") fica fora, venha do
    # fornecedor que vier: tirar uma marca de BRANDS deixa de a importar. Ver também
    # _PHONE_ONLY_BRANDS, que usa a mesma lista para classificar.
    df["MARCA"] = (df["MARCA"].astype(str) + " " + df["DESCRICAO"].astype(str)).apply(detect_brand)
    nao_listada = df["MARCA"] == "OUTRAS"
    if nao_listada.any():
        print(f"[FILTRO] {int(nao_listada.sum())} produtos de marcas fora de BRANDS excluídos")
        df = df[~nao_listada].copy()
    df["PRECO"] = pd.to_numeric(df["PRECO"], errors="coerce").fillna(0.0)
    df = df[df["PRECO"] > 0].copy()
    # Inclui produtos com stock disponível OU com chegada prevista; exclui o resto
    if "STOCK" in df.columns:
        stock_num   = pd.to_numeric(df["STOCK"], errors="coerce").fillna(0)
        chegada_num = pd.to_numeric(df["STOCK_CHEGADA"], errors="coerce").fillna(0) if "STOCK_CHEGADA" in df.columns else pd.Series(0, index=df.index)
        df = df[(stock_num > 0) | (chegada_num > 0)].copy()
    return df

_CONN_MARKER_RES = [(m, re.compile(r'\b' + m + r'\b')) for m in ("5G", "4G", "LTE")]

# _bucket_of() — chave de agrupamento partilhada pelas duas passagens do dedup:
# (categoria, maior capacidade de armazenamento). Mantém cada bucket com dezenas de
# produtos em vez de milhares, e evita repetir _prep()/capacity_tokens() por passagem.
def _bucket_of(r):
    caps = capacity_tokens(_prep(r["DESCRICAO"]))
    return (str(r.get("CATEGORIA", "")), max(map(_cap_gb, caps)) if caps else 0)

def _dedup(df):
    # 1ª passagem: guarda o registo mais barato por EAN / PN / descrição.
    # to_dict("records") em vez de iterrows(): evita construir uma Series pandas por linha.
    best = {}
    no_id_buckets: dict = {}          # (categoria, armazenamento) → chaves sem EAN/PN
    for r in df.sort_values("PRECO").to_dict("records"):
        key = None
        if r["EAN"]:
            key = "EAN:" + r["EAN"]
        elif str(r["PN"]).strip() not in ("", "nan", "0"):
            key = "PN:" + str(r["PN"]).strip()
        if key is None:
            # Sem EAN nem PN: compara por descrição, mas só contra os outros registos sem
            # identificador do mesmo bucket (categoria + armazenamento). Antes varria o
            # dicionário inteiro — com um fornecedor sem códigos de barras isso é O(n²).
            bkt = no_id_buckets.setdefault(_bucket_of(r), [])
            for k in bkt:
                if same_product(r["DESCRICAO"], best[k]["DESCRICAO"]):
                    key = k; break
            if key is None:
                key = "NAME:" + r["DESCRICAO"]
                bkt.append(key)
        if key not in best or r["PRECO"] < best[key]["PRECO"]:
            best[key] = dict(r)
        elif key.startswith(("EAN:", "PN:")):
            # Same EAN/PN but more expensive: enrich stored description with
            # connectivity info it may be missing (e.g. cheapest is "IPHONE 16E 128GB PRETO",
            # more expensive has "5G" → stored becomes "IPHONE 16E 128GB PRETO 5G")
            candidate = str(r["DESCRICAO"])
            # ...e com a tonalidade da cor, se o mais barato só trouxer a cor base
            stored = _enrich_color(best[key]["DESCRICAO"], candidate)
            for marker in ("5G", "4G", "LTE"):
                pat = re.compile(r'\b' + marker + r'\b')
                if pat.search(candidate) and not pat.search(stored):
                    stored = stored + " " + marker
                    break
            best[key]["DESCRICAO"] = stored

    # 2ª passagem: fuzzy dedup para produtos com EAN/PN distintos mas mesma descrição
    # in the first pass but are actually the same item (regional variants,
    # cross-supplier different naming). Bucketed by (category + capacity)
    # so each bucket is small (10-50 products) → total O(n * bucket_size).
    by_bucket: dict = {}
    for r in sorted(best.values(), key=lambda x: x["PRECO"]):
        bucket = by_bucket.setdefault(_bucket_of(r), [])
        dup = next((e for e in bucket if same_product(r["DESCRICAO"], e["DESCRICAO"])), None)
        if dup is None:
            bucket.append(r)
        else:
            # O vencedor (mais barato, já no bucket) pode ter só a cor base ("AZUL");
            # se o duplicado traz a tonalidade ("AZUL ULTRAMARINO"), herda-a.
            dup["DESCRICAO"] = _enrich_color(dup["DESCRICAO"], r["DESCRICAO"])
    result = pd.DataFrame([r for lst in by_bucket.values() for r in lst])
    if result.empty: return result
    result["PRECO"]     = pd.to_numeric(result["PRECO"], errors="coerce").fillna(0.0)
    result["DESCRICAO"] = result["DESCRICAO"].apply(clean_desc)
    result = result.sort_values("PRECO").drop_duplicates(subset=["DESCRICAO"], keep="first")
    return result

# ── auditoria pós-consolidação ────────────────────────────────────────────────
# Corre depois de cada consolidação (logo, depois de cada upload/sync) e responde a:
#   1. Entrou algum acessório (capa, carregador, película…) ou recondicionado?
#   2. Ficou algum produto duplicado entre fornecedores?
#   3. O preço escolhido é mesmo o mais baixo dos fornecedores que têm o produto?
#   4. Há preços suspeitos (telemóvel a 5 €)?
# Também calcula, por produto, o preço dos OUTROS fornecedores (coluna PRECOS_OUTROS).
_ACC_AUDIT_RE = re.compile(
    r'\b(CAPA|CAPAS|COVER|CASE|FUNDA|CARREGADOR|CHARGER|CABO|CABLE|PEL[IÍ]CULA|PROTECTOR|PROTETOR|'
    r'SCREEN PROTECTOR|TEMPERED GLASS|VIDRO TEMPERADO|POWERBANK|POWER BANK|AURICULAR|AURICULARES|'
    r'EARBUD|EARBUDS|HEADPHONE|AUSCULTADOR|SUPORTE|HOLDER|STYLUS|CANETA|ADAPTADOR|ADAPTER|DOCK)\b',
    re.IGNORECASE)
_SUSPECT_MIN_PRICE = {"SMARTPHONES": 30.0, "TABLETS": 60.0, "SMARTWATCHES": 20.0}
_AUDIT_MAX_EXAMPLES = 25

def audit_catalog(final, pool):
    """Verifica o catálogo final contra o conjunto de candidatos (pool = _prefilter(dados)).
    Devolve dict com contagens, exemplos e 'ok'. Acrescenta final['PRECOS_OUTROS']."""
    t0 = _time.time()
    rep = {"acessorios": [], "recondicionados": [], "duplicados": [],
           "preco_nao_minimo": [], "preco_suspeito": []}

    # 1. acessórios / recondicionados que tenham escapado
    for r in final.itertuples():
        d = str(r.DESCRICAO)
        tipo = _strip_accents(str(getattr(r, "TIPO_FONTE", "") or "").upper())
        # "INCLUI CARREGADOR + AURICULARES" (bundle) não faz do telemóvel um acessório
        d_chk = re.split(r'\b(?:INCLUI|INCLUYE|INCLUDES|WITH)\b', d)[0]
        if _ACC_AUDIT_RE.search(d_chk) or tipo in ACCESSORY_TYPES:
            rep["acessorios"].append({"desc": d, "fonte": r.FONTE, "preco": float(r.PRECO)})
        if _REFURB_RE.search(d) or _REFURB_RE.search(str(getattr(r, "MARCA", "") or "")):
            rep["recondicionados"].append({"desc": d, "fonte": r.FONTE, "preco": float(r.PRECO)})
        lim = _SUSPECT_MIN_PRICE.get(str(r.CATEGORIA), 0)
        if 0 < float(r.PRECO) < lim:
            rep["preco_suspeito"].append({"desc": d, "fonte": r.FONTE, "preco": float(r.PRECO)})

    # buckets do pool por (categoria, armazenamento) — mesma lógica da 2ª passagem do dedup
    def _bucket_key(cat, desc):
        p = _prep(desc)
        caps = capacity_tokens(p)
        return (cat, max(map(_cap_gb, caps)) if caps else 0), p

    # Linhas do pool cujo EAN já está no catálogo pertencem a um produto final (foram fundidas
    # por EAN, o critério mais fiável) — não podem ser "oferta mais barata perdida" de outro.
    def _ean(v):
        s = str(v).strip()
        return "" if s in ("", "nan", "None") else s
    final_eans = {_ean(e) for e in final["EAN"]} - {""}
    pool_b: dict = {}
    for r in pool.itertuples():
        key, p = _bucket_key(r.CATEGORIA, r.DESCRICAO)
        pool_b.setdefault(key, []).append((r.FONTE, float(r.PRECO), str(r.DESCRICAO), _ean(r.EAN)))

    # 2./3. por produto final: preços dos outros fornecedores e escolha do mínimo
    outros_col = []
    final_b: dict = {}
    for r in final.itertuples():
        key, p = _bucket_key(r.CATEGORIA, r.DESCRICAO)
        final_b.setdefault(key, []).append((r, p))
        f_ean = _ean(r.EAN)
        outros: dict = {}
        for fonte, preco, desc, ean in pool_b.get(key, []):
            if fonte == r.FONTE:
                continue
            if ean and ean == f_ean:
                mesmo = True                       # mesmo EAN: mesmo produto, sem dúvida
            elif ean and ean in final_eans:
                continue                           # pertence a OUTRO produto do catálogo
            else:
                mesmo = same_product(str(r.DESCRICAO), desc)
            if mesmo:
                outros[fonte] = min(preco, outros.get(fonte, float("inf")))
        outros_col.append(" · ".join(f"{k} {v:.2f}" for k, v in sorted(outros.items())))
        mais_barato = {k: v for k, v in outros.items() if v < float(r.PRECO) - 0.005}
        if mais_barato:
            rep["preco_nao_minimo"].append({"desc": str(r.DESCRICAO), "fonte": r.FONTE, "preco": float(r.PRECO),
                                            "outros": mais_barato})
    final["PRECOS_OUTROS"] = outros_col

    # 2. duplicados residuais no final: mesmo modelo/armazenamento, cor igual ou compatível,
    #    mesma conectividade declarada, capacidade compatível — mas de fontes diferentes
    for lst in final_b.values():
        for (a, pa), (b, pb) in itertools.combinations(lst, 2):
            if a.FONTE == b.FONTE or _fuzzy_key(pa) != _fuzzy_key(pb):
                continue
            ca = frozenset(w for w in pa.split() if w in _COLOR_WORDS)
            cb = frozenset(w for w in pb.split() if w in _COLOR_WORDS)
            if ca and cb and not _colors_mergeable(ca, cb, _fuzzy_key(pa)):
                continue
            if _conn_set(pa) != _conn_set(pb):
                continue
            if not capacity_match(capacity_tokens(pa), capacity_tokens(pb)):
                continue
            rep["duplicados"].append({"a": {"desc": str(a.DESCRICAO), "fonte": a.FONTE, "preco": float(a.PRECO)},
                                      "b": {"desc": str(b.DESCRICAO), "fonte": b.FONTE, "preco": float(b.PRECO)}})

    counts = {k: len(v) for k, v in rep.items()}
    problemas = counts["acessorios"] + counts["recondicionados"] + counts["duplicados"] + counts["preco_nao_minimo"]
    out = {
        "ok": problemas == 0,
        "problemas": problemas,
        "avisos": counts["preco_suspeito"],
        "contagens": counts,
        "exemplos": {k: v[:_AUDIT_MAX_EXAMPLES] for k, v in rep.items()},
        "total": int(len(final)),
        "por_fonte": {k: int(v) for k, v in final["FONTE"].value_counts().items()},
        "por_categoria": {k: int(v) for k, v in final["CATEGORIA"].value_counts().items()},
        "candidatos": int(len(pool)),
        "quando": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
        "duracao_s": round(_time.time() - t0, 1),
    }
    print(f"[AUDIT] {'OK' if out['ok'] else 'PROBLEMAS: ' + str(problemas)} — "
          + ", ".join(f"{k}={v}" for k, v in counts.items()) + f" ({out['duracao_s']}s)")
    return out

def clean_ean(v):
    if pd.isna(v): return ""
    s = str(v).strip()
    try:    s = str(int(float(s)))
    except (ValueError, OverflowError): s = re.sub(r"\D", "", s)
    return s.lstrip("0")

# ── excel ─────────────────────────────────────────────────────────────────────
def build_excel(final, tipo):
    sale_col = "PRECO_VIP" if tipo == "vip" else "PRECO_CLIENTE"
    sale_lbl = "Preço VIP C/IVA" if tipo == "vip" else "Preço C/IVA"

    HDR_FILL   = PatternFill("solid", fgColor="1B4FA0")
    HDR_FONT   = Font(bold=True, color="FFFFFF", size=11)
    BRAND_FILL = PatternFill("solid", fgColor="E8F0FB")
    BRAND_FONT = Font(bold=True, color="1B4FA0", size=14)
    ALT_FILL   = PatternFill("solid", fgColor="F7F8FA")
    PROD_FONT  = Font(color="111827", size=10)
    PRICE_FONT = Font(bold=True, color="000000", size=10)
    AL_L = Alignment(horizontal="left",  vertical="center")
    AL_R = Alignment(horizontal="right", vertical="center")
    AL_C = Alignment(horizontal="center",vertical="center")
    FMT  = '#,##0 "€"'

    wb_buf = io.BytesIO()
    with pd.ExcelWriter(wb_buf, engine="openpyxl") as writer:
        for cat in ["SMARTPHONES", "TABLETS", "SMARTWATCHES"]:
            rows   = []
            # Dentro de cada marca: agrupado por modelo (todas as variantes do Galaxy S25
            # seguidas), e dentro do modelo da menor para a maior capacidade e depois pelo
            # preço — em vez da lista corrida por valor, que espalhava o mesmo modelo pela
            # folha toda.
            subset = final[final["CATEGORIA"] == cat].copy()
            if not subset.empty:
                prep = subset["DESCRICAO"].apply(_prep)
                subset["_MODELO"]   = prep.apply(_fuzzy_key)
                subset["_ARMAZEM"]  = prep.apply(
                    lambda p: max(map(_cap_gb, capacity_tokens(p))) if capacity_tokens(p) else 0)
                subset = subset.sort_values(["MARCA", "_MODELO", "_ARMAZEM", sale_col])
            for marca in sorted(subset["MARCA"].dropna().unique()):
                rows.append({"DESCRIÇÃO": marca, sale_lbl: None})
                for _, r in subset[subset["MARCA"] == marca].iterrows():
                    venda = int(float(r[sale_col]) + 0.5) if pd.notna(r[sale_col]) and r[sale_col] > 0 else None
                    if venda is None: continue
                    rows.append({"DESCRIÇÃO": r["DESCRICAO"], sale_lbl: venda})
                rows.append({"DESCRIÇÃO": None, sale_lbl: None})
            pd.DataFrame(rows).to_excel(writer, sheet_name=cat, index=False)

    wb_buf.seek(0)
    wb = load_workbook(wb_buf)

    label_vip = "TABELA VIP" if tipo == "vip" else "TABELA REVENDA"
    exported  = datetime.now().strftime("%d/%m/%Y")
    for ws in wb.worksheets:
        ws.insert_rows(1)
        ws.row_dimensions[1].height = 42
        ws.row_dimensions[2].height = 22

        ws.merge_cells("A1:B1")
        title_cell = ws.cell(row=1, column=1)
        title_cell.value     = f"{app_settings_store.get('company_name', DEFAULT_COMPANY)} · {label_vip} · {ws.title} · {exported}"
        title_cell.font      = Font(bold=True, color="0D2B6E", size=18)
        title_cell.alignment = Alignment(horizontal="left", vertical="center")

        for cell in ws[2]:
            cell.font = HDR_FONT; cell.fill = HDR_FILL; cell.alignment = AL_C

        for i, row in enumerate(ws.iter_rows(min_row=3), start=3):
            v_a = row[0].value
            v_b = row[1].value
            is_brand = v_a and str(v_a).strip() and (v_b is None or v_b == "")
            if is_brand:
                row[0].font = BRAND_FONT; row[0].alignment = AL_L
                for cell in row: cell.fill = BRAND_FILL
            elif v_a and str(v_a).strip():
                row[0].font = PROD_FONT;  row[0].alignment = AL_L
                row[1].font = PRICE_FONT; row[1].alignment = AL_R
                row[1].number_format = FMT
                if i % 2 == 0:
                    for cell in row: cell.fill = ALT_FILL

        ws.column_dimensions["A"].width = 70
        ws.column_dimensions["B"].width = 18

    out = io.BytesIO()
    wb.save(out); out.seek(0)
    return out

# ── estado da sessão ──────────────────────────────────────────────────────────
MARGINS_FILE   = os.path.join(BASE_DIR, "margins.json")
SETTINGS_FILE  = os.path.join(BASE_DIR, "settings.json")
DEFAULT_COMPANY = "A Minha Empresa"   # nome por omissão; muda-se em Definições
STATIC_DIR     = os.path.join(BASE_DIR, "static")

def _load_settings():
    defaults = {"company_name": DEFAULT_COMPANY, "auto_sync": True}
    try:
        with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        defaults.update({k: data[k] for k in defaults if k in data})
    except (FileNotFoundError, ValueError):
        pass
    return defaults

def _save_settings():
    with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
        json.dump(app_settings_store, f, ensure_ascii=False, indent=2)

def _load_margins():
    try:
        with open(MARGINS_FILE, "r") as f:
            data = json.load(f)
        return {"normal": float(data.get("normal", 10.0)),
                "vip":    float(data.get("vip",     5.0))}
    except (FileNotFoundError, ValueError, KeyError):
        return {"normal": 10.0, "vip": 5.0}

uploaded           = {}
margins            = _load_margins()
app_settings_store = _load_settings()
suppliers_config   = _load_suppliers()
overrides          = {k: {} for k in suppliers_config}
sync_state         = {k: {"status": "idle", "last_sync": None, "error": None} for k in suppliers_config}
_final_cache: dict = {"df": None, "valid": False}   # invalidated on every sync

# Protege o estado partilhado do catálogo — uploaded, overrides, margins e _final_cache —
# contra acessos simultâneos. O servidor corre com threaded=True e a thread de auto_sync()
# substitui o DataFrame de um fornecedor a qualquer momento: sem este lock, um /generate a
# decorrer durante um sync podia consolidar metade dos fornecedores com dados antigos e a
# outra metade com dados novos, produzindo um Excel com preços de duas gerações misturados.
# É reentrante (RLock) porque get_final() chama funções que voltam a adquiri-lo.
# Regra: nunca segurar o lock durante I/O lento (download HTTP, leitura do ficheiro) — só à
# volta do bloco que publica o resultado no estado partilhado.
_catalog_lock = threading.RLock()

# ── download automático ───────────────────────────────────────────────────────
def _detect_ext(url, content_type):
    u, ct = url.lower(), content_type.lower()
    if u.endswith(".xlsx") or "excel" in u or "spreadsheetml" in ct or "ms-excel" in ct:
        return ".xlsx"
    if u.endswith(".xls"): return ".xls"
    return ".csv"

class _NoRetry(Exception):
    """Falha definitiva (ex.: URL exige login) — não repetir o download."""

# download_provider() — descarrega o ficheiro de um fornecedor (3 tentativas com backoff).
# Guarda primeiro num .tmp e só substitui o ficheiro real após validação, para que um
# download interrompido não corrompa os dados em uso.
def download_provider(src):
    import glob as _glob

    sync_state.setdefault(src, {"status": "idle", "last_sync": None, "error": None})
    url = suppliers_config.get(src, {}).get("url", "")
    if not url:
        sync_state[src].update(
            status="error",
            error="Sem URL — define o URL do fornecedor ou carrega o ficheiro em modo Manual")
        return
    sync_state[src].update(status="downloading", error=None)

    last_exc = None
    for attempt in range(1, 4):   # 3 tentativas: imediata, +10s, +30s
        try:
            r = req_lib.get(url, timeout=90)
            # Página web em vez de ficheiro (login/token em falta) — não vale a pena repetir
            head = r.content[:256].lstrip().lower()
            if r.status_code in (401, 403, 404) or head.startswith((b"<!doctype", b"<html")):
                raise _NoRetry(f"HTTP {r.status_code}: o URL devolveu uma página web em vez do "
                               f"ficheiro — o link exige login ou token")
            r.raise_for_status()
            if len(r.content) < 64:
                raise ValueError(f"Resposta do servidor parece vazia ({len(r.content)} bytes)")
            ext   = _detect_ext(url, r.headers.get("content-type", ""))
            fname = f"{src}_auto{ext}"
            path  = os.path.join(UPLOAD_FOLDER, fname)
            tmp   = os.path.join(UPLOAD_FOLDER, f"{src}_{os.getpid()}.tmp{ext}")
            with open(tmp, "wb") as f: f.write(r.content)
            df = read_file(tmp)
            if df.empty or len(df.columns) == 0:
                os.remove(tmp)
                raise ValueError("Ficheiro descarregado não contém dados válidos")
            os.replace(tmp, path)
            price_col = find_col(df, *PRICE_KWS)
            desc_col  = find_col(df, *DESC_COLS)
            with _catalog_lock:
                uploaded[src] = {"path": path, "name": fname, "rows": len(df),
                                 "cols": list(df.columns), "price_col": price_col,
                                 "desc_col": desc_col, "df": df}
                overrides[src] = {}
                _final_cache["valid"] = False
            sync_state[src].update(status="ok", last_sync=datetime.now().strftime("%H:%M:%S"))
            print(f"[SYNC {src.upper()}] {len(df)} linhas | price={price_col} | ext={ext}"
                  + (f" (tentativa {attempt})" if attempt > 1 else ""))
            return
        except Exception as exc:
            last_exc = exc
            # Clean up any tmp files from this attempt
            for stale in _glob.glob(os.path.join(UPLOAD_FOLDER, f"{src}_*.tmp*")):
                try: os.remove(stale)
                except OSError: pass
            if isinstance(exc, _NoRetry):
                break
            if attempt < 3:
                wait = 10 * attempt   # 10s depois da 1ª falha, 20s depois da 2ª
                print(f"[SYNC {src.upper()}] tentativa {attempt} falhou ({exc}), a aguardar {wait}s...")
                sync_state[src].update(status="a repetir...", error=str(exc))
                _time.sleep(wait)

    # Todas as tentativas falharam — o ficheiro anterior (se existir) continua em uso
    traceback.print_exc()
    msg = str(last_exc)
    if src in uploaded:
        msg += f" · a usar o último ficheiro carregado ({uploaded[src]['rows']} linhas)"
    sync_state[src].update(status="error", error=msg)

_SYNC_INTERVAL = 24 * 3600  # refresh automático a cada 24 horas

# auto_sync() — thread em background que descarrega todos os fornecedores automáticos
# e depois dorme 24h em loop. A 1ª execução acontece imediatamente ao arrancar.
def auto_sync():
    while True:
        for src, cfg in suppliers_config.items():
            if cfg.get("mode", "auto") == "auto" and cfg.get("url"):
                download_provider(src)
        _time.sleep(_SYNC_INTERVAL)

def _load_cached():
    # Carrega os ficheiros já descarregados do disco para que o UI tenha dados
    # imediatamente ao arrancar (sem esperar o primeiro ciclo de auto_sync).
    # Limpa ficheiros .tmp de execuções anteriores (ficam quando o servidor é interrompido a meio de um download)
    import glob as _g
    for stale in _g.glob(os.path.join(UPLOAD_FOLDER, "*.tmp*")):
        try:
            os.remove(stale)
            print(f"[CACHE] removido ficheiro temporario: {os.path.basename(stale)}")
        except OSError:
            pass

    for src in suppliers_config:
        if src in uploaded:
            continue
        # Primeiro os downloads automáticos ({src}_auto.*); depois, se não houver,
        # o upload manual mais recente ({src}_<nome do ficheiro>) — senão um ficheiro
        # carregado à mão perdia-se sempre que o servidor reiniciava.
        cands = [os.path.join(UPLOAD_FOLDER, f"{src}_auto{e}") for e in (".xlsx", ".xls", ".csv")]
        cands = [p for p in cands if os.path.exists(p)]
        if not cands:
            manual = [p for e in (".xlsx", ".xls", ".csv")
                        for p in _g.glob(os.path.join(UPLOAD_FOLDER, f"{src}_*{e}"))]
            cands  = sorted(manual, key=os.path.getmtime, reverse=True)[:1]
        for p in cands:
            try:
                df = read_file(p)
                if df.empty or len(df.columns) == 0:
                    continue
                price_col = find_col(df, *PRICE_KWS)
                desc_col  = find_col(df, *DESC_COLS)
                uploaded[src] = {"path": p, "name": os.path.basename(p),
                                 "rows": len(df), "cols": list(df.columns),
                                 "price_col": price_col, "desc_col": desc_col, "df": df}
                print(f"[CACHE] {src} carregado da cache: {len(df)} linhas")
                break
            except Exception as e:
                print(f"[CACHE] {src} falhou a carregar cache: {e}")

# Arranca apenas no processo principal (evita duplicar em modo debug com reloader)
if os.environ.get('WERKZEUG_RUN_MAIN') == 'true' or not os.environ.get('WERKZEUG_RUN_MAIN'):
    _load_cached()                                        # carrega cache do disco primeiro
    threading.Thread(target=auto_sync, daemon=True).start()  # refresh a cada 24h em background

# ── pipeline principal ────────────────────────────────────────────────────────
def _get_df(src):
    # Usa o DataFrame em memória se disponível; relê do disco caso contrário
    entry = uploaded[src]
    df = entry.get("df")
    return df if df is not None else read_file(entry["path"])

def _loaded_sources():
    # Fornecedores configurados que já têm ficheiro carregado (ordem do suppliers_config)
    return [s for s in suppliers_config if s in uploaded]

# get_final() — ponto de entrada do pipeline: map_source → consolidate → margens.
# _final_cache é invalidado em cada sync para garantir dados frescos.
def get_final(force=False):
    with _catalog_lock:
        if not force and _final_cache["valid"] and _final_cache["df"] is not None:
            return _final_cache["df"]
        srcs = _loaded_sources()
        if not srcs:
            raise ValueError("Nenhum ficheiro carregado")
        dados = pd.concat([
            map_source(_get_df(s), s, overrides.get(s, {})) for s in srcs
        ], ignore_index=True)
        pool  = _prefilter(dados)
        final = _dedup(pool)
        if final.empty:
            _final_cache.update(df=final, valid=True, audit=None)
            return final
        final["MARCA"]         = final.apply(
            lambda x: detect_brand(str(x.get("MARCA","")) + " " + str(x["DESCRICAO"])), axis=1)
        final["PRECO_CLIENTE"] = (final["PRECO"] * (1 + margins["normal"] / 100)).round(2)
        final["PRECO_VIP"]     = (final["PRECO"] * (1 + margins["vip"]    / 100)).round(2)
        # Verificação automática de cada consolidação (acessórios, recondicionados, duplicados,
        # preço mínimo) — resultado exposto em /audit e /preview.
        _final_cache["audit"] = audit_catalog(final, pool)
        watch_mask = final["CATEGORIA"] == "SMARTWATCHES"
        if watch_mask.any():
            final.loc[watch_mask, "DESCRICAO"] = final.loc[watch_mask, "DESCRICAO"].apply(clean_watch_desc)
        _final_cache.update(df=final, valid=True)
        return final

# ── rotas ─────────────────────────────────────────────────────────────────────
@app.route("/")
def index():
    r = render_template("index.html")
    from flask import make_response
    resp = make_response(r)
    resp.headers["Cache-Control"] = "no-store, no-cache, must-revalidate"
    return resp

@app.route("/sync", methods=["POST"])
def sync():
    data    = request.get_json() or {}
    src     = data.get("source")
    # Sync global (sem "source") só toca nos automáticos — os manuais não têm URL
    # e ficariam com um erro "URL não configurado" no cartão.
    sources = [src] if src in suppliers_config else [
        s for s, c in suppliers_config.items() if c.get("mode", "auto") == "auto"
    ]
    for s in sources:
        threading.Thread(target=download_provider, args=(s,), daemon=True).start()
    return jsonify({"ok": True, "sources": sources})

@app.route("/sync_status")
def sync_status():
    result = {}
    for src in suppliers_config:
        result[src] = {**sync_state.get(src, {"status": "idle", "last_sync": None, "error": None})}
        if src in uploaded:
            result[src]["rows"] = uploaded[src]["rows"]
            result[src]["name"] = uploaded[src]["name"]
    return jsonify({"ok": True, "state": result})

@app.route("/upload", methods=["POST"])
def upload():
    src = request.form.get("source")
    f   = request.files.get("file")
    if not f or src not in suppliers_config:
        return jsonify({"ok": False, "error": "Parâmetros inválidos"}), 400
    fname = os.path.basename(f.filename or "upload")
    fname = re.sub(r'[^\w.\-() ]', '_', fname) or "upload"
    path  = os.path.join(UPLOAD_FOLDER, f"{src}_{fname}")
    tmp   = path + ".tmp"
    f.save(tmp)
    try:
        df        = read_file(tmp)
        price_col = find_col(df, *PRICE_KWS)
        desc_col  = find_col(df, *DESC_COLS)
        if df.empty or not price_col:
            os.remove(tmp)
            return jsonify({"ok": False, "error": "Ficheiro vazio ou sem coluna de preço — o anterior mantém-se"}), 400
        # Só depois de o novo ficheiro ler bem: apaga TODOS os ficheiros antigos deste
        # fornecedor (uploads manuais anteriores e o *_auto.*) e fica só o novo.
        removed = _purge_supplier_files(src, keep=tmp)
        os.replace(tmp, path)
        with _catalog_lock:
            uploaded[src] = {"path": path, "name": fname, "rows": len(df),
                             "cols": list(df.columns), "price_col": price_col,
                             "desc_col": desc_col, "df": df}
            overrides[src] = {}
            # Upload manual bem-sucedido limpa qualquer erro de sync anterior
            sync_state[src] = {"status": "ok", "last_sync": None, "error": None}
            _clear_caches()
        print(f"[UPLOAD] {src}: {fname} ({len(df)} linhas) · {removed} ficheiro(s) antigo(s) apagado(s) · cache limpa")
        return jsonify({"ok": True, "name": fname, "rows": len(df), "removed": removed,
                        "cols": list(df.columns), "price_col": price_col, "desc_col": desc_col})
    except Exception:
        traceback.print_exc()
        try:
            os.remove(tmp)
        except OSError:
            pass
        return jsonify({"ok": False, "error": traceback.format_exc().splitlines()[-1]}), 500

def _purge_supplier_files(src, keep=None):
    # Apaga todos os ficheiros {src}_* da pasta uploads excepto `keep`.
    import glob as _g
    n = 0
    for p in _g.glob(os.path.join(UPLOAD_FOLDER, f"{src}_*")):
        if keep and os.path.abspath(p) == os.path.abspath(keep):
            continue
        try:
            os.remove(p); n += 1
        except OSError as e:
            print(f"[UPLOAD] não consegui apagar {os.path.basename(p)}: {e}")
    return n

def _clear_caches():
    # Catálogo consolidado + auditoria + cache de traduções: tudo recalculado no próximo pedido.
    _final_cache.update(df=None, valid=False, audit=None)
    translate_es_pt.cache_clear()

@app.route("/reload/<key>", methods=["POST"])
def reload_supplier(key):
    """Reprocessa o ficheiro que já está carregado, sem obrigar a carregar outro.
    Serve para voltar a correr a consolidação sobre os mesmos dados — por exemplo depois
    de mudar margens ou regras de catálogo."""
    import glob as _g
    if key not in suppliers_config:
        return jsonify({"ok": False, "error": "Fornecedor não encontrado"}), 404
    path = (uploaded.get(key) or {}).get("path")
    if not path or not os.path.exists(path):
        # ficheiro já não está em memória (servidor reiniciado): procura o mais recente no disco
        cands = [f for e in (".xlsx", ".xls", ".csv")
                   for f in _g.glob(os.path.join(UPLOAD_FOLDER, f"{key}_*{e}"))]
        cands = sorted(cands, key=os.path.getmtime, reverse=True)
        if not cands:
            return jsonify({"ok": False, "error": "Não há ficheiro carregado para reprocessar"}), 400
        path = cands[0]
    try:
        df = read_file(path)
        if df.empty:
            return jsonify({"ok": False, "error": "O ficheiro está vazio"}), 400
        price_col = find_col(df, *PRICE_KWS)
        desc_col  = find_col(df, *DESC_COLS)
        if not price_col:
            return jsonify({"ok": False, "error": "Ficheiro sem coluna de preço"}), 400
        name = os.path.basename(path)
        with _catalog_lock:
            uploaded[key] = {"path": path, "name": name, "rows": len(df), "cols": list(df.columns),
                             "price_col": price_col, "desc_col": desc_col, "df": df}
            sync_state[key] = {"status": "ok", "last_sync": datetime.now().strftime("%H:%M:%S"), "error": None}
            _clear_caches()
        print(f"[RELOAD {key.upper()}] {name}: {len(df)} linhas relidas · cache limpa")
        return jsonify({"ok": True, "rows": len(df), "name": name})
    except Exception:
        traceback.print_exc()
        return jsonify({"ok": False, "error": traceback.format_exc().splitlines()[-1]}), 500

@app.route("/suppliers", methods=["GET"])
def get_suppliers():
    result = {}
    for k, cfg in suppliers_config.items():
        result[k] = {
            **cfg,
            "state": sync_state.get(k, {"status": "idle", "last_sync": None, "error": None}),
            "rows":  uploaded[k]["rows"] if k in uploaded else None,
            "file":  uploaded[k]["name"] if k in uploaded else None,
        }
    return jsonify({"ok": True, "suppliers": result})

@app.route("/suppliers", methods=["POST"])
def add_supplier():
    data = request.get_json() or {}
    name = data.get("name", "").strip()
    url  = data.get("url",  "").strip()
    mode = data.get("mode", "auto")
    if not name:
        return jsonify({"ok": False, "error": "Nome obrigatório"}), 400
    key = re.sub(r'[^a-z0-9]', '_', name.lower())[:20].strip('_')
    if not key:
        return jsonify({"ok": False, "error": "Nome inválido"}), 400
    if key in suppliers_config:
        key = key + "_" + str(len(suppliers_config))
    suppliers_config[key] = {"name": name, "url": url, "mode": mode}
    sync_state[key]  = {"status": "idle", "last_sync": None, "error": None}
    overrides[key]   = {}
    _save_suppliers()
    return jsonify({"ok": True, "key": key})

@app.route("/suppliers/<key>", methods=["PUT"])
def update_supplier(key):
    if key not in suppliers_config:
        return jsonify({"ok": False, "error": "Fornecedor não encontrado"}), 404
    data = request.get_json() or {}
    for field in ("name", "url", "mode"):
        if field in data:
            suppliers_config[key][field] = data[field]
    # Mudar de modo ou definir URL limpa o erro antigo (ex.: "Sem URL" ao passar para manual)
    if ("mode" in data or "url" in data) and sync_state.get(key, {}).get("error"):
        sync_state[key]["error"]  = None
        sync_state[key]["status"] = "ok" if key in uploaded else "idle"
    _final_cache["valid"] = False
    _save_suppliers()
    return jsonify({"ok": True})

@app.route("/suppliers/<key>", methods=["DELETE"])
def delete_supplier(key):
    if key not in suppliers_config:
        return jsonify({"ok": False, "error": "Fornecedor não encontrado"}), 404
    name = suppliers_config[key].get("name", key)
    with _catalog_lock:
        for store in (suppliers_config, sync_state, overrides, uploaded):
            store.pop(key, None)
        removed = _purge_supplier_files(key)         # apaga ficheiros em uploads/ deste fornecedor
        _clear_caches()
    _save_suppliers()
    print(f"[SUPPLIERS] removido {key} ({name}) · {removed} ficheiro(s) apagado(s)")
    return jsonify({"ok": True, "removed_files": removed})

@app.route("/set_mapping", methods=["POST"])
def set_mapping():
    data = request.get_json()
    src  = data.get("source")
    if src not in suppliers_config:
        return jsonify({"ok": False, "error": "source inválido"}), 400
    if src not in uploaded:
        return jsonify({"ok": False, "error": "Ficheiro não carregado"}), 400
    with _catalog_lock:
        overrides[src] = {k: v for k, v in data.items()
                          if k in ("desc","brand","pn","ean","price") and v}
        df        = _get_df(src)
        price_col = overrides[src].get("price") or find_col(df, *PRICE_KWS)
        uploaded[src]["price_col"] = price_col
        # Mudar a coluna de preço muda os preços de custo: o catálogo consolidado em cache
        # deixa de ser válido e tem de ser recalculado no próximo /preview ou /generate.
        _clear_caches()
    samples   = df[price_col].dropna().head(5).astype(str).tolist() if price_col else []
    return jsonify({"ok": True, "price_col": price_col, "sample_prices": samples})

@app.route("/cols/<src>")
def get_cols(src):
    if src not in uploaded:
        return jsonify({"ok": False, "error": "Ficheiro não carregado"}), 400
    try:
        df   = _get_df(src)
        cols = list(df.columns)
        samp = {c: str(df[c].dropna().iloc[0]) if df[c].dropna().shape[0] > 0 else "" for c in cols}
        return jsonify({"ok": True, "cols": cols, "samples": samp,
                        "price_col": uploaded[src].get("price_col")})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500

@app.route("/margins", methods=["GET","POST"])
def set_margins():
    if request.method == "GET":
        return jsonify({"ok": True, **margins})
    data = request.get_json(silent=True) or {}
    try:
        normal = float(data.get("normal", margins["normal"]))
        vip    = float(data.get("vip",    margins["vip"]))
    except (TypeError, ValueError):
        return jsonify({"ok": False, "error": "Margens têm de ser números"}), 400
    if not (0 <= normal <= 500 and 0 <= vip <= 500):
        return jsonify({"ok": False, "error": "Margens fora do intervalo 0–500 %"}), 400
    with _catalog_lock:
        margins["normal"], margins["vip"] = normal, vip
        _final_cache["valid"] = False  # preços de venda dependem das margens
    with open(MARGINS_FILE, "w") as f:
        json.dump(margins, f)
    return jsonify({"ok": True, **margins})

@app.route("/settings", methods=["GET","POST"])
def settings_route():
    logo_file = next((f for f in ["logo.png","logo.jpg","logo.jpeg","logo.webp"]
                      if os.path.exists(os.path.join(STATIC_DIR, f))), None)
    if request.method == "GET":
        return jsonify({"ok": True, **app_settings_store,
                        "logo_url": f"/static/{logo_file}" if logo_file else None})
    data = request.get_json() or {}
    if "company_name" in data:
        app_settings_store["company_name"] = str(data["company_name"])[:60].strip() or DEFAULT_COMPANY
    if "auto_sync" in data:
        app_settings_store["auto_sync"] = bool(data["auto_sync"])
    _save_settings()
    return jsonify({"ok": True, **app_settings_store})

@app.route("/logo", methods=["POST"])
def upload_logo():
    f = request.files.get("file")
    if not f: return jsonify({"ok": False, "error": "Sem ficheiro"}), 400
    ext = os.path.splitext(f.filename)[1].lower()
    if ext not in (".png",".jpg",".jpeg",".webp"):
        return jsonify({"ok": False, "error": "Usa PNG, JPG ou WEBP"}), 400
    for old in ["logo.png","logo.jpg","logo.jpeg","logo.webp"]:
        p = os.path.join(STATIC_DIR, old)
        if os.path.exists(p):
            try: os.remove(p)
            except: pass
    dest = os.path.join(STATIC_DIR, f"logo{ext}")
    f.save(dest)
    return jsonify({"ok": True, "logo_url": f"/static/logo{ext}"})

@app.route("/logo", methods=["DELETE"])
def delete_logo():
    for name in ["logo.png","logo.jpg","logo.jpeg","logo.webp"]:
        p = os.path.join(STATIC_DIR, name)
        if os.path.exists(p):
            try: os.remove(p)
            except: pass
    return jsonify({"ok": True})

@app.route("/preview")
def preview():
    if not _loaded_sources():
        return jsonify({"ok": False, "error": "Faltam ficheiros"}), 400
    try:
        final  = get_final()
        counts = final["CATEGORIA"].value_counts().to_dict()
        cols   = ["DESCRICAO","CATEGORIA","MARCA","EAN","PRECO","PRECO_CLIENTE","PRECO_VIP","FONTE",
                  "STOCK_CHEGADA","DATA_CHEGADA","PRECOS_OUTROS"]
        sample = final[[c for c in cols if c in final.columns]].to_dict("records")
        return jsonify({"ok": True,
            "smartphones":  int(counts.get("SMARTPHONES",  0)),
            "tablets":      int(counts.get("TABLETS",      0)),
            "smartwatches": int(counts.get("SMARTWATCHES", 0)),
            "total": len(final), "sample": sample,
            "audit": _final_cache.get("audit")})
    except Exception:
        traceback.print_exc()
        return jsonify({"ok": False, "error": traceback.format_exc().splitlines()[-1]}), 500

@app.route("/audit")
def audit_route():
    # Relatório da última verificação (recalcula o catálogo se for preciso)
    if not _loaded_sources():
        return jsonify({"ok": False, "error": "Faltam ficheiros"}), 400
    try:
        get_final()
        return jsonify({"ok": True, "audit": _final_cache.get("audit")})
    except Exception:
        traceback.print_exc()
        return jsonify({"ok": False, "error": traceback.format_exc().splitlines()[-1]}), 500

@app.route("/generate/<tipo>")
def generate(tipo):
    if tipo not in ("vip","normal"):
        return jsonify({"ok": False, "error": "tipo inválido"}), 400
    if not _loaded_sources():
        return jsonify({"ok": False, "error": "Faltam ficheiros"}), 400
    try:
        final = get_final()
        if final.empty:
            return jsonify({"ok": False, "error": "Nenhum produto encontrado."}), 400
        label = "VIP" if tipo == "vip" else "Normal"
        stamp = datetime.now().strftime("%Y-%m-%d_%H-%M")
        return send_file(build_excel(final, tipo), as_attachment=True,
                         download_name=f"Catalogo_{label}_{stamp}.xlsx",
                         mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    except Exception:
        traceback.print_exc()
        return jsonify({"ok": False, "error": traceback.format_exc().splitlines()[-1]}), 500

@app.route("/debug")
def debug():
    lines = ["<pre style='font:13px monospace;padding:1rem;background:#111;color:#eee'>"]
    for src in suppliers_config:
        if src in uploaded:
            u = uploaded[src]
            lines += [f"=== {src.upper()} ===",
                      f"Ficheiro : {u['name']}",
                      f"Linhas   : {u['rows']}",
                      f"Preço col: {u.get('price_col','NÃO DETETADA')}",
                      f"Colunas  : {u['cols']}", ""]
            try:    lines.append(_get_df(src).head(3).to_string())
            except Exception as e: lines.append(f"Erro: {e}")
            lines.append("")
        else:
            lines.append(f"=== {src.upper()} === NÃO CARREGADO\n")
    lines.append("</pre>")
    return "\n".join(lines)

@app.route("/debug/product")
def debug_product():
    q = request.args.get("q", "").strip().upper()
    if not q:
        return "<pre>Usa: /debug/product?q=SAMSUNG+S26+ULTRA</pre>"

    lines = ["<pre style='font:13px monospace;padding:1rem;background:#111;color:#eee;white-space:pre-wrap'>",
             f"PESQUISA: {q}\n{'='*60}"]

    # Mostra o que cada fornecedor tem em bruto
    for src in suppliers_config:
        if src not in uploaded:
            lines.append(f"\n[{src.upper()}] NÃO CARREGADO")
            continue
        try:
            raw = _get_df(src)
            mapped = map_source(raw, src, overrides.get(src, {}))
            mapped["DESCRICAO_NORM"] = mapped["DESCRICAO"].apply(norm)
            hits = mapped[mapped["DESCRICAO_NORM"].str.contains(q, na=False, regex=False)]
            lines.append(f"\n[{src.upper()}] {len(hits)} resultado(s):")
            for _, r in hits.head(10).iterrows():
                lines.append(f"  DESC : {r['DESCRICAO']}")
                lines.append(f"  EAN  : {r['EAN']}  |  PN: {r['PN']}")
                lines.append(f"  PRECO: {r['PRECO']}  |  STOCK: {r.get('STOCK','—')}")
                lines.append("")
        except Exception as e:
            lines.append(f"\n[{src.upper()}] ERRO: {e}")

    # Mostra o resultado consolidado
    lines.append(f"\n{'='*60}\nRESULTADO CONSOLIDADO:")
    try:
        final = get_final()
        hits = final[final["DESCRICAO"].str.upper().str.contains(q, na=False, regex=False)]
        for _, r in hits.head(10).iterrows():
            lines.append(f"  DESC     : {r['DESCRICAO']}")
            lines.append(f"  FONTE    : {r['FONTE']}")
            lines.append(f"  PRECO    : {r['PRECO']}")
            lines.append(f"  VIP      : {r['PRECO_VIP']}")
            lines.append(f"  REVENDA  : {r['PRECO_CLIENTE']}")
            lines.append("")
    except Exception as e:
        lines.append(f"ERRO: {e}")

    lines.append("</pre>")
    return "\n".join(lines)

if __name__ == "__main__":
# host="127.0.0.1": o servidor fica acessível apenas a partir desta própria máquina.
# Esta configuração é importante porque a aplicação não possui autenticação integrada.
# Se fosse disponibilizada na rede, qualquer pessoa com acesso poderia, sem necessidade
# de login, apagar fornecedores, alterar tabelas de preços, modificar margens e consultar
# os URLs dos fornecedores, que podem conter credenciais de acesso aos preços de custo.
#
# NÃO alterar para "0.0.0.0" nem disponibilizar através de router, port forwarding ou túnel
# sem implementar primeiro uma camada de segurança. Para acesso remoto, utilizar uma VPN
# (por exemplo, Tailscale) ou colocar a aplicação atrás de um proxy com autenticação e HTTPS.
#
# O servidor Flask utilizado neste projeto é destinado a desenvolvimento local e não deve
# ser diretamente exposto à Internet.
    app.run(host="127.0.0.1", debug=False, port=5050, threaded=True)
