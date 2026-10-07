"""
Product classification: agricultura familiar vs. abarrotes.
To add a new keyword or spelling variant, edit CULTIVADOS or ABARROTES.
Order matters: CULTIVADOS is checked first, and the first match wins.
"""
import re
from texto import normalize_text

CULTIVADOS = [
    # frutas
    'banano', 'bananano',                          # triple-n typo
    'platano', 'pina', 'papaya', 'sandia', 'melon', 'mango',
    'naranja', 'limon', 'limom', 'limo',           # limon typos
    'manzana', 'aguacate', 'jamaica', 'tamarindo',
    'guayaba', 'fresa', 'mora', 'arandano', 'orandano',
    # verduras / hortalizas
    'tomate', 'miltomate', 'cebolla', 'zanahoria', 'ejote',
    'guisquil', 'gusiquil', 'guisqul',             # guisquil typos
    'guicoy', 'ayote', 'calabaza', 'remolacha', 'repollo',
    'brocoli', 'brocoly',                          # brocoli typo
    'coliflor', 'papa', 'camote', 'yuca', 'malanga',
    'espinaca', 'bledo', 'rabano', 'lechuga', 'pepino',
    'chipolin', 'chipilin',
    # hierbas / aromaticas
    'perejil', 'ajo', 'apio', 'cilantro', 'oregano', 'romero',
    'hierba', 'hierba buena', 'hierbabuena', 'hirbabuena',
    'mashan', 'apazote', 'apasote',                # apazote misspelling
    'zacate', 'tusa', 'laurel', 'tomio', 'tomillo', 'albahaca',
    # granos frescos
    'maiz', 'cebada', 'cabada',                   # cebada typo
    'trigo', 'arveja', 'haba', 'azote',
    'ajonjoli', 'ajonjolin',                       # ajonjoli variant spelling
    # chiles cultivados (qualified only -- bare "chile" stays unmatched)
    'chile pimiento', 'chile pimento', 'chile pimienta',            # pimento typo
    'chile cobanero', 'chile verde', 'chile jalapeno', 'chile chiltepe',
    'chile dulce', 'chile morron',
    # frijol cultivado
    'frijol ejotero', 'frijol tierno', 'frijol negro', 'frijol vaina real',  'frijol seco',
    'frijol rojo', 'frijol colorado', 'frijol blanco', 'frijol en grano',
    # más
    'carne de res', 'res', 'pescado', 'huevo', 'pollo', 'pechuga', 'pierna', 'muslo',
    'tomillo', 'clavo', 'pimienta', 'comino', 'achiote', 'achote',          # achiote typo
    'canela', 'laurel', 'laure', 'pepitoria', 'pepitorio', 'mani', 'mania', 'manilla',
]

ABARROTES = [
    # semillas secas / procesadas
    'pepita', 'frijol sellado', 'paq de', 'paquete de', 'chiper', 'australian', 'anchor', 'australia', 'suli',
    # proteina animal
    'carne', 'embutido', 'chorizo', 'salchicha', 'jamon',
    # lacteos
    'crema', 'leche', 'queso', 'yogur', 'mantequilla', 'margarina',
    # panaderia
    'pan', 'pirujo',
    'tostada', 'tortilla', 'galleta', 'chocolate',
    # pasta / cereales procesados
    'pasta', 'espagueti', 'fideo', 'macarron', 'espaqueti',
    'codito',                                       # catches "pasta codito" / "pasto codito"
    'avena', 'abena',                               # avena typo
    'corazon de trigo',
    'chaomein', 'chow mein', 'chao mein', 'chaumein', 'cahomein',
    'mosh',
    # harinas / mezclas
    'maseca', 'incaparina', 'protemas', 'atol', 'harina', 'pinol',
    # aceites / condimentos
    'aceite', 'sal', 'azucar', 'vinagre',
    'pimiento en polvo',
    # otros
    'arroz', 'consome', 'concentrado', 'levadura', 'agua pura', 'bebida',
    # chiles procesados / secos
    'chile seco', 'chile rojo', 'chile en polvo', 'chile molido',
    'chile pasa', 'chila pasa',                    # chila typo
    'chile guaque', 'chile guaca',                 # guaca typo (muy comun)
    'chile chocolate', 'chile negro',
]


def fuzzy_match_category(description, cultivados=CULTIVADOS, abarrotes=ABARROTES, threshold=80):
    """
    Accent-insensitive, word-boundary matching with Spanish plural tolerance.
    Handles pdfplumber multi-line cells and concatenated units like 'espagueti180g'.
    Returns ('agricultura' | 'abarrotes' | 'unmatched', matched_keyword).
    """
    if not description:
        return ('unmatched', None)

    desc_norm = normalize_text(description)
    # Collapse newlines/tabs/repeated whitespace from multi-line PDF cells
    desc_norm = re.sub(r'\s+', ' ', desc_norm).strip()
    # Split letter/digit runs so "espagueti180g" -> "espagueti 180 g"
    desc_norm = re.sub(r'([a-z])(\d)', r'\1 \2', desc_norm)
    desc_norm = re.sub(r'(\d)([a-z])', r'\1 \2', desc_norm)

    # (e?s)? tolerates Spanish plurals: banano/bananos, limon/limones, etc.
    for kw in cultivados:
        if re.search(r'\b' + re.escape(kw) + r'(e?s)?\b', desc_norm):
            return ('agricultura', kw)
    for kw in abarrotes:
        if re.search(r'\b' + re.escape(kw) + r'(e?s)?\b', desc_norm):
            return ('abarrotes', kw)

    return ('unmatched', None)


def clasificar_factura(factura):
    """
    Adds 'categoria' and 'palabra_clave' to each line of a factura (from
    extraccion.procesar_pdf) and returns (agri_sum, abar_sum).
    Matching uses the full row text, as in the original tool.
    """
    agri_sum, abar_sum = 0, 0
    for linea in factura['lineas']:
        category, matched_word = fuzzy_match_category(linea['texto_fila'])
        linea['categoria'] = category
        linea['palabra_clave'] = matched_word
        if category == 'agricultura':
            agri_sum += linea['total']
        elif category == 'abarrotes':
            abar_sum += linea['total']
    factura['total_agri'] = agri_sum
    factura['total_abar'] = abar_sum
    return agri_sum, abar_sum
