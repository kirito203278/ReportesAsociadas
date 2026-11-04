"""
Reporte de Asociados — Procesador de PDF (Betterware)
======================================================
App local (Streamlit). Sube el PDF "Actividad de Mis Asociadas", extrae los datos
(código, teléfono, pedidos y TOTAL DE COMPRAS DEL CATÁLOGO ACTUAL) del texto directo,
y usa IA (Claude o Grok) SOLO para reconstruir los nombres correctos. Genera Excel
y PDF formales, ordenados por compras del catálogo actual (desc) y desempate por pedidos.

La API key NO se guarda en el código: se pega en la app o se lee de variable de entorno.
"""

import io
import os
import re
import json
import streamlit as st


# ============================================================================
# CARGA DE API KEYS DESDE ARCHIVO .env (si existe, junto a app.py)
# El .env NO es parte del código; lo creas tú y guardas ahí tus keys.
# Formato del .env:
#   ANTHROPIC_API_KEY=sk-ant-...
#   XAI_API_KEY=xai-...
# ============================================================================
def cargar_env():
    ruta = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env")
    if not os.path.exists(ruta):
        return
    try:
        with open(ruta, "r", encoding="utf-8") as f:
            for linea in f:
                linea = linea.strip()
                if not linea or linea.startswith("#") or "=" not in linea:
                    continue
                clave, valor = linea.split("=", 1)
                clave = clave.strip()
                valor = valor.strip().strip('"').strip("'")
                if clave and valor and clave not in os.environ:
                    os.environ[clave] = valor
    except Exception:
        pass


def cargar_streamlit_secrets():
    """Carga API keys desde Streamlit Secrets cuando la app está en la nube."""
    try:
        for clave in ("ANTHROPIC_API_KEY", "GEMINI_API_KEY", "XAI_API_KEY", "CLAVE_APP"):
            if clave not in os.environ:
                valor = st.secrets.get(clave, "")
                if valor:
                    os.environ[clave] = str(valor)
    except Exception:
        pass


cargar_env()
cargar_streamlit_secrets()

# ============================================================================
# UTILIDADES
# ============================================================================
def title_es(nombre):
    minus = {"de", "la", "las", "los", "del", "y", "da", "e"}
    palabras = nombre.strip().lower().split()
    out = []
    for i, p in enumerate(palabras):
        out.append(p if (p in minus and i != 0) else (p[:1].upper() + p[1:]))
    return " ".join(out)


def limpiar_dinero(txt):
    if txt is None:
        return 0
    s = re.sub(r"[^\d.]", "", str(txt))
    if s == "":
        return 0
    try:
        return int(round(float(s)))
    except ValueError:
        return 0


# ============================================================================
# EXTRACCIÓN DE DATOS (texto directo, sin OCR)
# Los datos numéricos salen perfectos. El nombre crudo es de respaldo por si no se usa IA.
#
# IMPORTANTE: cada asociada SIEMPRE tiene una línea con "Sem:NN / AAAA" seguida de
# importes semanales variables, total de pedidos y total de compras. Esa línea es
# el ANCLA fiable de cada fila. El teléfono PUEDE NO EXISTIR, y el código a veces
# queda en una línea aparte (cuando el nombre es largo y se parte en varias líneas).
# Por eso ya NO se exige teléfono ni que el código esté pegado al inicio de la línea:
# así no se salta ninguna asociada.
# ============================================================================

# Patrón base para encontrar la parte financiera de cada asociada.
#
# El PDF puede cambiar de catálogo y traer 4, 5 o más semanas visibles:
#   Ejemplo 1: S23 S24 S25 S26
#   Ejemplo 2: S27 S28 S29 S30 S31
#
# Por eso ya NO contamos semanas fijas. Primero encontramos "Sem:NN / AAAA" y luego
# leemos los últimos campos de la fila desde la derecha:
#   ... semanas ... TotalPedidosActual TotalComprasActual Referidas SaldoPuntos
#   TotalPedidosAnterior TotalComprasAnterior
#
# Esto evita jalar el Total Compras del catálogo anterior y soporta PDFs con diferente
# cantidad de columnas semanales.
SEMANA_FILA = re.compile(r"Sem:\s*\d+\s*/\s*\d{4}")
TOKEN_NUMERICO = re.compile(r"\$[\d,]+|\d[\d,]*")

# Línea que contiene ÚNICAMENTE un código (con posibles símbolos * + ~ - al inicio).
SOLO_CODIGO = re.compile(r"^[\*\+\~\-\s]*(\d{6,8})\s*$")


def _nombre_fallback(prefijo):
    """Arma un nombre de respaldo (sin IA) con el fragmento alfabético del prefijo."""
    frag = re.sub(r"\(cid:[^)]*\)", " ", prefijo)        # quita basura (cid:..)
    frag = re.sub(r"\d+", " ", frag)                      # quita números (código/teléfono)
    frag = re.sub(r"[\*\+\~\$/]", " ", frag)              # quita símbolos
    frag = re.sub(r"\bSem:\b", " ", frag)
    frag = re.sub(r"\b(PLUS|BASE|Inactivo)\b", " ", frag, flags=re.IGNORECASE)
    return re.sub(r"\s+", " ", frag).strip()




# ----------------------------
# NOMBRES POR COORDENADAS
# ----------------------------
# En estos PDFs, pdfplumber extrae los números bien por línea, pero los nombres
# pueden quedar arriba/abajo de la línea que contiene "Sem:NN / AAAA". Por eso
# también leemos la posición visual de las palabras dentro de la columna "Nombre".
CID_RE = re.compile(r"\(cid:(\d+)\)")

# Mapa para palabras que vienen dañadas por la fuente del PDF.
# Ejemplos: 5RGUtJXH] -> Rodríguez, 3(f$ -> Peña, *XWLpUUH] -> Gutiérrez.
CIFRA_NOMBRE = {
    "5": "R", "R": "o", "G": "d", "U": "r", "H": "e", "X": "u",
    "J": "g", "]": "z", "&": "C", "$": "A", "(": "E", "3": "P",
    "L": "i", "O": "l", "F": "c", "D": "a", "Q": "n", "V": "s",
    "-": "J", "W": "t", "p": "é", "i": "á", "t": "í", "y": "ó",
    "x": "ñ", ")": "F", "0": "M", "2": "O", "*": "G", "f": "Ñ",
    "1": "N", "8": "ú", "=": "z", "/": "L",
}

PALABRAS_NO_NOMBRE = {
    "base", "plus", "inactivo", "distribuido", "con", "r", "ctbt", "me",
    "asociada", "asociado", "nombre", "fecha", "ingreso", "total", "pedidos",
    "compras", "referidas", "saldo", "puntos", "nivel", "cierre",
}


def _cid_a_texto(txt):
    """Convierte piezas tipo (cid:42) a su carácter base."""
    return CID_RE.sub(lambda m: chr(int(m.group(1))), str(txt))


def _descifrar_si_aplica(txt_original):
    """Descifra SOLO palabras que vienen como cid o con símbolos de la fuente dañada."""
    txt = _cid_a_texto(txt_original)
    # No tocar palabras normales como Estrada, Garcia, Elizabeth, etc.
    # Solo descifrar si el token venía como cid o trae símbolos/dígitos del código dañado.
    if "(cid:" in str(txt_original) or re.search(r"[0-9\]\&\$\(\)\*\-]", txt):
        return "".join(CIFRA_NOMBRE.get(ch, ch) for ch in txt)
    return txt


def _limpiar_nombre_desde_partes(partes):
    """Limpia y une palabras de nombre detectadas visualmente."""
    palabras = []
    for parte in partes:
        txt = _descifrar_si_aplica(parte)
        # Mantener espacios cuando una palabra dañada contiene dos palabras: Molina Peña.
        txt = re.sub(r"[^A-Za-zÁÉÍÓÚÜáéíóúüÑñ]+", " ", txt)
        for p in txt.split():
            if not p:
                continue
            if p.lower() in PALABRAS_NO_NOMBRE:
                continue
            palabras.append(p)

    nombre = " ".join(palabras)
    nombre = re.sub(r"\bGuadarram\s+A\b", "Guadarrama", nombre, flags=re.IGNORECASE)
    nombre = re.sub(r"\bGuillermin\s+A\b", "Guillermina", nombre, flags=re.IGNORECASE)
    nombre = re.sub(r"\bVelzquez\b", "Velázquez", nombre, flags=re.IGNORECASE)
    nombre = re.sub(r"\bVernica\b", "Verónica", nombre, flags=re.IGNORECASE)
    nombre = re.sub(r"\bLucia\b", "Lucía", nombre, flags=re.IGNORECASE)
    return re.sub(r"\s+", " ", nombre).strip()


def _extraer_nombres_por_coordenadas(page):
    """Devuelve {codigo: nombre} leyendo la columna de nombres por posición visual.

    Esto evita que salgan nombres como '(revisar nombre)' cuando el nombre está
    partido en varias líneas alrededor de la fila financiera.
    """
    nombres = {}
    try:
        words = page.extract_words(
            x_tolerance=2,
            y_tolerance=3,
            keep_blank_chars=False,
            use_text_flow=False,
        )
    except Exception:
        return nombres

    # Cada fila real tiene una palabra tipo Sem:39, Sem:27, etc.
    sems = [w for w in words if _cid_a_texto(w.get("text", "")).startswith("Sem:")]
    sems.sort(key=lambda w: w["top"])
    if not sems:
        return nombres

    for i, sem in enumerate(sems):
        prev_top = sems[i - 1]["top"] if i > 0 else sem["top"] - 60
        next_top = sems[i + 1]["top"] if i + 1 < len(sems) else sem["top"] + 60
        top_b = (prev_top + sem["top"]) / 2
        bot_b = (sem["top"] + next_top) / 2

        fila_words = [w for w in words if top_b - 1 <= w["top"] < bot_b + 1]

        # Código: 6 a 8 dígitos, normalmente a la izquierda de la columna de nombre.
        codigos = []
        for w in fila_words:
            raw = _cid_a_texto(w.get("text", ""))
            dig = re.sub(r"\D", "", raw)
            if 6 <= len(dig) <= 8 and w.get("x0", 9999) < 200:
                codigos.append(dig)
        if not codigos:
            continue
        codigo = codigos[0]

        # Nombre: columna visual entre código y teléfono.
        # En este formato Betterware suele caer entre x≈198 y x≈276.
        partes_nombre = []
        for w in fila_words:
            x0 = w.get("x0", 0)
            raw = _cid_a_texto(w.get("text", ""))
            if 198 <= x0 <= 276:
                if raw.startswith("Sem:"):
                    continue
                if re.fullmatch(r"\$?[\d,]+", raw):
                    continue
                partes_nombre.append(w.get("text", ""))

        nombre = _limpiar_nombre_desde_partes(partes_nombre)
        if nombre:
            nombres[codigo] = title_es(nombre)

    return nombres


def extraer(pdf_bytes):
    """Devuelve (filas, texto_crudo). filas con datos correctos + nombre crudo de respaldo.

    Procesa página por página. Para cada línea con el ancla financiera crea una fila;
    el teléfono es opcional y, si el código no está en esa línea, lo recupera de una
    línea vecina que contenga solo el código.
    """
    import pdfplumber
    paginas = []
    nombres_por_codigo = {}
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            paginas.append((page.extract_text() or "").split("\n"))
            nombres_por_codigo.update(_extraer_nombres_por_coordenadas(page))
    texto_crudo = "\n".join("\n".join(p) for p in paginas)

    filas = []
    for lineas in paginas:
        codigos_usados = set()
        for idx, ln in enumerate(lineas):
            m = SEMANA_FILA.search(ln)
            if not m:
                continue

            # Tomamos la información DESPUÉS de "Sem:NN / AAAA".
            # Ejemplo con 5 semanas:
            # $2,800 $0 $0 $0 $0 1 $2,800 0 66,115 4 $9,453
            # Los últimos 6 tokens siempre son:
            # pedidos_actuales, compras_actuales, referidas, saldo_puntos,
            # pedidos_anteriores, compras_anteriores.
            resto = ln[m.end():]
            tokens = TOKEN_NUMERICO.findall(resto)

            if len(tokens) < 6:
                continue

            try:
                pedidos = int(tokens[-6].replace(",", ""))
                compras_actuales = limpiar_dinero(tokens[-5])
                compras_anteriores = limpiar_dinero(tokens[-1])
                pedidos_anteriores = int(tokens[-2].replace(",", ""))
            except Exception:
                continue

            prefijo = ln[:m.start()]
            # En el prefijo solo hay números de código (6-8 dígitos) y/o teléfono (>=9).
            nums = re.findall(r"\d{6,}", prefijo)
            codigo = next((n for n in nums if 6 <= len(n) <= 8), None)
            telefono = next((n for n in nums if len(n) >= 9), "")

            # Si el código no venía en la línea (nombre largo partido), buscarlo en una
            # línea cercana. Acepta tanto una línea con SOLO el código como una que
            # empiece con el código seguido de un fragmento del nombre
            # (p. ej. "14503899 Lopez Fuentes"). Nunca toma el código de otra fila de datos.
            if codigo is None:
                for off in (1, -1, 2, -2, 3, -3):
                    j = idx + off
                    if not (0 <= j < len(lineas)):
                        continue
                    cand = lineas[j]
                    if "Sem:" in cand:          # esa línea es otra fila de datos: no tomar su código
                        continue
                    mc = re.match(r"^[\*\+\~\-\s]*(\d{6,8})\b", cand)
                    if mc and mc.group(1) not in codigos_usados:
                        codigo = mc.group(1)
                        break
            if codigo:
                codigos_usados.add(codigo)

            nombre = ""
            if codigo:
                nombre = nombres_por_codigo.get(codigo, "")
            if not nombre:
                nombre = _nombre_fallback(prefijo)
                nombre = title_es(nombre) if nombre else "(revisar nombre)"

            filas.append({
                "codigo": codigo or "(revisar codigo)",
                "nombre": nombre,
                "telefono": telefono,
                "pedidos": pedidos,
                # Se conserva la llave "ingresos" para no romper el resto del código,
                # pero el valor real es Total Compras del Catálogo Actual.
                "ingresos": compras_actuales,
                "compras_catalogo_actual": compras_actuales,
                "pedidos_catalogo_actual": pedidos,
                "compras_catalogo_anterior": compras_anteriores,
                "pedidos_catalogo_anterior": pedidos_anteriores,
            })
    return filas, texto_crudo


def extraer_total_oficial_catalogo_actual(texto_crudo):
    """Intenta leer el total general oficial del PDF original.

    En los PDFs Betterware normalmente aparece una línea final así:
    Total $24,366 $0 $0 $0 $0 21 $24,366 0 81 $91,655

    El primer importe después de "Total" y el importe después de Total Pedidos
    corresponden a compras del Catálogo Actual. Si no se encuentra, devuelve None.
    """
    candidatos = []
    for ln in texto_crudo.splitlines():
        if not re.search(r"\bTotal\b", ln, flags=re.IGNORECASE):
            continue
        if "Sem:" in ln:
            continue
        tokens = TOKEN_NUMERICO.findall(ln)
        # Línea total con semanas + pedidos actual + compras actual + anterior.
        if len(tokens) >= 8:
            # Ejemplo con 5 semanas: $24,366 $0 $0 $0 $0 21 $24,366 0 81 $91,655
            # Los últimos 6 tokens son: pedidos_actual, compras_actual, referidas,
            # saldo_puntos si existe, pedidos_anterior, compras_anterior. En la línea
            # Total, puede faltar saldo_puntos, por eso preferimos el último importe
            # antes del bloque de catálogo anterior: tokens[-4] en este formato.
            try:
                compras_actual = limpiar_dinero(tokens[-4])
                pedidos_actual = int(str(tokens[-5]).replace(",", ""))
                candidatos.append((pedidos_actual, compras_actual, ln.strip()))
            except Exception:
                pass
    if not candidatos:
        return None
    # Tomamos el último candidato porque el total general suele estar en la última página.
    return candidatos[-1]


# ============================================================================
# LIMPIEZA DE NOMBRES CON IA (Claude o Grok)
# ============================================================================
def limpiar_nombres_con_ia(texto_crudo, filas, proveedor, api_key):
    codigos = [f["codigo"] for f in filas]
    prompt = (
        "Eres un experto extrayendo nombres de un reporte de Betterware ('Actividad de Mis Asociadas'). "
        "En el texto, cada asociada tiene un CÓDIGO numérico y, cerca de él, su nombre completo, que suele "
        "venir PARTIDO en varias líneas (por ejemplo 'MARIA', 'DEL', 'ROSARIO', 'CHAVEZ', 'AVILES' en "
        "líneas distintas). Une esos fragmentos para formar el nombre completo de cada código.\n\n"
        "INSTRUCCIONES:\n"
        "- Incluye SIEMPRE el nombre COMPLETO: nombre(s) y TODOS los apellidos. No omitas ningún apellido, "
        "aunque venga en una línea aparte o rodeado de números o símbolos.\n"
        "- Escribe los nombres con sus acentos y la Ñ correctos en español (María, Núñez, Peña, Fernández, "
        "Rodríguez, Gutiérrez, Garduño, Castañeda, etc.).\n"
        "- TEXTO CIFRADO: por la fuente del PDF, ALGUNOS nombres vienen 'revueltos' (mezcla de mayúsculas, "
        "dígitos y símbolos, p. ej. '5RGUtJXH]', '-XiUH]', '3LxD', '&DxHGR', '9HOi]TXH]'). Eso es un nombre. "
        "Descífralo con esta clave (carácter mostrado -> letra real): "
        "5->R R->o G->d U->r H->e X->u J->g ]->z &->C $->A (->E 3->P L->i O->l F->c D->a Q->n V->s -->J ; "
        "y para acentos/ñ: i->á t->í y->ó x->ñ f->Ñ . "
        "Ejemplos: '5RGUtJXH]'->Rodríguez, '-XiUH]'->Juárez, '3LxD'->Piña, '&DxHGR'->Cañedo, "
        "'9HOi]TXH]'->Velázquez. Aplica esto SOLO a las palabras revueltas de nombres; NUNCA a números, "
        "teléfonos ni a 'Sem:'.\n"
        "- Si algún carácter del texto está dañado o ilegible, dedúcelo según cómo se escribe normalmente "
        "ese nombre en español. No cambies un nombre por otro distinto (por ejemplo, NO conviertas Piña en "
        "Lira, ni Juárez en Núñez).\n"
        "- Formato de salida: cada nombre como 'Nombre Apellido Apellido' (primera letra de cada palabra en "
        "mayúscula; partículas como de, del, la, los en minúscula).\n\n"
        "Responde ÚNICAMENTE con un JSON válido (sin texto adicional, sin bloques de código): un objeto "
        "donde la llave es el código (string) y el valor es el nombre completo. "
        "Debes incluir los " + str(len(codigos)) + " códigos. Códigos: " + ", ".join(codigos) +
        "\n\nTEXTO CRUDO:\n" + texto_crudo
    )
    if proveedor == "Claude":
        contenido = _llamar_claude(prompt, api_key)
    elif proveedor == "Grok":
        contenido = _llamar_grok(prompt, api_key)
    else:
        contenido = _llamar_gemini(prompt, api_key)

    contenido = contenido.strip()
    contenido = re.sub(r"^```(?:json)?", "", contenido).strip()
    contenido = re.sub(r"```$", "", contenido).strip()
    ini, fin = contenido.find("{"), contenido.rfind("}")
    if ini != -1 and fin != -1:
        contenido = contenido[ini:fin + 1]

    # Intento 1: parsear el JSON completo
    mapa = {}
    try:
        mapa = json.loads(contenido)
    except json.JSONDecodeError:
        # Intento 2 (respaldo): si vino cortado, rescatar pares "codigo": "nombre" con regex
        for m in re.finditer(r'"(\d{6,8})"\s*:\s*"([^"]+)"', contenido):
            mapa[m.group(1)] = m.group(2)
        if not mapa:
            raise RuntimeError("La IA devolvió un formato que no se pudo leer.")

    for f in filas:
        val = mapa.get(f["codigo"])
        if val and str(val).strip():
            f["nombre"] = title_es(str(val).strip())
    return filas


def _llamar_claude(prompt, api_key):
    import urllib.request, urllib.error
    body = json.dumps({
        "model": "claude-haiku-4-5-20251001",
        "max_tokens": 4096,
        "messages": [{"role": "user", "content": prompt}],
    }).encode("utf-8")
    req = urllib.request.Request(
        "https://api.anthropic.com/v1/messages", data=body,
        headers={"content-type": "application/json", "x-api-key": api_key,
                 "anthropic-version": "2023-06-01"}, method="POST")
    try:
        with urllib.request.urlopen(req, timeout=120) as resp:
            data = json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        detalle = e.read().decode("utf-8", "ignore")
        raise RuntimeError(f"HTTP {e.code} de Anthropic: {detalle}")
    return "".join(b.get("text", "") for b in data.get("content", []) if b.get("type") == "text")


def _llamar_grok(prompt, api_key):
    import urllib.request, urllib.error
    body = json.dumps({
        "model": "grok-4.3",
        "messages": [{"role": "user", "content": prompt}],
        "max_completion_tokens": 4096,
    }).encode("utf-8")
    req = urllib.request.Request(
        "https://api.x.ai/v1/chat/completions", data=body,
        headers={"content-type": "application/json", "authorization": "Bearer " + api_key},
        method="POST")
    try:
        with urllib.request.urlopen(req, timeout=120) as resp:
            data = json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        detalle = e.read().decode("utf-8", "ignore")
        raise RuntimeError(f"HTTP {e.code} de xAI/Grok: {detalle}")
    return data["choices"][0]["message"]["content"]


def _llamar_gemini(prompt, api_key):
    import urllib.request, urllib.error, socket
    modelo = "gemini-2.5-flash"
    url = f"https://generativelanguage.googleapis.com/v1beta/models/{modelo}:generateContent"
    body = json.dumps({
        "contents": [{"parts": [{"text": prompt}]}],
        "generationConfig": {"temperature": 0, "maxOutputTokens": 65536,
                              "responseMimeType": "application/json",
                              "thinkingConfig": {"thinkingBudget": 0}},
    }).encode("utf-8")
    req = urllib.request.Request(
        url, data=body,
        headers={"content-type": "application/json", "x-goog-api-key": api_key},
        method="POST")
    try:
        with urllib.request.urlopen(req, timeout=300) as resp:
            data = json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        detalle = e.read().decode("utf-8", "ignore")
        raise RuntimeError(f"HTTP {e.code} de Gemini: {detalle}")
    except (socket.timeout, TimeoutError):
        raise RuntimeError("Gemini tardó demasiado (timeout). Intenta de nuevo.")
    # La respuesta viene en candidates[0].content.parts[*].text
    try:
        partes = data["candidates"][0]["content"]["parts"]
        return "".join(p.get("text", "") for p in partes)
    except (KeyError, IndexError):
        raise RuntimeError(f"Respuesta inesperada de Gemini: {json.dumps(data)[:500]}")


# ============================================================================
# EXCEL
# ============================================================================
def generar_excel(filas, titulo_reporte):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.formatting.rule import DataBarRule

    MORADO, MORADO_OSC, NARANJA, BLANCO, LILA = "702080", "4A1560", "F25C00", "FFFFFF", "F3ECF7"
    thin = Side(style="thin", color="D8CCE2")
    borde = Border(left=thin, right=thin, top=thin, bottom=thin)
    headers = ["#", "Nombre del Asociado", "Pedidos", "Compras Catálogo Actual"]

    def construir(ws, subtitulo, registros):
        """Llena una hoja con la tabla de marca (título, logo, encabezado, datos, total, barras)."""
        ws.sheet_view.showGridLines = False

        # Título
        ws.merge_cells("A1:D1")
        c = ws["A1"]; c.value = titulo_reporte
        c.font = Font(name="Arial", size=16, bold=True, color=BLANCO)
        c.fill = PatternFill("solid", fgColor=MORADO)
        c.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 34
        try:
            from openpyxl.drawing.image import Image as XLImage
            lp = _ruta_junto_app("logo.png")
            if os.path.exists(lp):
                img = XLImage(lp); img.height = 40; img.width = int(40 * 271 / 301)
                ws.add_image(img, "A1")
        except Exception:
            pass

        # Subtítulo naranja (marca + descripción de la hoja)
        ws.merge_cells("A2:D2")
        s = ws["A2"]; s.value = f"INNquietus · {subtitulo}"
        s.font = Font(name="Arial", size=11, bold=True, color=BLANCO)
        s.fill = PatternFill("solid", fgColor=NARANJA)
        s.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[2].height = 22

        # Encabezados
        for i, h in enumerate(headers, start=1):
            cell = ws.cell(row=3, column=i, value=h)
            cell.font = Font(name="Arial", size=11, bold=True, color=BLANCO)
            cell.fill = PatternFill("solid", fgColor=MORADO_OSC)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = borde
        ws.row_dimensions[3].height = 28

        # Datos (columna A = número de asociado)
        fila = 4
        for idx, r in enumerate(registros):
            relleno = LILA if idx % 2 else BLANCO
            vals = [idx + 1, r["nombre"], r["pedidos"], r["ingresos"]]
            for col, val in enumerate(vals, start=1):
                cell = ws.cell(row=fila, column=col, value=val)
                cell.font = Font(name="Arial", size=10)
                cell.fill = PatternFill("solid", fgColor=relleno)
                cell.border = borde
                if col == 1:      # #
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                elif col == 2:    # Nombre
                    cell.alignment = Alignment(horizontal="left", vertical="center")
                elif col == 3:    # Pedidos
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                else:             # Compras
                    cell.alignment = Alignment(horizontal="right", vertical="center"); cell.number_format = '"$"#,##0'
            fila += 1

        # Total (etiqueta en A-B, sumas en C=pedidos y D=compras)
        ws.merge_cells(start_row=fila, start_column=1, end_row=fila, end_column=2)
        ws.cell(row=fila, column=1, value="TOTAL")
        for col in (1, 2, 3, 4):
            cell = ws.cell(row=fila, column=col)
            cell.font = Font(name="Arial", size=11, bold=True, color=BLANCO)
            cell.fill = PatternFill("solid", fgColor=MORADO)
            cell.border = borde
        ws.cell(row=fila, column=1).alignment = Alignment(horizontal="center", vertical="center")
        if registros:
            ws.cell(row=fila, column=3, value=f"=SUM(C4:C{fila-1})").alignment = Alignment(horizontal="center", vertical="center")
            e = ws.cell(row=fila, column=4, value=f"=SUM(D4:D{fila-1})")
        else:
            ws.cell(row=fila, column=3, value=0).alignment = Alignment(horizontal="center", vertical="center")
            e = ws.cell(row=fila, column=4, value=0)
        e.alignment = Alignment(horizontal="right", vertical="center"); e.number_format = '"$"#,##0'

        # Barras de datos moradas en la columna de compras (D)
        if registros:
            regla = DataBarRule(start_type="min", end_type="max", color="B583C9", showValue=True)
            ws.conditional_formatting.add(f"D4:D{fila-1}", regla)

        for col, w in zip("ABCD", (6, 44, 14, 26)):
            ws.column_dimensions[col].width = w
        ws.freeze_panes = "A4"

    # Separar: activas = 1 o más pedidos; inactivas = 0 pedidos
    activos = [r for r in filas if r["pedidos"] >= 1]
    inactivos = [r for r in filas if r["pedidos"] == 0]

    wb = Workbook()
    ws1 = wb.active; ws1.title = "Reporte"
    construir(ws1, "Todos · Ordenado por compras del catálogo actual", filas)                     # tabla original completa
    construir(wb.create_sheet("Activos"), "Activos · 1 o más pedidos", activos)
    construir(wb.create_sheet("Inactivos"), "Inactivos · 0 pedidos", inactivos)

    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf


# ============================================================================
# PDF
# ============================================================================
def generar_pdf(filas, titulo_reporte):
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import letter
    from reportlab.lib.units import cm
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image as RLImage
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.enums import TA_CENTER

    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=letter, topMargin=1.3*cm, bottomMargin=1.5*cm,
                            leftMargin=1.2*cm, rightMargin=1.2*cm)
    estilos = getSampleStyleSheet()
    marca_style = ParagraphStyle("marca", parent=estilos["Normal"], fontSize=11, alignment=TA_CENTER,
                                 textColor=colors.HexColor("#F25C00"), spaceAfter=2, leading=13)
    titulo_style = ParagraphStyle("titulo", parent=estilos["Title"], textColor=colors.HexColor("#702080"),
                                  fontSize=18, alignment=TA_CENTER, spaceAfter=4)
    sub_style = ParagraphStyle("sub", parent=estilos["Normal"], fontSize=9, alignment=TA_CENTER,
                               textColor=colors.HexColor("#404040"), spaceAfter=12)

    elementos = []
    logo_path = _ruta_junto_app("logo.png")
    if os.path.exists(logo_path):
        try:
            logo = RLImage(logo_path, width=1.7*cm, height=1.7*cm*301/271)  # mantiene proporción
            logo.hAlign = "CENTER"
            elementos += [logo, Spacer(1, 4)]
        except Exception:
            pass
    elementos += [Paragraph("<b>INNquietus</b>", marca_style),
                  Paragraph(titulo_reporte, titulo_style),
                  Paragraph("Catálogo Actual · Ordenado por compras del catálogo actual", sub_style), Spacer(1, 6)]

    def banda(texto, color_hex):
        """Banner de color de ancho completo para encabezar cada sección."""
        b = Table([[texto]], colWidths=[17*cm])
        b.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor(color_hex)),
            ("TEXTCOLOR", (0, 0), (-1, -1), colors.white),
            ("FONTNAME", (0, 0), (-1, -1), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, -1), 12),
            ("LEFTPADDING", (0, 0), (-1, -1), 10),
            ("TOPPADDING", (0, 0), (-1, -1), 6),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ]))
        return b

    def construir_tabla(registros):
        header_style = ParagraphStyle(
            "header_tabla",
            parent=estilos["Normal"],
            fontName="Helvetica-Bold",
            fontSize=7,
            leading=8,
            textColor=colors.white,
            alignment=TA_CENTER,
        )
        data = [[
            Paragraph("#", header_style),
            Paragraph("Nombre del Asociado", header_style),
            Paragraph("Pedidos", header_style),
            Paragraph("Compras Catálogo Actual", header_style),
        ]]
        tot_p = tot_i = 0
        for i, r in enumerate(registros, start=1):
            data.append([str(i), r["nombre"], str(r["pedidos"]), f"${r['ingresos']:,}"])
            tot_p += r["pedidos"]; tot_i += r["ingresos"]
        data.append(["", "TOTAL", str(tot_p), f"${tot_i:,}"])
        t = Table(data, colWidths=[1.0*cm, 8.4*cm, 2.6*cm, 5.0*cm], repeatRows=1)
        t.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#4A1560")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, -1), 8),
            ("ALIGN", (0, 0), (0, -1), "CENTER"),
            ("ALIGN", (1, 0), (1, -1), "LEFT"),
            ("ALIGN", (2, 0), (2, -1), "CENTER"),
            ("ALIGN", (3, 0), (3, -1), "RIGHT"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("ROWBACKGROUNDS", (0, 1), (-1, -2), [colors.white, colors.HexColor("#F3ECF7")]),
            ("BACKGROUND", (0, -1), (-1, -1), colors.HexColor("#702080")),
            ("TEXTCOLOR", (0, -1), (-1, -1), colors.white),
            ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"),
            ("ALIGN", (1, -1), (1, -1), "CENTER"),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#D8CCE2")),
            ("TOPPADDING", (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ]))
        return t

    activos = [r for r in filas if r["pedidos"] >= 1]
    inactivos = [r for r in filas if r["pedidos"] == 0]

    # 1) Tabla completa (la original)
    elementos += [banda(f"TODOS LOS ASOCIADOS  ({len(filas)})", "#702080"), Spacer(1, 4),
                  construir_tabla(filas)]
    # 2) Activas
    elementos += [Spacer(1, 18), banda(f"ASOCIADOS ACTIVOS  ({len(activos)})", "#F25C00"),
                  Spacer(1, 4), construir_tabla(activos)]
    # 3) Inactivas
    elementos += [Spacer(1, 18), banda(f"ASOCIADOS INACTIVOS  ({len(inactivos)})", "#8A7E93"),
                  Spacer(1, 4), construir_tabla(inactivos)]

    doc.build(elementos)
    buf.seek(0)
    return buf


# ============================================================================
# GOOGLE SHEETS (gratis; requiere credentials.json de Google Cloud junto a app.py)
# Crea una hoja nueva en TU Google Drive con los datos y devuelve su URL.
# La primera vez abre el navegador para que autorices; luego guarda token.json
# y ya no vuelve a pedir permiso.
# ============================================================================
GOOGLE_SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.file",  # solo archivos creados por esta app
]


def _ruta_junto_app(nombre):
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), nombre)


def _leer_secret_json(nombre):
    """Lee un JSON guardado en Streamlit Secrets como texto multilínea."""
    try:
        valor = st.secrets.get(nombre, "")
        if valor:
            return json.loads(str(valor))
    except Exception:
        pass
    return None


def _credenciales_google():
    """Obtiene credenciales OAuth para Google Sheets.

    En local puede usar credentials.json/token.json.
    En Streamlit Cloud usa GOOGLE_TOKEN_JSON y GOOGLE_CREDENTIALS_JSON desde Secrets.
    """
    from google.oauth2.credentials import Credentials
    from google.auth.transport.requests import Request
    from google_auth_oauthlib.flow import InstalledAppFlow

    cred_path = _ruta_junto_app("credentials.json")
    token_path = _ruta_junto_app("token.json")

    creds = None

    # 1) NUBE: token guardado en Streamlit Secrets.
    token_info = _leer_secret_json("GOOGLE_TOKEN_JSON")
    if token_info:
        creds = Credentials.from_authorized_user_info(token_info, GOOGLE_SCOPES)

    # 2) LOCAL: token.json junto a app.py.
    elif os.path.exists(token_path):
        creds = Credentials.from_authorized_user_file(token_path, GOOGLE_SCOPES)

    # 3) Refrescar token si está vencido.
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())

    # 4) Si ya es válido, listo.
    if creds and creds.valid:
        return creds

    # 5) Si no hay token válido, solo local permite abrir navegador para autorizar.
    cred_info = _leer_secret_json("GOOGLE_CREDENTIALS_JSON")
    if cred_info:
        raise RuntimeError(
            "Google Sheets está configurado con credentials en Secrets, pero falta "
            "GOOGLE_TOKEN_JSON válido. Genera token.json en tu laptop y pégalo en Secrets."
        )

    if not os.path.exists(cred_path):
        raise RuntimeError(
            "Falta 'credentials.json' junto a app.py o GOOGLE_CREDENTIALS_JSON/GOOGLE_TOKEN_JSON "
            "en Streamlit Secrets."
        )

    flow = InstalledAppFlow.from_client_secrets_file(cred_path, GOOGLE_SCOPES)
    creds = flow.run_local_server(port=0)  # abre el navegador para autorizar en local
    with open(token_path, "w", encoding="utf-8") as f:
        f.write(creds.to_json())
    return creds


def subir_a_google_sheets(filas, titulo):
    """Crea una hoja de cálculo en el Drive del usuario con 3 pestañas y devuelve su URL."""
    import gspread
    creds = _credenciales_google()
    gc = gspread.authorize(creds)
    sh = gc.create(titulo.strip() or "Reporte de Asociados")

    MORADO = {"red": 112/255, "green": 32/255, "blue": 128/255}   # #702080
    LILA   = {"red": 243/255, "green": 236/255, "blue": 247/255}  # #F3ECF7
    BLANCO = {"red": 1, "green": 1, "blue": 1}

    def llenar(ws, registros):
        total_ped = sum(r["pedidos"] for r in registros)
        total_ing = sum(r["ingresos"] for r in registros)
        data = [["Nombre del Asociado", "Total de Pedidos", "Total Compras Catálogo Actual"]]
        for r in registros:
            data.append([r["nombre"], r["pedidos"], r["ingresos"]])
        data.append(["TOTAL", total_ped, total_ing])
        sh.values_update(f"'{ws.title}'!A1",
                         params={"valueInputOption": "USER_ENTERED"},
                         body={"values": data})
        n = len(data)
        try:  # el formato es opcional: si falla, los datos igual quedan bien
            enc = {"backgroundColor": MORADO,
                   "textFormat": {"bold": True, "foregroundColor": BLANCO},
                   "horizontalAlignment": "CENTER"}
            ws.format("A1:C1", enc)                    # encabezado morado
            ws.format(f"A{n}:C{n}", enc)               # fila TOTAL morada
            ws.format(f"C2:C{n}", {"numberFormat": {"type": "CURRENCY", "pattern": '"$"#,##0'}})
            ws.freeze(rows=1)
            reqs = [
                {"updateDimensionProperties": {
                    "range": {"sheetId": ws.id, "dimension": "COLUMNS", "startIndex": 0, "endIndex": 1},
                    "properties": {"pixelSize": 300}, "fields": "pixelSize"}},
                {"updateDimensionProperties": {
                    "range": {"sheetId": ws.id, "dimension": "COLUMNS", "startIndex": 1, "endIndex": 3},
                    "properties": {"pixelSize": 140}, "fields": "pixelSize"}},
            ]
            if n > 3:  # bandas alternadas (zebra) nativas en las filas de datos: 1 sola petición
                reqs.append({"addBanding": {"bandedRange": {
                    "range": {"sheetId": ws.id, "startRowIndex": 1, "endRowIndex": n - 1,
                              "startColumnIndex": 0, "endColumnIndex": 3},
                    "rowProperties": {"firstBandColor": BLANCO, "secondBandColor": LILA}}}})
            sh.batch_update({"requests": reqs})
        except Exception:
            pass

    activos = [r for r in filas if r["pedidos"] >= 1]
    inactivos = [r for r in filas if r["pedidos"] == 0]

    ws1 = sh.sheet1
    ws1.update_title("Reporte")
    llenar(ws1, filas)                                                    # tabla completa
    llenar(sh.add_worksheet(title="Activos", rows=max(len(activos) + 6, 12), cols=4), activos)
    llenar(sh.add_worksheet(title="Inactivos", rows=max(len(inactivos) + 6, 12), cols=4), inactivos)
    return sh.url


# ============================================================================
# ACCESO PRIVADO
# ============================================================================
def _obtener_clave_app():
    """Obtiene la contraseña de acceso desde Secrets o variable de entorno."""
    try:
        clave = st.secrets.get("CLAVE_APP", "")
        if clave:
            return str(clave)
    except Exception:
        pass
    return os.environ.get("CLAVE_APP", "")


def _validar_acceso():
    """Muestra pantalla de contraseña y devuelve True si el usuario ya entró."""
    import hmac

    clave_correcta = _obtener_clave_app().strip()

    # Si no configuras CLAVE_APP, la app queda abierta. Útil para pruebas locales.
    if not clave_correcta:
        return True

    if st.session_state.get("acceso_autorizado"):
        return True

    # Estilos solo para la pantalla de acceso.
    # También ocultamos el texto de Streamlit “Press Enter to apply”, que hacía
    # que el campo se viera amontonado junto al ícono de mostrar contraseña.
    # Usamos st.html (no st.markdown) para inyectar el CSS de forma estable
    # entre versiones de Streamlit.
    st.html("""
        <style>
        div[data-testid="InputInstructions"]{ display:none !important; }
        [data-testid="stSidebar"]{ display:none !important; }
        header[data-testid="stHeader"]{ background:transparent !important; }
        .stApp{ background:
            radial-gradient(620px 440px at 12% 6%, rgba(242,92,0,.22), transparent 60%),
            radial-gradient(720px 520px at 90% 16%, rgba(112,32,128,.26), transparent 60%),
            radial-gradient(680px 520px at 50% 104%, rgba(242,229,0,.18), transparent 60%),
            linear-gradient(160deg,#f6eefc 0%, #f2e9f8 100%) !important;
            background-attachment:fixed !important; }
        .block-container{ max-width:480px !important; padding-top:11vh !important; }

        /* La tarjeta de vidrio del login (solo la marcada con key) */
        [class*="st-key-glasscard"]{
            background:rgba(255,255,255,.55) !important; border:1px solid rgba(255,255,255,.65) !important;
            border-radius:24px !important; box-shadow:0 24px 60px rgba(112,32,128,.16) !important;
            padding:30px 30px 26px !important;
            backdrop-filter:blur(20px) saturate(150%) !important; -webkit-backdrop-filter:blur(20px) saturate(150%) !important;
        }
        /* El formulario no debe dibujar su propia caja */
        [data-testid="stForm"]{ border:none !important; background:transparent !important; padding:0 !important; box-shadow:none !important; }
        /* Enmascara el campo de acceso con puntitos SIN ser type=password
           (así ningún navegador ofrece guardar la contraseña) */
        [class*="st-key-clave_login"] input{
            -webkit-text-security:disc !important; text-security:disc !important;
        }
        /* espaciado solo a los puntitos, no al texto de ayuda */
        [class*="st-key-clave_login"] input:not(:placeholder-shown){ letter-spacing:.12em; }
        .login-top{ text-align:center; }
        .login-icon{
            width:70px;height:70px;margin:2px auto 16px;border-radius:22px;
            display:flex;align-items:center;justify-content:center;font-size:33px;
            background:linear-gradient(135deg,#702080 0%,#F25C00 100%);
            box-shadow:0 14px 30px rgba(112,32,128,.30);
        }
        .login-title{ margin:0;color:#4A1560;font-size:25px;font-weight:800;letter-spacing:-.01em; }
        .login-subtitle{ margin:9px 6px 2px;color:#7a6f84;font-size:14px;line-height:1.5; }
        .login-barra{ height:5px;border-radius:4px;margin:16px 0 6px;
            background:linear-gradient(90deg,#702080 0 33%,#F2E500 33% 66%,#F25C00 66% 100%); }

        /* Campo del login: se estiliza el contenedor real de Streamlit (Root), no el input */
        [data-testid="stTextInputRootElement"]{
            height:52px !important;border:1.5px solid rgba(112,32,128,.18) !important;border-radius:13px !important;
            background:rgba(255,255,255,.75) !important;
        }
        [data-testid="stTextInputRootElement"]:focus-within{ border-color:#702080 !important;
            box-shadow:0 0 0 3px rgba(112,32,128,.13) !important; }
        [data-testid="stTextInputRootElement"] [data-baseweb="base-input"],
        [data-testid="stTextInputRootElement"] input{
            background:transparent !important;border:none !important;box-shadow:none !important;
            height:100% !important;font-size:16px !important;color:#2b2330 !important;
        }
        /* Botón Entrar (cubre botón normal y de formulario) */
        .stButton>button, [data-testid="stFormSubmitButton"] button{
            background:linear-gradient(135deg,#702080 0%,#F25C00 120%) !important;color:#fff !important;
            font-weight:800 !important;border:none !important;border-radius:13px !important;height:50px !important;
            box-shadow:0 12px 24px rgba(112,32,128,.26) !important;
        }
        .stButton>button *, [data-testid="stFormSubmitButton"] button *{ color:#fff !important; }
        .stButton>button:hover, [data-testid="stFormSubmitButton"] button:hover{ filter:brightness(1.07); }
        </style>
    """)

    # Todo el login dentro de UNA sola tarjeta centrada.
    # Usamos st.form para que al presionar Enter en la contraseña intente entrar.
    with st.container(border=False, key="glasscard_login"):
        st.html("""
            <div class="login-top">
                <div class="login-icon">🔐</div>
                <div class="login-title">Acceso privado</div>
                <div class="login-subtitle">Ingresa la contraseña para usar el extractor de reportes Betterware.</div>
                <div class="login-barra"></div>
            </div>
        """)
        with st.form("login_form", border=False, clear_on_submit=False):
            # Campo de texto normal (NO type="password") enmascarado con CSS.
            # Así el navegador no lo reconoce como contraseña y no ofrece guardarla.
            clave_ingresada = st.text_input(
                "Contraseña de acceso",
                placeholder="Escribe la contraseña",
                label_visibility="collapsed",
                key="clave_login",
            )
            entrar = st.form_submit_button("Entrar", type="primary", use_container_width=True)

    if entrar:
        if hmac.compare_digest(clave_ingresada.strip(), clave_correcta):
            st.session_state["acceso_autorizado"] = True
            st.rerun()
        else:
            st.error("Contraseña incorrecta. Revisa e intenta de nuevo.")

    st.stop()


# ============================================================================
# INTERFAZ
# ============================================================================
def main():
    st.set_page_config(page_title="Reporte de Asociados", page_icon="📊", layout="centered")
    st.html("""
        <style>
        :root{
          --morado:#702080; --morado-osc:#4A1560; --naranja:#F25C00;
          --amarillo:#F2E500; --tinta:#2b2330; --borde:#ece3f2;
        }
        .stApp{ background:
            radial-gradient(620px 440px at 10% 8%, rgba(242,92,0,.20), transparent 60%),
            radial-gradient(720px 520px at 92% 14%, rgba(112,32,128,.24), transparent 60%),
            radial-gradient(680px 560px at 50% 104%, rgba(242,229,0,.16), transparent 60%),
            linear-gradient(160deg,#f6eefc 0%, #f2e9f8 100%) !important; }
        header[data-testid="stHeader"]{ background:transparent !important; }
        .block-container{ max-width:760px !important; padding-top:2.2rem !important; padding-bottom:3rem !important; }
        h1,h2,h3,h4,h5,h6{ color:var(--morado-osc) !important; letter-spacing:-.01em; }
        .stCaption, .stMarkdown p{ color:var(--tinta) !important; }
        div[data-testid="InputInstructions"]{ display:none !important; }

        /* ===== Encabezado (hero) ===== */
        .head{ display:flex; align-items:center; gap:16px; }
        .logo-tile{ width:58px;height:58px;border-radius:15px;display:flex;align-items:center;justify-content:center;
          font-size:27px;background:linear-gradient(135deg,var(--morado),var(--naranja));
          box-shadow:0 10px 22px rgba(112,32,128,.24); overflow:hidden; }
        .logo-tile img{ width:100%;height:100%;object-fit:cover; }
        .kicker{ font-size:13px;font-weight:800;letter-spacing:3px;text-transform:uppercase;color:var(--naranja); }
        .kicker b{ color:var(--morado); }
        .title{ font-size:38px;font-weight:800;color:var(--morado-osc);margin:2px 0 0;line-height:1.05; }
        .barL{ height:6px;width:120px;border-radius:4px;margin:12px 0 2px;
          background:linear-gradient(90deg,var(--morado) 0 33%,var(--amarillo) 33% 66%,var(--naranja) 66% 100%); }
        .sec{ font-size:13px;font-weight:800;letter-spacing:2px;text-transform:uppercase;color:#9a86a8;margin:24px 0 8px; }

        /* ===== Inputs ===== */
        /* Streamlit pone borde y altura en el contenedor (stTextInputRootElement),
           NO en el input. Se estiliza el contenedor y el input queda transparente.
           OJO: nunca escribir etiquetas con signos de menor/mayor dentro de este
           bloque de estilos: el filtro DOMPurify de st.html borraria todo el bloque. */
        [data-testid="stTextInputRootElement"]{ height:46px !important;
          background:rgba(255,255,255,.75) !important;
          border:1.5px solid rgba(112,32,128,.16) !important;border-radius:12px !important; }
        [data-testid="stTextInputRootElement"]:focus-within{ border-color:var(--morado) !important;
          box-shadow:0 0 0 3px rgba(112,32,128,.12) !important; }
        [data-testid="stTextInputRootElement"] [data-baseweb="base-input"],
        [data-testid="stTextInputRootElement"] input{ background:transparent !important;border:none !important;
          box-shadow:none !important;height:100% !important;color:var(--tinta) !important; }

        /* ===== Zona de carga (vidrio) ===== */
        [data-testid="stFileUploader"]{
          background:rgba(255,255,255,.5) !important; border:1.5px dashed rgba(112,32,128,.38) !important; border-radius:16px;
          box-shadow:0 12px 34px rgba(112,32,128,.10);
          backdrop-filter:blur(16px) saturate(150%) !important; -webkit-backdrop-filter:blur(16px) saturate(150%) !important;
        }
        [data-testid="stFileUploaderDropzone"], [data-testid="stFileUploaderDropzone"] section{
          background:transparent !important; color:var(--tinta) !important; padding:10px 6px !important;
        }
        [data-testid="stFileUploaderDropzone"] button{
          background:var(--morado) !important; color:#fff !important; border:none !important; border-radius:10px !important;
        }
        [data-testid="stFileUploaderDropzone"] button *{ color:#fff !important; }

        /* Traducción de los textos internos del uploader (Streamlit los pone en inglés) */
        [data-testid="stFileUploaderDropzoneInstructions"] span{ display:none !important; }
        [data-testid="stFileUploaderDropzoneInstructions"] small{ display:none !important; }
        [data-testid="stFileUploaderDropzoneInstructions"] > div::before{
          content:"Arrastra o selecciona tu PDF"; display:block; font-weight:700; color:var(--tinta); font-size:15px;
        }
        [data-testid="stFileUploaderDropzoneInstructions"] > div::after{
          content:"Límite 20 MB · PDF"; display:block; color:#8a7e93; font-size:13px; margin-top:2px;
        }
        [data-testid="stFileUploaderDropzone"] button{ font-size:0 !important; }
        [data-testid="stFileUploaderDropzone"] button::after{ content:"Seleccionar"; font-size:14px !important; font-weight:700; }

        /* ===== Botón primario "Generar" ===== */
        .stButton>button[kind="primary"], [data-testid="stBaseButton-primary"]{
          background:linear-gradient(135deg,var(--morado) 0%, var(--naranja) 120%) !important;
          color:#fff !important;font-weight:800 !important;font-size:17px !important;
          border:none !important;border-radius:14px !important;padding:14px 22px !important;
          box-shadow:0 14px 28px rgba(112,32,128,.26) !important; transition:filter .15s, transform .05s;
        }
        .stButton>button[kind="primary"]:hover, [data-testid="stBaseButton-primary"]:hover{ filter:brightness(1.07); }
        .stButton>button[kind="primary"]:active{ transform:translateY(1px); }
        .stButton>button[kind="primary"]:disabled, [data-testid="stBaseButton-primary"]:disabled{
          background:#d9c7e2 !important;color:#fdf6ff !important;box-shadow:none !important;cursor:not-allowed !important;filter:none !important;
        }

        /* ===== Botones secundarios (Sheets, cerrar sesión) y descargas (vidrio) ===== */
        [data-testid="stBaseButton-secondary"],
        [data-testid="stDownloadButton"] button{
          background:rgba(255,255,255,.82) !important;color:var(--morado) !important;border:1.5px solid rgba(112,32,128,.55) !important;
          border-radius:12px !important;font-weight:700 !important;
          box-shadow:0 6px 18px rgba(112,32,128,.10) !important;
          backdrop-filter:blur(10px) saturate(150%) !important; -webkit-backdrop-filter:blur(10px) saturate(150%) !important;
        }
        [data-testid="stBaseButton-secondary"] *, [data-testid="stDownloadButton"] button *{ color:var(--morado) !important; }
        [data-testid="stDownloadButton"] button:hover, [data-testid="stBaseButton-secondary"]:hover{ background:#fff !important; }

        /* ===== Advertencias / avisos (vidrio, conservando su color) ===== */
        [data-testid="stAlert"]{
          border-radius:14px !important;
          box-shadow:0 10px 28px rgba(112,32,128,.10) !important;
          backdrop-filter:blur(8px) saturate(140%) !important; -webkit-backdrop-filter:blur(8px) saturate(140%) !important;
          overflow:hidden !important;
        }

        /* ===== Quitar cualquier marco/caja interna: el fondo cubre todo ===== */
        body, [data-testid="stAppViewContainer"], [data-testid="stMain"],
        [data-testid="stMainBlockContainer"], section.main, .main{
          background:transparent !important; border:none !important; box-shadow:none !important;
        }

        /* ===== Tarjetas de métricas (vidrio) ===== */
        [data-testid="stMetric"]{
          background:rgba(255,255,255,.55) !important;border:1px solid rgba(255,255,255,.65) !important;border-radius:16px !important;
          padding:14px 18px !important;box-shadow:0 10px 30px rgba(112,32,128,.10) !important;
          backdrop-filter:blur(16px) saturate(150%) !important; -webkit-backdrop-filter:blur(16px) saturate(150%) !important;
        }
        [data-testid="stMetricLabel"] *{ color:#6a5f74 !important;font-weight:700 !important; }
        [data-testid="stMetricValue"]{ color:var(--morado) !important;font-weight:800 !important; }

        /* ===== Tarjetas de vidrio: SOLO las marcadas con key (no columnas ni el main) ===== */
        [class*="st-key-glasscard"]{
          background:rgba(255,255,255,.55) !important;border:1px solid rgba(255,255,255,.65) !important;border-radius:20px !important;
          box-shadow:0 18px 50px rgba(112,32,128,.12) !important;padding:12px 20px 16px !important;
          backdrop-filter:blur(18px) saturate(150%) !important; -webkit-backdrop-filter:blur(18px) saturate(150%) !important;
        }
        [data-testid="stDataFrame"]{ border-radius:12px;overflow:hidden;border:1px solid rgba(112,32,128,.12); }

        /* ===== Selector de IA tipo segmentos (pills) ===== */
        [data-testid="stSegmentedControl"] [role="radiogroup"], [data-baseweb="button-group"]{
          background:rgba(112,32,128,.08) !important;border-radius:12px !important;padding:4px !important;gap:4px !important;
        }
        [data-testid="stSegmentedControl"] button{
          border:none !important;border-radius:9px !important;background:transparent !important;
          color:#6d6175 !important;font-weight:700 !important;
        }
        [data-testid="stSegmentedControl"] button[aria-checked="true"], [data-testid="stSegmentedControl"] button[kind="primary"]{
          background:#fff !important;color:var(--morado) !important;box-shadow:0 4px 12px rgba(112,32,128,.18) !important;
        }
        [data-testid="stSegmentedControl"] button *{ color:inherit !important; }
        </style>
    """)

    _validar_acceso()

    # Logo embebido (base64) para el encabezado.
    import base64
    logo_path = _ruta_junto_app("logo.png")
    logo_inner = "📊"
    if os.path.exists(logo_path):
        with open(logo_path, "rb") as _f:
            _b64 = base64.b64encode(_f.read()).decode()
        logo_inner = f"<img src='data:image/png;base64,{_b64}'>"

    # ---- Encabezado ----
    st.html(f"""
        <div class="head">
          <div class="logo-tile">{logo_inner}</div>
          <div>
            <div class="kicker"><b>INN</b>quietus · Betterware</div>
            <div class="title">Reporte de Asociados</div>
          </div>
        </div>
        <div class="barL"></div>
    """)
    st.caption("Sube el PDF de actividad, genera el Excel/PDF y crea la hoja de Google. Procesa en un clic.")

    # ---------- Configuración ----------
    st.html('<div class="sec">Configuración</div>')
    with st.container(border=False, key="glasscard_config"):
        titulo_reporte = st.text_input("Título del reporte", value="Reporte de Asociados")
        usar_ia = st.checkbox("🤖 Corregir nombres con IA (recomendado)", value=True)
        proveedor, api_key = "Claude", ""
        if usar_ia:
            st.markdown("**Proveedor de IA**")
            proveedor = st.segmented_control(
                "Proveedor de IA", ["Claude", "Gemini"],
                default="Claude", selection_mode="single", label_visibility="collapsed")
            if not proveedor:
                proveedor = "Claude"
            if proveedor == "Claude":
                env_var = "ANTHROPIC_API_KEY"
            elif proveedor == "Grok":
                env_var = "XAI_API_KEY"
            else:
                env_var = "GEMINI_API_KEY"
            env_key = os.environ.get(env_var, "")
            if env_key:
                api_key = env_key
                st.caption(f"🔒 API Key de {proveedor} cargada del .env (oculta por seguridad).")
            else:
                api_key = st.text_input(
                    f"API Key de {proveedor}", value="", type="password",
                    help="Se usa solo en esta sesión. No se guarda en el código.")

    # ---------- Subir PDF ----------
    st.html('<div class="sec">Sube tu PDF</div>')
    archivo = st.file_uploader("Arrastra o selecciona el PDF", type=["pdf"],
                               label_visibility="collapsed")

    # ---------- Generar ----------
    st.session_state.setdefault("procesando", False)
    procesando = st.session_state["procesando"]
    etiqueta_btn = "⏳ Generando…" if procesando else "✨ Generar"
    generar = st.button(
        etiqueta_btn, type="primary", use_container_width=True,
        disabled=procesando or archivo is None,
    )
    if archivo is None and not procesando:
        st.caption("Sube un PDF para habilitar el botón **Generar**.")

    # Al pulsar: marcamos "procesando" y recargamos, para que el botón quede
    # bloqueado y muestre "Generando…" mientras se hace el trabajo (evita doble clic).
    if generar and archivo is not None:
        st.session_state["procesando"] = True
        st.rerun()

    # ========================================================================
    # PROCESO (esto SÍ consume créditos de IA): corre SOLO cuando está marcado
    # "procesando". El resultado se guarda en st.session_state para no repetirlo
    # en cada recarga; tocar otros controles o descargar NO vuelve a llamar a la IA.
    # ========================================================================
    if st.session_state["procesando"]:
        st.session_state.pop("error_proceso", None)
        if archivo is None:
            st.session_state["procesando"] = False
            st.rerun()
        try:
            with st.spinner("Leyendo el PDF…"):
                pdf_bytes = archivo.getvalue()
                filas, texto_crudo = extraer(pdf_bytes)
            if not filas:
                st.session_state["error_proceso"] = (
                    "No detecté filas de asociados. ¿El PDF tiene el formato esperado?")
            else:
                aviso_ia = None  # (tipo, mensaje[, detalle]) para mostrarlo desde la sesión
                if usar_ia and api_key.strip():
                    with st.spinner(f"Corrigiendo nombres con {proveedor}…"):
                        try:
                            filas = limpiar_nombres_con_ia(texto_crudo, filas, proveedor, api_key.strip())
                            aviso_ia = ("success", "Nombres corregidos con IA ✅")
                        except Exception as e:
                            import traceback
                            aviso_ia = ("error", f"La IA falló ({e}). Muestro los nombres sin corregir.",
                                        traceback.format_exc())
                elif usar_ia:
                    aviso_ia = ("warning", "Marcaste usar IA pero falta la API Key. Procesé sin IA.")

                filas.sort(key=lambda r: (-r["ingresos"], -r["pedidos"]))

                total_extraido = sum(r["ingresos"] for r in filas)
                pedidos_extraidos = sum(r["pedidos"] for r in filas)
                total_oficial = extraer_total_oficial_catalogo_actual(texto_crudo)
                aviso_total = None
                if total_oficial:
                    pedidos_oficiales, compras_oficiales, linea_total = total_oficial
                    if compras_oficiales != total_extraido or pedidos_oficiales != pedidos_extraidos:
                        aviso_total = (
                            pedidos_oficiales, compras_oficiales,
                            pedidos_extraidos, total_extraido, linea_total,
                        )

                # Se genera el Excel y el PDF una sola vez y se guardan en la sesión.
                st.session_state["resultado"] = {
                    "filas": filas,
                    "titulo": titulo_reporte,
                    "excel": generar_excel(filas, titulo_reporte).getvalue(),
                    "pdf": generar_pdf(filas, titulo_reporte).getvalue(),
                    "aviso_ia": aviso_ia,
                    "aviso_total": aviso_total,
                }
                st.session_state.pop("sheets_url", None)  # el link de Sheets es de otro reporte
        except Exception as e:
            st.session_state["error_proceso"] = f"Error al leer el PDF: {e}"
        finally:
            st.session_state["procesando"] = False
        st.rerun()

    # Error del último intento (si lo hubo)
    if st.session_state.get("error_proceso"):
        st.error(st.session_state["error_proceso"])

    # ========================================================================
    # RENDER (gratis): se dibuja en CADA recarga leyendo de la sesión.
    # No reprocesa nada; por eso el resultado permanece en pantalla aunque
    # toques la interfaz, y solo cambia al procesar otro PDF.
    # ========================================================================
    res = st.session_state.get("resultado")
    if res:
        aviso = res.get("aviso_ia")
        if aviso:
            if aviso[0] == "success":
                st.success(aviso[1])
            elif aviso[0] == "warning":
                st.warning(aviso[1])
            elif aviso[0] == "error":
                st.error(aviso[1])
                if len(aviso) > 2:
                    with st.expander("Ver detalle técnico del error"):
                        st.code(aviso[2])

        filas = res["filas"]
        aviso_total = res.get("aviso_total")
        if aviso_total:
            pedidos_oficiales, compras_oficiales, pedidos_extraidos, total_extraido, linea_total = aviso_total
            diferencia = total_extraido - compras_oficiales
            signo = "+" if diferencia > 0 else ""
            st.warning(
                f"⚠️ Diferencia de total: PDF ${compras_oficiales:,} vs suma ${total_extraido:,} "
                f"({signo}${diferencia:,}). Se conserva la suma por asociada."
            )

        # --- Resumen en tarjetas (KPIs) ---
        _tot_pedidos = sum(r["pedidos"] for r in filas)
        _tot_compras = sum(r["ingresos"] for r in filas)
        st.html('<div class="sec">Resumen</div>')
        m1, m2, m3 = st.columns(3)
        m1.metric("Asociados", f"{len(filas)}")
        m2.metric("Total de pedidos", f"{_tot_pedidos:,}")
        m3.metric("Compras del catálogo", f"${_tot_compras:,}")

        # --- Tabla + descargas ---
        st.html('<div class="sec">Detalle y descargas</div>')
        with st.container(border=False, key="glasscard_tabla"):
            st.dataframe(
                [{"#": i, "Nombre del Asociado": r["nombre"], "Total de Pedidos": r["pedidos"],
                  "Total Compras Catálogo Actual": f"${r['ingresos']:,}"} for i, r in enumerate(filas, 1)],
                use_container_width=True, hide_index=True)

            col1, col2 = st.columns(2)
            with col1:
                st.download_button("⬇️ Descargar Excel", res["excel"], file_name="Reporte_Asociados.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True)
            with col2:
                st.download_button("⬇️ Descargar PDF", res["pdf"], file_name="Reporte_Asociados.pdf",
                    mime="application/pdf", use_container_width=True)

        # --- Abrir en Google Sheets (usa el resultado YA procesado; NO llama a la IA) ---
        st.html('<div class="sec">Google Sheets</div>')
        if st.button("📗 Crear y abrir en Google Sheets", use_container_width=True):
            with st.spinner("Creando la hoja en tu Google Drive…"):
                try:
                    url = subir_a_google_sheets(res["filas"], res["titulo"])
                    st.session_state["sheets_url"] = url
                    try:
                        import webbrowser
                        webbrowser.open_new_tab(url)  # abre en el navegador de tu equipo
                    except Exception:
                        pass
                    st.success("Hoja creada ✅")
                except ModuleNotFoundError:
                    st.error(
                        "Google Sheets no está disponible porque faltan dependencias. "
                        "Revisa que requirements.txt esté actualizado y reinicia/despliega de nuevo la app."
                    )
                except Exception as e:
                    st.error(f"No se pudo crear la hoja: {e}")
                    import traceback
                    with st.expander("Ver detalle técnico"):
                        st.code(traceback.format_exc())
        if st.session_state.get("sheets_url"):
            st.markdown(f"🔗 [Abrir la hoja en Google Sheets]({st.session_state['sheets_url']})")

    # ---- Cerrar sesión (abajo, centrado) ----
    st.html("<div style='height:20px'></div>")
    lc, cc, rc = st.columns([1, 1.4, 1])
    with cc:
        if st.button("↩ Cerrar sesión", use_container_width=True, key="logout_btn"):
            st.session_state.pop("acceso_autorizado", None)
            st.rerun()

    # Pie de página de marca
    st.markdown(
        "<hr style='border:none;border-top:1px solid #e7dcef;margin:18px 0 10px;'>"
        "<div style='text-align:center;color:#9b8aa6;font-size:12px;'>"
        "<b style='color:#702080;'>INNquietus</b> · Reporte de Asociados</div>",
        unsafe_allow_html=True)


if __name__ == "__main__":
    main()