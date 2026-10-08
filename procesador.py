"""
procesador.py
Motor de procesamiento de informes de tasacion.

- Usa la API de Claude (Anthropic) para ENTENDER el peritaje en lenguaje
  libre y extraer los datos estructurados, sin importar como esten escritos.
- Completa la plantilla Excel virgen preservando logos, estilos y formulas.

Requiere la variable de entorno ANTHROPIC_API_KEY (configurada en Render).
"""

import os
import re
import json
import zipfile
import unicodedata
from io import BytesIO

from openpyxl import load_workbook
import anthropic


# Modelo de Claude a usar (Sonnet 4.6: buen equilibrio costo/calidad)
MODELO_CLAUDE = "claude-sonnet-4-6"


# ============================================================
#  UTILIDADES
# ============================================================

def normalizar(texto):
    """Pasa a mayusculas y quita tildes."""
    if texto is None:
        return ""
    texto = str(texto).upper()
    texto = unicodedata.normalize("NFD", texto)
    texto = "".join(c for c in texto if unicodedata.category(c) != "Mn")
    return texto.strip()


def a_numero(valor):
    """Convierte un string a numero ENTERO, quitando $, puntos y comas.
    Pensado para MONTOS en pesos (no tiene decimales).
    """
    if valor is None:
        return 0
    if isinstance(valor, (int, float)):
        return int(valor)
    s = re.sub(r"[^\d]", "", str(valor))
    return int(s) if s else 0


def a_numero_decimal(valor):
    """
    Como a_numero pero preservando decimales. Pensado para CANTIDADES
    de mano de obra (ej: '1,5' horas, '0.75 panos').

    Interpreta:
      - 1.5 o '1.5' o '1,5'  -> 1.5
      - '1.500'              -> 1500 (miles, no decimal)
      - '1.500,50'           -> 1500.5
      - 15                   -> 15
    """
    if valor is None or valor == "":
        return 0
    if isinstance(valor, (int, float)):
        return valor
    s = str(valor).strip().replace("$", "").replace(" ", "")
    if not s:
        return 0

    try:
        if "," in s:
            # Formato AR: coma es decimal, punto es separador de miles.
            s = s.replace(".", "").replace(",", ".")
            return float(s)
        if s.count(".") == 1:
            # Un solo punto: 1-2 digitos despues = decimal, 3 = miles.
            partes = s.split(".")
            if len(partes[1]) <= 2:
                return float(s)
            # Miles: "1.500" -> 1500
            return int(partes[0] + partes[1])
        if s.count(".") >= 2:
            # Miles: "1.500.000" -> 1500000
            return int(s.replace(".", ""))
        # Solo digitos
        n = re.sub(r"[^\d]", "", s)
        return int(n) if n else 0
    except Exception:
        return 0


# ============================================================
#  ESTRUCTURA DE DATOS
# ============================================================

def data_vacia():
    return {
        "numeroSiniestro": "", "fechaSiniestro": "", "fechaInspeccion": "",
        "asegurado": "", "marca": "", "modelo": "", "anio": "", "dominio": "",
        "chasis": "", "kilometraje": "", "sumaAsegurada": "", "franquiciaVeh": "",
        "tallerNombre": "", "tallerDireccion": "", "tallerLocalidad": "",
        "tallerTelefono": "", "tallerEmail": "",
        "danos": [],  # lista de {accion, pieza, precio}
        "manoObra": {
            "pintura": 0, "chapa": 0, "mecanica": 0, "tapiceria": 0, "varios": 0,
            "pinturaValor": 0, "chapaValor": 0,
            "mecanicaValor": 0, "tapiceriaValor": 0,
        },
        "franquicia": 0,
        "observaciones": "",
    }


# ============================================================
#  CEREBRO: ENTENDER EL PERITAJE CON LA API DE CLAUDE
# ============================================================

# Instrucciones para Claude: como perito, que datos extraer y en que formato.
_PROMPT_SISTEMA_PARSEO = """Sos un perito tasador de seguros automotores con \
experiencia. Recibis el texto libre de una peritacion (puede venir \
desordenado, con abreviaturas, datos de poliza mezclados, etc.) y tu tarea \
es EXTRAER los datos y devolverlos en formato JSON estricto.

Reglas de interpretacion:
- Las piezas bajo "SUSTITUIR" o "CAMBIAR" tienen accion "CAMBIAR".
- Las piezas bajo "PINTAR" o "REPARAR" tienen accion "REPARAR".

- PIEZAS DEL PERITAJE (pre-extraidas): si el texto incluye un bloque
  "=== PIEZAS DEL PERITAJE (extraidas del Excel) ===", ese bloque es la
  lista OFICIAL de piezas del peritaje, ya extraida por otro sistema.
  Cada linea tiene el formato "  N. [ACCION] Nombre de pieza".
  * En el campo "danos" del JSON, USA EXACTAMENTE esta lista, en el mismo
    orden, con los mismos nombres y las mismas acciones. NO agregues
    piezas nuevas, NO cambies los nombres, NO cambies las acciones.
  * Lo unico que puede variar es el precio: si hay una cotizacion pegada
    ("=== COTIZACION DE REPUESTOS ==="), asignale a cada pieza a CAMBIAR
    el precio que le corresponda de la cotizacion (matcheando por nombre
    con tolerancia a abreviaciones y variantes).
  * Si una pieza no aparece en la cotizacion, su precio va en 0.
  * Las piezas a REPARAR siempre llevan precio 0.

- PERITAJE DESDE EXCEL (GRILLA): si el texto incluye un bloque que
  empieza con "=== PERITAJE DESDE EXCEL (GRILLA) ===", ese bloque es el
  volcado de una planilla Excel de peritacion, celda por celda, con sus
  coordenadas (ej: "A55=Paragolpe | B55=X | D55=95000"). Interpretala
  asi:
  * La planilla tiene una grilla de danos dividida en SECTORES. Cada
    sector tiene un encabezado de texto (ej: "PARTE DELANTERA",
    "PARTE TRASERA", "LADO IZQUIERDO", "LADO DERECHO", "PARTE INTERIOR",
    "MOTOR", "CHASIS", "TREN TRASERO", "TREN DELANTERO", "OTROS").
  * Cada sector tiene un par de columnas marcadas "A" y "B" en su fila
    de encabezado, y una columna "Precio". Las piezas son los textos
    que estan a la izquierda de esas columnas.
  * Una "X" en la columna "A" del sector significa que esa pieza va a
    CAMBIAR. Una "X" en la columna "B" significa REPARAR.
  * IMPORTANTE: la "X" solo cuenta como marca cuando es el UNICO
    contenido de la celda (o "X" con un espacio). Si la "X" aparece
    DENTRO del texto de otra celda (ej: el nombre "Soportes x4" o
    "Tornillos 6x40" o cualquier palabra que contenga la letra x), eso
    NO es una marca. Tampoco lo es una celda que tenga un numero o
    cualquier otro texto. SOLO una celda cuyo unico contenido sea la
    letra "X" (mayuscula o minuscula) cuenta como marca.
  * Si una pieza no tiene "X" en NINGUNA de sus dos columnas de marca,
    NO la incluyas en "danos". Solo se cargan las piezas marcadas.
  * Recorré TODA la grilla y detectá TODAS las "X", en todos los
    sectores (las columnas A/B de cada sector estan en distintas
    posiciones: pueden ser B/C, G/H, L/M, etc.).
  * El nombre de cada pieza se arma combinando el SECTOR donde esta con
    el texto de la fila, para que quede sin ambiguedad. Ejemplos:
    "Paragolpe" en el sector "PARTE TRASERA" -> "Paragolpe trasero".
    "Guardabarro der." en "PARTE TRASERA" -> "Guardabarro trasero
    derecho". "Guardabarro der." en "PARTE DELANTERA" -> "Guardabarro
    delantero derecho".
  * IGNORÁ los precios que vengan en la columna "Precio" del Excel. Los
    precios NO se toman del Excel del peritaje.
  * Del Excel tambien extraé los datos del vehiculo, asegurado, taller,
    mano de obra y observaciones si estan presentes.

- PRECIOS DE REPUESTOS: el texto puede incluir, despues de una linea
  "=== COTIZACION DE REPUESTOS ===", una tabla o lista de precios de
  repuestos (con columnas tipo Repuesto, Precio s/IVA, Precio c/IVA,
  etc.). Si esa cotizacion esta presente:
  * Para cada pieza a CAMBIAR, asignale su precio usando SIEMPRE la
    columna "Precio c/IVA" (el valor CON IVA).
  * Emparejá cada repuesto de la cotizacion con la pieza del peritaje
    aunque esten escritos distinto. Ejemplos: "PPE Del" del peritaje =
    "Paragolpe delantero" de la cotizacion; "Felpa Bajo Capot" =
    "Felpa / manta bajo capot"; "Optica Izq" = "Optica delantera
    izquierda".
  * El precio es solo el numero entero, sin signo $ ni puntos de miles.
  * Si una pieza a CAMBIAR no aparece en la cotizacion, su precio es 0.
- Tambien se acepta un precio escrito al lado de la pieza en el peritaje
  (ej: "Capot $180000").
- Si NO hay ninguna cotizacion, todas las piezas a CAMBIAR van con
  precio 0.
- Las piezas a REPARAR siempre llevan precio 0.
- Mano de obra: interpretar todas las modalidades de escritura. Ejemplos:
  "7 panos", "pint 7", "7p" -> pintura = 7.
  "chapa 3", "3 dias", "3d", "3 jornadas" -> chapa = 3.
  "12 hs mecanica", "mecanica 12" -> mecanica = 12.
  LAS CANTIDADES PUEDEN TENER DECIMALES (horas y medias, pañitos
  parciales, etc). Interpretar coma Y punto como separador decimal:
  "mec 1,5" o "mec 1.5" -> mecanica = 1.5 (NO 15).
  "pint 7,25" -> pintura = 7.25.
  "0,5 dias chapa" -> chapa = 0.5.
- "carga de gas", "varios", "service", "cristaleria", "vidrieria",
  "alineacion", "electricidad" -> NO son cantidades de unidades; son
  items con un SUBTOTAL EN PESOS. Sumá todos esos subtotales y guardalos
  en el campo manoObra.varios COMO MONTO EN PESOS, no como cantidad.
  Ejemplos:
  * "Cristaleria 1 hora $100.000" -> manoObra.varios = 100000.
  * "Carga de gas $50.000" + "Service $30.000" -> manoObra.varios = 80000.
  * Si no hay items extra, manoObra.varios = 0.
- El numero de siniestro puede venir solo (un numero suelto) o con etiqueta.
- La fecha de inspeccion puede venir como "fecha ip 16/05/26", "16-05-26", \
etc., en cualquier parte del texto. En el Excel suele ser "Dia de
Inspeccion"; si viene como numero serial de Excel, convertilo a fecha.
- La suma asegurada es solo el numero, sin texto pegado.
- Los datos del taller / lugar de inspeccion (nombre, direccion, localidad) \
pueden aparecer en cualquier parte; extraelos si los encontras.
- Las observaciones son la descripcion en prosa del perito (ej: "visto en \
domicilio particular..."), NO los datos de poliza, productor, vigencia, etc.
- Si un dato no aparece, dejalo como cadena vacia "" o 0 segun corresponda.

Devolve UNICAMENTE un objeto JSON valido, sin texto antes ni despues, sin \
backticks, con esta estructura exacta:
{
  "numeroSiniestro": "",
  "fechaSiniestro": "",
  "fechaInspeccion": "",
  "asegurado": "",
  "marca": "",
  "modelo": "",
  "anio": "",
  "dominio": "",
  "chasis": "",
  "kilometraje": "",
  "sumaAsegurada": "",
  "tallerNombre": "",
  "tallerDireccion": "",
  "tallerLocalidad": "",
  "franquicia": 0,
  "manoObra": {"pintura": 0, "chapa": 0, "mecanica": 0,
               "tapiceria": 0, "varios": 0},
  "danos": [{"accion": "CAMBIAR", "pieza": "nombre", "precio": 0}],
  "observaciones": ""
}"""


def _cliente_claude():
    """Crea el cliente de la API de Claude. Lee la key del entorno."""
    api_key = os.environ.get("ANTHROPIC_API_KEY", "").strip()
    if not api_key:
        raise RuntimeError(
            "Falta la variable de entorno ANTHROPIC_API_KEY. "
            "Configurala en Render (Environment).")
    return anthropic.Anthropic(api_key=api_key)


def _extraer_json(texto):
    """Extrae el primer objeto JSON de un texto, tolerando backticks."""
    t = texto.strip()
    # Quitar fences ```json ... ```
    t = re.sub(r"^```(?:json)?\s*|\s*```$", "", t,
               flags=re.IGNORECASE | re.MULTILINE)
    # Tomar desde la primera { hasta la ultima }
    ini = t.find("{")
    fin = t.rfind("}")
    if ini >= 0 and fin > ini:
        t = t[ini:fin + 1]
    return json.loads(t)


def _mejorar_observaciones(texto):
    """
    Reescribe las observaciones del perito en tono tecnico-profesional,
    sin agregar, quitar ni cambiar el sentido de la informacion.
    Si falla la llamada, devuelve el texto original sin tocar.
    """
    if not texto or not texto.strip():
        return texto

    # Limpiar el texto: colapsar espacios y tabs multiples, eliminar
    # saltos de linea sueltos, sacar encabezados como "OBSERVACIONES:"
    # que a veces vienen pegados desde la grilla del peritaje.
    import re as _re
    texto_limpio = str(texto)
    # Colapsar tabs, espacios y saltos de linea en un solo espacio.
    texto_limpio = _re.sub(r"[\t\r\n ]+", " ", texto_limpio).strip()
    # Sacar la palabra "OBSERVACIONES" o "OBSERVACIONES:" al inicio.
    texto_limpio = _re.sub(r"^OBSERVACIONES\s*:?\s*", "", texto_limpio,
                           flags=_re.IGNORECASE)
    if not texto_limpio:
        return ""

    instruccion = (
        "Sos un asistente de redaccion para informes de tasacion de "
        "seguros de automotores. Recibiras un texto crudo con las "
        "observaciones que anoto un perito durante la inspeccion. "
        "Tu tarea: reescribirlas en tono tecnico, formal y profesional, "
        "con oraciones completas, buena ortografia y puntuacion. "
        "REGLAS ESTRICTAS:\n"
        "- NO agregues informacion, conclusiones ni datos que no esten "
        "en el texto original. Nada de inventar.\n"
        "- NO quites ningun hecho ni dato mencionado.\n"
        "- NO cambies el sentido. Si el perito expresa una duda o "
        "hipotesis, debe seguir siendo una duda o hipotesis.\n"
        "- Si el texto original es muy breve (una o dos frases), "
        "expandilo minimamente a una redaccion completa sin agregar "
        "informacion nueva. Ejemplo: 'Motivo: Incendio' se puede "
        "reescribir como 'El motivo de la inspeccion es un siniestro "
        "por incendio.'\n"
        "- Mantene siglas y abreviaturas si aparecen (DT, PT, etc).\n"
        "- Devolve UNICAMENTE el texto reescrito, en una o varias "
        "oraciones prolijas. Sin comillas, sin comentarios, sin "
        "encabezado 'Observaciones:' ni nada por el estilo."
    )

    try:
        cliente = _cliente_claude()
        resp = cliente.messages.create(
            model=MODELO_CLAUDE,
            max_tokens=1000,
            system=instruccion,
            messages=[{"role": "user", "content": texto_limpio}],
        )
        salida = "".join(
            b.text for b in resp.content
            if getattr(b, "type", None) == "text"
        ).strip()
        return salida or texto_limpio
    except Exception:
        # Si falla la llamada, al menos devolver el texto limpio
        # (sin tabs multiples), que es mejor que el crudo.
        return texto_limpio


def extraer_encabezado_desde_adjuntos(archivos, texto_opcional=""):
    """
    Extrae los datos del encabezado desde adjuntos (PDF, imágenes) y
    opcionalmente texto pegado, usando Claude Vision.

    archivos: lista de dicts {"nombre": str, "bytes": bytes}.
    texto_opcional: texto adicional pegado por el usuario (puede ser "").

    Devuelve un dict con los campos encontrados (los que falten vienen
    como ""). Si no hay nada que procesar o la llamada falla, devuelve
    un dict vacío (no rompe el flujo del informe).

    Campos extraídos: numeroSiniestro, asegurado, marca, modelo, anio,
    dominio, chasis, sumaAsegurada, franquicia, tallerTelefono,
    tallerEmail. NO extrae tallerDireccion ni fechaInspeccion (esos los
    carga el perito a mano).
    """
    vacio = {}
    archivos = archivos or []
    texto_opcional = (texto_opcional or "").strip()

    if not archivos and not texto_opcional:
        return vacio

    try:
        # 1) Convertir cada archivo a una o varias imágenes PNG (base64).
        imagenes = []  # lista de dicts {"media_type": str, "data": b64}
        for arch in archivos[:5]:  # tope de 5 archivos
            nombre = (arch.get("nombre") or "").lower()
            datos = arch.get("bytes") or b""
            if not datos:
                continue
            if nombre.endswith(".pdf") or datos[:4] == b"%PDF":
                # Convertir PDF a imágenes (máx 5 páginas)
                imgs_pdf = _pdf_a_imagenes_base64(datos, max_paginas=5)
                imagenes.extend(imgs_pdf)
            else:
                # Imagen directa (jpg, png, etc.)
                b64 = _imagen_a_base64(datos)
                media_type = _detectar_media_type(datos, nombre)
                if b64 and media_type:
                    imagenes.append({"media_type": media_type, "data": b64})
            if len(imagenes) >= 10:  # tope global
                break

        if not imagenes and not texto_opcional:
            return vacio

        # 2) Preparar el contenido multimodal para Claude.
        contenido = []
        for img in imagenes[:10]:
            contenido.append({
                "type": "image",
                "source": {
                    "type": "base64",
                    "media_type": img["media_type"],
                    "data": img["data"],
                },
            })

        prompt_texto = (
            "Analizá los documentos adjuntos (y el texto opcional al "
            "final, si hay) y extraé los datos del siniestro para "
            "cargarlos en un informe de tasación.\n\n"
            "Devolvé SOLO un objeto JSON válido con estos campos. Si un "
            "dato no aparece en los documentos, dejá el string vacío "
            '"":\n'
            "{\n"
            '  "numeroSiniestro": "numero de siniestro, pieza, PZA o STRO",\n'
            '  "asegurado": "nombre completo del asegurado o cliente",\n'
            '  "marca": "marca del vehiculo (FORD, CHEVROLET, etc.)",\n'
            '  "modelo": "modelo del vehiculo (KA, CRUZE, etc.)",\n'
            '  "anio": "año del vehiculo (4 digitos)",\n'
            '  "dominio": "patente/dominio (ej: AB779VJ)",\n'
            '  "chasis": "numero de chasis",\n'
            '  "sumaAsegurada": "suma asegurada en pesos, SOLO el numero entero sin $ ni puntos",\n'
            '  "franquicia": "franquicia en pesos, SOLO el numero entero sin $ ni puntos",\n'
            '  "tallerTelefono": "telefono de contacto",\n'
            '  "tallerEmail": "email de contacto"\n'
            "}\n\n"
            "REGLAS:\n"
            "- Devolvé SOLO el JSON, sin texto antes ni despues, sin "
            "markdown, sin comentarios.\n"
            "- Los montos SOLO como numero entero (ej: 15638400, no "
            "'$15.638.400').\n"
            "- Si un dato aparece de varias formas, elegí la mas "
            "completa.\n"
            "- 'numeroSiniestro' puede figurar como 'N° Siniestro', "
            "'Pieza', 'PZA', 'STRO', 'Siniestro N°'.\n"
            "- Para marcas y modelos: usar la forma mas comun (FORD, no "
            "'Ford Motor Company').\n"
        )
        if texto_opcional:
            prompt_texto += (
                "\n--- TEXTO ADICIONAL PEGADO POR EL USUARIO ---\n"
                + texto_opcional
            )

        contenido.append({"type": "text", "text": prompt_texto})

        # 3) Llamar a Claude Vision.
        cliente = _cliente_claude()
        resp = cliente.messages.create(
            model=MODELO_CLAUDE,
            max_tokens=1500,
            messages=[{"role": "user", "content": contenido}],
        )
        salida = "".join(
            b.text for b in resp.content
            if getattr(b, "type", None) == "text"
        ).strip()

        parsed = _extraer_json(salida)

        # 4) Normalizar los campos esperados.
        campos = ("numeroSiniestro", "asegurado", "marca", "modelo",
                  "anio", "dominio", "chasis", "sumaAsegurada",
                  "franquicia", "tallerTelefono", "tallerEmail")
        resultado = {}
        for k in campos:
            v = parsed.get(k, "")
            if v is None:
                v = ""
            resultado[k] = str(v).strip() if not isinstance(v, (int, float)) else v
        # Montos siempre como número
        for k in ("sumaAsegurada", "franquicia"):
            resultado[k] = a_numero(resultado[k])

        return resultado
    except Exception:
        # Si algo falla, no rompemos el flujo del informe.
        return vacio


def _pdf_a_imagenes_base64(pdf_bytes, max_paginas=5, dpi=150):
    """
    Convierte un PDF a imágenes PNG en base64 (una por página).
    Limita a max_paginas y usa dpi moderado para no inflar el payload.
    """
    try:
        import pymupdf  # type: ignore
        import base64 as _b64
        imagenes = []
        doc = pymupdf.open(stream=pdf_bytes, filetype="pdf")
        try:
            total = min(len(doc), max_paginas)
            # matrix para escalar a dpi deseado (72 dpi = base)
            escala = dpi / 72.0
            mat = pymupdf.Matrix(escala, escala)
            for i in range(total):
                page = doc.load_page(i)
                pix = page.get_pixmap(matrix=mat, alpha=False)
                png_bytes = pix.tobytes("png")
                imagenes.append({
                    "media_type": "image/png",
                    "data": _b64.b64encode(png_bytes).decode("ascii"),
                })
        finally:
            doc.close()
        return imagenes
    except Exception:
        return []


def _imagen_a_base64(datos):
    """Convierte bytes de imagen a string base64 ASCII."""
    try:
        import base64 as _b64
        return _b64.b64encode(datos).decode("ascii")
    except Exception:
        return ""


def _detectar_media_type(datos, nombre):
    """Detecta el media type de una imagen por firma o extensión."""
    # Firmas conocidas
    if datos[:3] == b"\xff\xd8\xff":
        return "image/jpeg"
    if datos[:8] == b"\x89PNG\r\n\x1a\n":
        return "image/png"
    if datos[:6] in (b"GIF87a", b"GIF89a"):
        return "image/gif"
    if datos[:4] == b"RIFF" and datos[8:12] == b"WEBP":
        return "image/webp"
    # Fallback por extensión
    ext = nombre.rsplit(".", 1)[-1].lower() if "." in nombre else ""
    mapa = {
        "jpg": "image/jpeg", "jpeg": "image/jpeg",
        "png": "image/png", "gif": "image/gif", "webp": "image/webp",
    }
    return mapa.get(ext, "")


def _fecha_dd_mm_aaaa(valor):
    """
    Normaliza una fecha al formato DD/MM/AAAA. Acepta:
      - "2026-05-16"     (ISO)             -> "16/05/2026"
      - "2026/05/16"                       -> "16/05/2026"
      - "16-05-2026" / "16-5-26"           -> "16/05/2026"
      - "16/05/2026" o "16/5/26"           -> "16/05/2026" (ya esta)
      - numero serial Excel (ej 46177)     -> "16/05/2026"
      - cualquier otra cosa -> se devuelve tal cual (no se rompe)
    """
    if not valor:
        return ""
    s = str(valor).strip()
    if not s:
        return ""

    # Numero serial de Excel (dias desde 1899-12-30)
    try:
        if s.replace(".", "").replace(",", "").isdigit():
            n = int(float(s.replace(",", ".")))
            if 1000 < n < 80000:
                import datetime as _dt
                base = _dt.date(1899, 12, 30)
                d = base + _dt.timedelta(days=n)
                return f"{d.day:02d}/{d.month:02d}/{d.year:04d}"
    except (ValueError, OverflowError):
        pass

    # Reemplazar separadores varios por uno solo
    import re as _re
    partes = _re.split(r"[-/\.]", s)
    if len(partes) != 3:
        return s  # no se puede parsear, dejar tal cual
    try:
        a, b, c = (int(p) for p in partes)
    except ValueError:
        return s

    # Si la primera parte tiene 4 digitos -> es AAAA-MM-DD
    if len(partes[0]) == 4:
        anio, mes, dia = a, b, c
    else:
        dia, mes, anio = a, b, c
        if anio < 100:  # "26" -> "2026"
            anio += 2000

    # Validacion basica
    if not (1 <= dia <= 31 and 1 <= mes <= 12):
        return s
    return f"{dia:02d}/{mes:02d}/{anio:04d}"


def parsear_texto(texto):
    """
    Entiende el texto libre de la peritacion usando la API de Claude
    y devuelve el dict de datos estructurado.

    Si Claude falla o el texto esta vacio, devuelve data_vacia().
    """
    data = data_vacia()
    if not texto or not texto.strip():
        return data

    cliente = _cliente_claude()
    respuesta = cliente.messages.create(
        model=MODELO_CLAUDE,
        max_tokens=8000,
        system=_PROMPT_SISTEMA_PARSEO,
        messages=[{
            "role": "user",
            "content": "Texto de la peritacion:\n\n" + texto,
        }],
    )

    # Juntar el texto devuelto por Claude
    salida = ""
    for bloque in respuesta.content:
        if getattr(bloque, "type", None) == "text":
            salida += bloque.text

    try:
        parsed = _extraer_json(salida)
    except Exception:
        # Si Claude no devolvio JSON valido, no se pierde el informe:
        # se devuelve data vacia y el resto del flujo sigue.
        return data

    # Volcar los campos al dict, con valores por defecto seguros
    for campo in ("numeroSiniestro", "fechaSiniestro", "fechaInspeccion",
                  "asegurado", "marca", "modelo", "anio", "dominio",
                  "chasis", "kilometraje", "sumaAsegurada",
                  "tallerNombre", "tallerDireccion", "tallerLocalidad",
                  "observaciones"):
        v = parsed.get(campo, "")
        data[campo] = str(v).strip() if v is not None else ""

    # Normalizar las fechas a formato DD/MM/AAAA, vengan como vengan.
    data["fechaSiniestro"] = _fecha_dd_mm_aaaa(data["fechaSiniestro"])
    data["fechaInspeccion"] = _fecha_dd_mm_aaaa(data["fechaInspeccion"])

    data["franquicia"] = a_numero(parsed.get("franquicia", 0))

    mo = parsed.get("manoObra", {}) or {}
    # Cantidades: pueden ser decimales (ej: 1.5 horas de mecanica).
    # Varios: siempre monto entero en pesos.
    for k in ("pintura", "chapa", "mecanica", "tapiceria"):
        data["manoObra"][k] = a_numero_decimal(mo.get(k, 0))
    data["manoObra"]["varios"] = a_numero(mo.get("varios", 0))

    danos = []
    for d in parsed.get("danos", []) or []:
        accion = normalizar(d.get("accion", ""))
        accion = "CAMBIAR" if accion not in ("REPARAR",) else "REPARAR"
        pieza = str(d.get("pieza", "")).strip()
        # Solo las piezas a CAMBIAR llevan precio; las REPARAR van en 0.
        precio = a_numero(d.get("precio", 0)) if accion == "CAMBIAR" else 0
        if pieza:
            danos.append({"accion": accion, "pieza": pieza,
                          "precio": precio})
    data["danos"] = danos

    # Reescritura automatica de observaciones en tono tecnico. Si la
    # llamada falla (timeout, red, etc), la funcion devuelve el texto
    # original y el informe se genera igual.
    if data.get("observaciones"):
        try:
            data["observaciones"] = _mejorar_observaciones(data["observaciones"])
        except Exception:
            pass

    return data
# ============================================================
#  PRESERVACION DE IMAGENES (LOGOS)
# ============================================================

def _reinyectar_imagenes(plantilla_bytes, generado_bytes):
    """
    Restaura los logos de la plantilla en el informe generado, PERO solo
    si openpyxl los perdio al guardar.

    En algunos entornos (segun version de openpyxl / del sistema), al
    hacer load_workbook + save las imagenes ancladas se descartan.
    Esta funcion lo detecta y, si faltan, reinyecta a nivel ZIP:
      - xl/media/*      (las imagenes)
      - xl/drawings/*   (los anclajes y sus rels)
    y ademas, para cada hoja que tenia imagenes:
      - agrega la relacion al drawing en su xl/worksheets/_rels/sheetN.xml.rels
      - inserta la etiqueta <drawing r:id="..."/> dentro de sheetN.xml
        (sin esto, Excel NO muestra la imagen aunque el archivo este)
    y declara los drawings en [Content_Types].xml.

    Si openpyxl ya conservo las imagenes, NO toca nada.
    Si la plantilla no tiene imagenes, tampoco actua.
    """
    try:
        zp = zipfile.ZipFile(BytesIO(plantilla_bytes))
        zg = zipfile.ZipFile(BytesIO(generado_bytes))
    except zipfile.BadZipFile:
        return generado_bytes

    nombres_p = set(zp.namelist())
    nombres_g = set(zg.namelist())

    # La plantilla no tiene imagenes -> nada que reinyectar
    if not any(n.startswith("xl/media/") for n in nombres_p):
        return generado_bytes

    # El generado YA conservo las imagenes -> no tocar
    if any(n.startswith("xl/media/") for n in nombres_g):
        return generado_bytes

    # --- Mapear: cada sheetN -> que drawing usa (segun la plantilla) ---
    # En la plantilla, xl/worksheets/_rels/sheetN.xml.rels apunta al drawing.
    sheet_a_drawing = {}  # 'sheet1' -> 'drawing1'
    for n in nombres_p:
        m = re.match(r"xl/worksheets/_rels/(sheet\d+)\.xml\.rels$", n)
        if m:
            contenido = zp.read(n).decode("utf-8")
            md = re.search(r'Target="\.\./drawings/(drawing\d+\.xml)"',
                           contenido)
            if md:
                sheet_a_drawing[m.group(1)] = md.group(1)

    # Archivos de imagen a copiar tal cual de la plantilla
    grupo = [n for n in nombres_p
             if n.startswith("xl/media/") or n.startswith("xl/drawings/")]

    # --- Fusionar [Content_Types].xml ---
    ct = zg.read("[Content_Types].xml").decode("utf-8")
    extras = ""
    for dw in sorted(set(sheet_a_drawing.values())):
        part = "/xl/drawings/" + dw
        if part not in ct:
            extras += ('<Override PartName="%s" ContentType='
                       '"application/vnd.openxmlformats-officedocument'
                       '.drawing+xml"/>' % part)
    if 'Extension="png"' not in ct:
        extras = ('<Default Extension="png" ContentType="image/png"/>'
                  + extras)
    if extras:
        ct = ct.replace("</Types>", extras + "</Types>")

    def _agregar_drawing_a_rels(rels_xml, drawing_file):
        """Garantiza una relacion al drawing y devuelve (xml_nuevo, rId).
        Si ya hay una relacion de tipo drawing, reutiliza su Id."""
        # Si ya existe una relacion de drawing, usar ese rId
        m = re.search(
            r'<Relationship[^>]*relationships/drawing"[^>]*Id="(rId\d+)"',
            rels_xml)
        if not m:
            m = re.search(
                r'<Relationship[^>]*Id="(rId\d+)"[^>]*relationships/drawing"',
                rels_xml)
        if m:
            rid = m.group(1)
            # Reescribir esa relacion para que apunte al drawing correcto
            rels_xml = re.sub(
                r'<Relationship[^>]*relationships/drawing"[^>]*/>',
                '', rels_xml)
            rels_xml = re.sub(
                r'<Relationship[^>]*Id="' + rid + r'"[^>]*/>',
                '', rels_xml)
        else:
            usados = re.findall(r'Id="rId(\d+)"', rels_xml)
            n = max([int(x) for x in usados], default=0) + 1
            rid = "rId%d" % n
        rel = ('<Relationship Id="%s" Type="http://schemas.openxmlformats'
               '.org/officeDocument/2006/relationships/drawing" '
               'Target="../drawings/%s"/>' % (rid, drawing_file))
        nuevo = rels_xml.replace("</Relationships>",
                                 rel + "</Relationships>")
        return nuevo, rid

    def _insertar_drawing_en_sheet(sheet_xml, rid):
        """Garantiza <drawing r:id=.../> antes de </worksheet>.
        Si ya hay una etiqueta <drawing>, la reemplaza por la correcta.
        Tambien garantiza que el namespace 'r' este declarado, ya que
        sin xmlns:r el atributo r:id provoca 'unbound prefix' en Excel."""
        # Quitar cualquier <drawing.../> existente (puede estar roto)
        sheet_xml = re.sub(r'<drawing\b[^>]*/>', '', sheet_xml)
        # Asegurar que el tag raiz <worksheet ...> declare xmlns:r
        if "xmlns:r=" not in sheet_xml:
            ns = ('xmlns:r="http://schemas.openxmlformats.org/'
                  'officeDocument/2006/relationships"')
            sheet_xml = re.sub(
                r'(<worksheet\b)', r'\1 ' + ns, sheet_xml, count=1)
        tag = '<drawing r:id="%s"/>' % rid
        return sheet_xml.replace("</worksheet>", tag + "</worksheet>")

    out = BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        escritos = set()

        for item in zg.namelist():
            # Content_Types fusionado
            if item == "[Content_Types].xml":
                zout.writestr(item, ct)
                escritos.add(item)
                continue

            # Los .rels de hojas se manejan junto a su sheet -> saltar aqui
            if re.match(r"xl/worksheets/_rels/sheet\d+\.xml\.rels$", item):
                continue

            # sheetN.xml: insertar la etiqueta <drawing>
            ms = re.match(r"xl/worksheets/(sheet\d+)\.xml$", item)
            if ms and ms.group(1) in sheet_a_drawing:
                sheet_id = ms.group(1)
                dw = sheet_a_drawing[sheet_id]
                rels_path = ("xl/worksheets/_rels/%s.xml.rels"
                             % sheet_id)
                if rels_path in nombres_g:
                    rels_xml = zg.read(rels_path).decode("utf-8")
                else:
                    rels_xml = ('<?xml version="1.0" encoding="UTF-8" '
                                'standalone="yes"?>'
                                '<Relationships xmlns="http://schemas'
                                '.openxmlformats.org/package/2006/'
                                'relationships"></Relationships>')
                rels_nuevo, rid = _agregar_drawing_a_rels(rels_xml, dw)
                sheet_xml = zg.read(item).decode("utf-8")
                sheet_xml = _insertar_drawing_en_sheet(sheet_xml, rid)
                zout.writestr(item, sheet_xml)
                zout.writestr(rels_path, rels_nuevo)
                escritos.add(item)
                escritos.add(rels_path)
                continue

            zout.writestr(item, zg.read(item))
            escritos.add(item)

        # Copiar imagenes y drawings de la plantilla
        for item in grupo:
            if item not in escritos:
                zout.writestr(item, zp.read(item))
                escritos.add(item)

    return out.getvalue()



# ============================================================
#  COMPLETAR LA PLANTILLA (PRESERVA LOGOS Y FORMATO)
# ============================================================

def _rango_de(ws, fila, columna):
    """Devuelve el rango combinado que contiene a (fila, columna), o None."""
    for rango in ws.merged_cells.ranges:
        if (rango.min_row <= fila <= rango.max_row and
                rango.min_col <= columna <= rango.max_col):
            return rango
    return None


def _valor_celda(ws, fila, columna):
    """
    Lee el valor de una celda de forma segura.
    Si es parte de un rango combinado, el valor real esta en la celda
    ancla (esquina superior izquierda).
    """
    rango = _rango_de(ws, fila, columna)
    if rango is not None:
        return ws.cell(row=rango.min_row, column=rango.min_col).value
    return ws.cell(row=fila, column=columna).value


def _escribir(ws, fila, columna, valor):
    """
    Escribe un valor en una celda manejando celdas combinadas.

    FIX del error "'MergedCell' object attribute 'value' is read-only":
    si la celda destino cae dentro de un rango combinado, se DESHACE el
    merge antes de escribir y se vuelve a combinar despues. Asi la celda
    deja de ser una MergedCell inmutable.
    """
    rango = _rango_de(ws, fila, columna)
    if rango is not None:
        rango_str = str(rango)
        anc_row, anc_col = rango.min_row, rango.min_col
        ws.unmerge_cells(rango_str)
        ws.cell(row=anc_row, column=anc_col).value = valor
        ws.merge_cells(rango_str)
    else:
        ws.cell(row=fila, column=columna).value = valor


def _es_celda_etiqueta(valor_celda, buscado):
    """
    Determina si una celda ES una etiqueta de campo (no que la contiene
    como substring accidental, ej. 'ANO' dentro de 'EMILIANO').

    El texto buscado debe aparecer como PALABRA COMPLETA dentro de una
    celda que se comporte como etiqueta (texto corto o terminado en ':').
    """
    v = normalizar(valor_celda)
    b = normalizar(buscado)
    if not v or not b:
        return False
    # 'b' como palabra completa dentro de 'v'
    if not re.search(r"(?<![A-Z0-9])" + re.escape(b) + r"(?![A-Z0-9])", v):
        return False
    # La celda debe parecer etiqueta: termina en ':'/'=' o es texto corto
    v_limpio = v.rstrip(": =").strip()
    if v.endswith(":") or v.endswith("=") or len(v_limpio) <= 35:
        return True
    return False


def _buscar_celda_etiqueta(ws, textos_buscados, ocurrencia=1):
    """
    Busca la celda que ES una etiqueta y devuelve la coordenada (fila, columna)
    de la celda donde escribir el valor (la primera celda vacia a la derecha).
    Devuelve None si no encuentra la etiqueta.
    """
    encontradas = 0
    for fila_celdas in ws.iter_rows():
        for celda in fila_celdas:
            if not normalizar(celda.value):
                continue
            for buscado in textos_buscados:
                if _es_celda_etiqueta(celda.value, buscado):
                    encontradas += 1
                    if encontradas == ocurrencia:
                        fila, col = celda.row, celda.column
                        for c in range(col + 1, col + 9):
                            if _valor_celda(ws, fila, c) in (None, ""):
                                return (fila, c)
                        return (fila, col + 1)
    return None


def _copiar_imagenes(ws_origen, ws_destino):
    """
    Copia las imagenes (logos) de una hoja a otra. Necesario porque
    openpyxl.copy_worksheet() no las clona automaticamente.

    NOTA: esta funcion solo deja MARCADAS las imagenes en la hoja
    destino (con los bytes correctos). Las dimensiones y posiciones
    finales se corrigen despues de guardar el archivo, en el paso
    de post-procesamiento del ZIP (_arreglar_drawings_en_zip), porque
    openpyxl tiene un bug con copy_worksheet + anchors que a veces
    rompe el renderizado en Excel.
    """
    try:
        from copy import deepcopy
        from io import BytesIO as _BIO
        from openpyxl.drawing.image import Image as _XLImage
        for img in getattr(ws_origen, "_images", []):
            datos = None
            ref = getattr(img, "ref", None)
            if hasattr(ref, "read"):
                try:
                    ref.seek(0)
                    datos = ref.read()
                    ref.seek(0)
                except Exception:
                    datos = None
            if not datos:
                continue
            nueva = _XLImage(_BIO(datos))
            try:
                anchor_copia = deepcopy(img.anchor)
                if hasattr(anchor_copia, "pic"):
                    anchor_copia.pic = None
                if hasattr(anchor_copia, "graphicFrame"):
                    anchor_copia.graphicFrame = None
                nueva.anchor = anchor_copia
            except Exception:
                pass
            ws_destino.add_image(nueva)
    except Exception:
        pass


def _arreglar_drawings_en_zip(archivo_bytes, hoja_origen_titulo,
                              hojas_clon_titulos):
    """
    Post-procesamiento: pisa el contenido del drawing.xml de cada
    hoja clonada con una copia EXACTA del drawing de la hoja origen,
    reescribiendo los IDs internos (cNvPr id) para que sean unicos.

    Esto es necesario porque la copia que hace openpyxl deja un XML
    valido en estructura pero con metadatos que Excel a veces no
    renderiza (problema visual, no de datos). Pisar el XML con uno
    bueno garantiza que las imagenes se vean igual que en la hoja
    original.
    """
    import zipfile
    import re
    from io import BytesIO as _BIO

    if not hojas_clon_titulos:
        return archivo_bytes

    # 1) Abrir el zip y averiguar el mapeo sheet<->drawing
    zin = zipfile.ZipFile(_BIO(archivo_bytes), "r")
    try:
        # workbook.xml lista las hojas en el orden que apareceran
        wb_xml = zin.read("xl/workbook.xml").decode("utf-8")
        # mapa nombre_hoja -> numero (sheetN). Las hojas aparecen
        # en orden en workbook.xml, asi que numeramos por orden.
        nombres_orden = re.findall(r'<sheet[^>]*name="([^"]+)"', wb_xml)
        sheet_por_nombre = {nombre: i + 1 for i, nombre in enumerate(nombres_orden)}

        # Para cada sheetN.xml.rels, encontrar el drawing al que
        # apunta (sheetX -> drawingY).
        drawing_de_sheet = {}
        for sheet_num in sheet_por_nombre.values():
            rels_path = f"xl/worksheets/_rels/sheet{sheet_num}.xml.rels"
            try:
                rels = zin.read(rels_path).decode("utf-8")
            except KeyError:
                continue
            m = re.search(r'Target="[^"]*drawing(\d+)\.xml"', rels)
            if m:
                drawing_de_sheet[sheet_num] = int(m.group(1))

        # 2) Leer el contenido del drawing de la hoja origen.
        sheet_origen = sheet_por_nombre.get(hoja_origen_titulo)
        if not sheet_origen:
            return archivo_bytes
        drawing_origen = drawing_de_sheet.get(sheet_origen)
        if not drawing_origen:
            return archivo_bytes
        xml_origen = zin.read(
            f"xl/drawings/drawing{drawing_origen}.xml").decode("utf-8")
        rels_origen = zin.read(
            f"xl/drawings/_rels/drawing{drawing_origen}.xml.rels").decode("utf-8")

        # 3) Construir el contenido a inyectar para cada hoja clonada.
        cambios = {}  # path -> contenido nuevo (bytes)
        # Pedimos los IDs maximos vistos para evitar colisiones.
        id_counter = 1000  # arrancamos alto para no colisionar
        for nombre_clon in hojas_clon_titulos:
            sheet_clon = sheet_por_nombre.get(nombre_clon)
            if not sheet_clon:
                continue
            drawing_clon = drawing_de_sheet.get(sheet_clon)
            if not drawing_clon:
                continue
            # Generar XML del drawing del clon, igual al origen pero
            # con cNvPr id distintos para evitar colisiones globales.
            xml_nuevo = xml_origen
            ids_a_reemplazar = re.findall(r'cNvPr id="(\d+)"', xml_nuevo)
            for viejo in ids_a_reemplazar:
                id_counter += 1
                xml_nuevo = xml_nuevo.replace(
                    f'cNvPr id="{viejo}"',
                    f'cNvPr id="{id_counter}"', 1)
            cambios[f"xl/drawings/drawing{drawing_clon}.xml"] = xml_nuevo.encode("utf-8")
            # Los .rels los dejamos como esten (openpyxl ya genero
            # las referencias correctas a las imagenes copiadas en
            # xl/media/).
            # PERO si por algun motivo la copia genero referencias a
            # imagenes inexistentes, mejor reescribimos el .rels para
            # que apunte a las MISMAS imagenes que la hoja origen.
            cambios[f"xl/drawings/_rels/drawing{drawing_clon}.xml.rels"] = \
                rels_origen.encode("utf-8")
    finally:
        zin.close()

    # 4) Reescribir el ZIP con los cambios.
    salida = _BIO()
    zin = zipfile.ZipFile(_BIO(archivo_bytes), "r")
    zout = zipfile.ZipFile(salida, "w", zipfile.ZIP_DEFLATED)
    try:
        for item in zin.infolist():
            datos = cambios.get(item.filename) or zin.read(item.filename)
            zout.writestr(item, datos)
    finally:
        zin.close()
        zout.close()
    return salida.getvalue()


def _agregar_sufijo_titulo_danos(ws, sufijo):
    """
    Agrega un sufijo (ej: " (hoja 2 de piezas)") a la celda que tenga
    el titulo "DESCRIPCION DE DAÑOS" de la hoja clonada.
    """
    try:
        for fila in ws.iter_rows():
            for celda in fila:
                v = celda.value
                if isinstance(v, str) and "DESCRIPCI" in v.upper() and "DA" in v.upper():
                    # Evitar agregar dos veces el mismo sufijo.
                    if sufijo not in v:
                        celda.value = v + sufijo
                    return
    except Exception:
        pass


def _formula_total_repuestos(hojas, data):
    """
    Arma una fórmula Excel que sume los precios de las piezas a
    CAMBIAR en la Hoja 2 (y sus clones si hay). Ejemplos:
      =SUM(Hoja2!F12:F41)
      =SUM(Hoja2!F12:F41)+SUM('Hoja2 (2)'!F12:F41)

    Si por alguna razón no se puede armar la fórmula, cae a un
    número fijo con la suma calculada en Python.
    """
    try:
        PIEZAS_POR_HOJA = 30
        # Detectar la columna de precio y la fila inicial mirando la Hoja 2.
        if len(hojas) <= 1:
            raise ValueError("no hay Hoja2")
        ws2 = hojas[1]
        fila_inicio = None
        col_precio = None
        for fila_celdas in ws2.iter_rows():
            for celda in fila_celdas:
                v = normalizar(celda.value)
                if v in ("ACCION", "ACCION:"):
                    fila_inicio = celda.row + 1
                if "PRECIO" in v:
                    col_precio = celda.column
            if fila_inicio and col_precio:
                break
        if not (fila_inicio and col_precio):
            raise ValueError("no se detectaron encabezados")

        col_letra = _col_letra(col_precio)
        fila_fin = fila_inicio + PIEZAS_POR_HOJA - 1

        # Todas las hojas cuyo título arranca como el de Hoja2.
        titulo_base = ws2.title
        hojas_piezas = [ws2.title]
        for ws in hojas[0].parent.worksheets:
            if ws.title != titulo_base and ws.title.startswith(titulo_base + " ("):
                hojas_piezas.append(ws.title)

        partes = []
        for titulo in hojas_piezas:
            # Nombres con espacios/paréntesis se escapan con comillas simples.
            titulo_ref = f"'{titulo}'" if (" " in titulo or "(" in titulo) else titulo
            partes.append(f"SUM({titulo_ref}!{col_letra}{fila_inicio}:{col_letra}{fila_fin})")
        return "=" + "+".join(partes)
    except Exception:
        # Fallback: número fijo con la suma calculada.
        return sum(
            (d.get("precio", 0) or 0)
            for d in data.get("danos", [])
            if d.get("accion") == "CAMBIAR"
        )


def _col_letra(n):
    """Convierte índice de columna 1-based a letra Excel (1='A', 27='AA')."""
    letras = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        letras = chr(65 + r) + letras
    return letras


def completar_plantilla(plantilla_bytes, data):
    """
    Carga la plantilla virgen (formato NUEVO: una sola hoja con todo),
    la completa con los datos y devuelve los bytes del informe final.

    Estructura de la plantilla:
      - Filas 7-15: datos del encabezado (siniestro, asegurado, vehiculo,
        taller, telefono/email).
      - Fila 16: titulo "DESCRIPCIÓN DE DAÑOS".
      - Fila 17: encabezados (ACCION / PIEZA / PRECIO).
      - Filas 18-50: 33 filas disponibles para piezas.
      - Fila 51: titulos "OBSERVACIONES" (A:C) y "COTIZACIÓN DE MANO DE
        OBRA" (D:H).
      - Filas 52-57: a la IZQUIERDA (A:C) van las observaciones; a la
        DERECHA (D:H) van los items de mano de obra (5 filas:
        Pintura, Chapa, Mecanica, Tapiceria, Varios).
      - Filas 58-61: totales con FORMULAS que ya vienen en la plantilla
        (NO tocar).

    Las fórmulas G58 (MANO DE OBRA), G59 (REPUESTOS), G60 (FRANQUICIA)
    y G61 (NETO) se calculan automáticamente en Excel.
    """
    wb = load_workbook(BytesIO(plantilla_bytes))
    ws = wb.worksheets[0]  # Solo hay una hoja

    # ---------- 1) DATOS DEL ENCABEZADO ----------
    # Mapeo de etiqueta -> celda donde va el VALOR (al lado derecho).
    # La etiqueta esta en una celda y el valor en la siguiente, segun los
    # merges detectados en la plantilla.

    # Fila 7: N° Siniestro | Nombre Aseg/Terc | Fecha IP
    _escribir(ws, 7, 3, data.get("numeroSiniestro", ""))     # C7
    _escribir(ws, 7, 5, data.get("asegurado", ""))           # E7:F7
    _escribir(ws, 7, 8, data.get("fechaInspeccion", ""))     # H7

    # Fila 9: Marca | Modelo | Año
    _escribir(ws, 9, 2, data.get("marca", ""))               # B9:C9
    _escribir(ws, 9, 5, data.get("modelo", ""))              # E9:F9
    _escribir(ws, 9, 8, data.get("anio", ""))                # H9

    # Fila 11: Dominio | Chasis | Suma aseg | Franq.
    _escribir(ws, 11, 2, data.get("dominio", ""))            # B11
    _escribir(ws, 11, 4, data.get("chasis", ""))             # D11
    _escribir(ws, 11, 6, a_numero(data.get("sumaAsegurada", 0)))  # F11
    franq = a_numero(data.get("franquicia", 0))
    if franq:
        _escribir(ws, 11, 8, franq)                          # H11

    # Fila 13: Taller | Dirección IP (+ Localidad concatenada)
    _escribir(ws, 13, 2, data.get("tallerNombre", ""))       # B13:C13
    dir_insp = (data.get("tallerDireccion") or "").strip()
    loc_insp = (data.get("tallerLocalidad") or "").strip()
    if dir_insp and loc_insp:
        direccion_completa = f"{dir_insp} - {loc_insp}"
    else:
        direccion_completa = dir_insp or loc_insp
    _escribir(ws, 13, 5, direccion_completa)                 # E13:H13

    # Fila 15: Teléfono | Email
    _escribir(ws, 15, 2, data.get("tallerTelefono", ""))     # B15:C15
    _escribir(ws, 15, 5, data.get("tallerEmail", ""))        # E15:H15

    # ---------- 2) PIEZAS: con expansion dinamica si hay > 33 ----------
    # La plantilla trae 33 filas de piezas (18-50). Si el peritaje tiene
    # mas, insertamos filas antes de la fila 51 (OBSERVACIONES), de
    # modo que toda la seccion de abajo se desplaza y las formulas de
    # totales se actualizan al nuevo rango.
    FILA_INICIO_PIEZAS = 18
    FILA_FIN_PIEZAS_BASE = 50  # fin en la plantilla sin ampliar

    danos = data.get("danos") or []
    extras = max(0, len(danos) - (FILA_FIN_PIEZAS_BASE - FILA_INICIO_PIEZAS + 1))
    if extras > 0:
        _ampliar_filas_piezas(ws, FILA_FIN_PIEZAS_BASE, extras)

    FILA_FIN_PIEZAS = FILA_FIN_PIEZAS_BASE + extras
    MAX_PIEZAS = FILA_FIN_PIEZAS - FILA_INICIO_PIEZAS + 1
    piezas_a_escribir = danos[:MAX_PIEZAS]

    for i, dano in enumerate(piezas_a_escribir):
        fila = FILA_INICIO_PIEZAS + i
        accion = (dano.get("accion") or "").upper()
        pieza = dano.get("pieza", "")
        _escribir(ws, fila, 1, accion)      # A: acción
        _escribir(ws, fila, 2, pieza)       # B:F (merged): pieza
        if accion == "CAMBIAR":
            precio = a_numero(dano.get("precio", 0) or 0)
            _escribir(ws, fila, 7, precio)  # G:H (merged): precio

    # ---------- 3) OBSERVACIONES ----------
    # Despues de ampliar: la celda OBSERVACIONES "titulo" se desplazo
    # extras filas hacia abajo. Las filas de contenido son:
    #    (52 + extras) a (57 + extras)
    fila_obs_titulo = 51 + extras
    fila_obs_inicio = 52 + extras
    fila_obs_fin = 57 + extras
    _preparar_celda_observaciones(ws, fila_obs_inicio, fila_obs_fin)

    obs = (data.get("observaciones") or "").strip()
    if obs:
        _escribir(ws, fila_obs_inicio, 1, obs)
        try:
            from openpyxl.styles import Alignment
            celda = ws.cell(row=fila_obs_inicio, column=1)
            celda.alignment = Alignment(
                wrap_text=True,
                vertical="top",
                horizontal=(celda.alignment.horizontal if celda.alignment
                            else "left")
            )
        except Exception:
            pass

    # ---------- 4) MANO DE OBRA (desplazadas por extras) ----------
    # Las filas de MO estaban en 53-57; despues de insertar extras
    # filas antes de la 51, pasan a 53+extras a 57+extras.
    FILA_INICIO_MO = 53 + extras
    FILA_FIN_MO = 57 + extras
    MAX_MO = FILA_FIN_MO - FILA_INICIO_MO + 1  # 5 filas
    # La fila "Varios" (con E:F merged) tambien se desplaza.
    FILA_VARIOS = 57 + extras

    def _escribir_mo(fila, concepto, cantidad, unitario):
        """
        Escribe una fila de mano de obra. Las filas 53-56 tienen formula
        G=E*F que se calcula sola. La fila 57 ("Varios") tiene E57:F57
        mergeadas, asi que no se puede escribir cantidad/unitario por
        separado: ahi se escribe el TOTAL directamente en G57 (sin
        formula), que igualmente entra en el SUM de G58.
        """
        _escribir(ws, fila, 4, concepto or "")
        if fila == FILA_VARIOS:
            # Fila especial "Varios": E:F mergeado. Escribir total
            # directo en G (merged G:H).
            cant_dec = a_numero_decimal(cantidad)
            unit_int = a_numero(unitario)
            total = cant_dec * unit_int
            if total == 0 and unit_int > 0:
                # Si solo vino el monto como "unitario" (sin cantidad),
                # usar ese valor como total.
                total = unit_int
            if total:
                _escribir(ws, fila, 7, int(total) if total == int(total) else total)
        else:
            # Cantidades pueden ser decimales (ej: 1.5 horas de mecanica).
            cant_dec = a_numero_decimal(cantidad)
            _escribir(ws, fila, 5, cant_dec)   # E
            # Si la cantidad tiene parte decimal, cambiar formato de la
            # celda para que Excel lo muestre con decimales (ej "1,5")
            # y no redondee visualmente (ej "2"). Si es entero, mantener
            # el formato original de la plantilla.
            try:
                if isinstance(cant_dec, float) and cant_dec != int(cant_dec):
                    ws.cell(row=fila, column=5).number_format = "0.##"
            except Exception:
                pass
            if a_numero(unitario):
                _escribir(ws, fila, 6, a_numero(unitario))   # F

    items_mo = data.get("manoObra_items") or []
    if items_mo:
        # Modo flexible: usar los items tal cual los trajo el peritaje.
        for i, item in enumerate(items_mo[:MAX_MO]):
            fila = FILA_INICIO_MO + i
            _escribir_mo(fila, item.get("concepto", ""),
                         item.get("cantidad", 0),
                         item.get("unitario", 0))
        # Filas sobrantes: dejar cantidad y unitario en 0.
        for i in range(len(items_mo), MAX_MO):
            fila = FILA_INICIO_MO + i
            if fila != FILA_VARIOS:
                _escribir(ws, fila, 5, 0)
                _escribir(ws, fila, 6, 0)
    else:
        # Modo clasico: usar manoObra con los 4 conceptos fijos + varios.
        mo = data.get("manoObra") or {}
        conceptos = [
            ("Pintura",  mo.get("pintura", 0),   mo.get("pinturaValor", 0)),
            ("Chapa",    mo.get("chapa", 0),     mo.get("chapaValor", 0)),
            ("Mecanica", mo.get("mecanica", 0),  mo.get("mecanicaValor", 0)),
            ("Tapiceria",mo.get("tapiceria", 0), mo.get("tapiceriaValor", 0)),
            # Varios: en modo clasico viene como monto directo. Se maneja
            # en _escribir_mo (va a G57 como total).
            ("Varios",   0, mo.get("varios", 0)),
        ]
        for i, (concepto, cant, unit) in enumerate(conceptos[:MAX_MO]):
            fila = FILA_INICIO_MO + i
            _escribir_mo(fila, concepto, cant, unit)

    # ---------- 5) GUARDAR ----------
    salida = BytesIO()
    wb.save(salida)
    salida.seek(0)
    generado = salida.read()

    # PRESERVAR LOGOS: openpyxl a veces no copia las imagenes de la
    # plantilla al guardar. _reinyectar_imagenes las copia desde la
    # plantilla original a nivel ZIP, si faltan.
    generado = _reinyectar_imagenes(plantilla_bytes, generado)

    # Si se insertaron filas extra (peritaje > 33 piezas), las
    # imagenes/dibujos que estaban por debajo quedaron en su posicion
    # original y hay que desplazarlas tambien. Esto se hace a nivel
    # XML en el ZIP porque openpyxl no mueve los anchors de shapes.
    if extras > 0:
        # Fila de corte 1-indexed: desplazamos las imagenes cuya
        # fila XML sea >= (51-1 = 50 en 0-indexed, osea Excel 51).
        generado = _desplazar_imagenes_debajo_de(generado, 51, extras)

    # Arreglar Content_Types.xml DESPUES de reinyectar las imagenes:
    # openpyxl a veces se olvida de declarar las extensiones (ej:
    # jpeg) y Excel no puede abrir el archivo. Esto garantiza que
    # cada extension de imagen presente tenga su ContentType.
    generado = _arreglar_content_types_imagenes(generado)

    return generado


def _desplazar_imagenes_debajo_de(archivo_bytes, fila_excel_1indexed, extras):
    """
    Desplaza las imagenes/dibujos que estan por debajo de una cierta
    fila `extras` filas hacia abajo, modificando los <xdr:row> en los
    drawings XML a nivel ZIP.

    openpyxl insert_rows() no actualiza los anchors de las imagenes
    (sobre todo las que son shapes/DrawingML complejos que openpyxl no
    maneja). Esto lo arregla a nivel XML.

    Las filas en el XML del drawing son 0-indexed, por eso restamos 1
    al convertir desde el 1-indexed de Excel.
    """
    if extras <= 0:
        return archivo_bytes
    try:
        import zipfile
        import re as _re
        from io import BytesIO as _BIO

        fila_corte_xml = fila_excel_1indexed - 1  # convertir a 0-indexed
        cambios = {}

        zin = zipfile.ZipFile(_BIO(archivo_bytes), "r")
        try:
            for nombre in zin.namelist():
                if not (nombre.startswith("xl/drawings/drawing")
                        and nombre.endswith(".xml")):
                    continue
                contenido = zin.read(nombre).decode("utf-8")

                def _reemplazar(match):
                    valor = int(match.group(1))
                    if valor >= fila_corte_xml:
                        return f"<xdr:row>{valor + extras}</xdr:row>"
                    return match.group(0)

                nuevo = _re.sub(
                    r"<xdr:row>(\d+)</xdr:row>",
                    _reemplazar, contenido)
                if nuevo != contenido:
                    cambios[nombre] = nuevo.encode("utf-8")
        finally:
            zin.close()

        if not cambios:
            return archivo_bytes

        salida = _BIO()
        zin = zipfile.ZipFile(_BIO(archivo_bytes), "r")
        zout = zipfile.ZipFile(salida, "w", zipfile.ZIP_DEFLATED)
        try:
            for item in zin.infolist():
                if item.filename in cambios:
                    zout.writestr(item, cambios[item.filename])
                else:
                    zout.writestr(item, zin.read(item.filename))
        finally:
            zin.close()
            zout.close()
        return salida.getvalue()
    except Exception:
        return archivo_bytes


def _arreglar_content_types_imagenes(archivo_bytes):
    """
    Asegura que [Content_Types].xml declare el ContentType para cada
    extension de imagen presente en xl/media/. Sin esta declaracion,
    Excel se queja al abrir el archivo (ventana "Hemos encontrado un
    problema con el contenido de...").
    """
    try:
        import zipfile
        import re as _re
        from io import BytesIO as _BIO

        zin = zipfile.ZipFile(_BIO(archivo_bytes), "r")
        try:
            nombres = zin.namelist()
            ct_xml = zin.read("[Content_Types].xml").decode("utf-8")

            # Extensiones de imagenes en el ZIP
            extensiones = set()
            for n in nombres:
                if n.startswith("xl/media/") and "." in n:
                    ext = n.rsplit(".", 1)[1].lower()
                    extensiones.add(ext)

            if not extensiones:
                return archivo_bytes

            # Mapa de extension -> content type
            mimes = {
                "jpeg": "image/jpeg",
                "jpg":  "image/jpeg",
                "png":  "image/png",
                "gif":  "image/gif",
                "bmp":  "image/bmp",
                "tiff": "image/tiff",
                "tif":  "image/tiff",
                "emf":  "image/x-emf",
                "wmf":  "image/x-wmf",
            }

            extras = ""
            for ext in extensiones:
                if f'Extension="{ext}"' not in ct_xml and ext in mimes:
                    extras += (f'<Default Extension="{ext}" '
                               f'ContentType="{mimes[ext]}"/>')

            if not extras:
                return archivo_bytes

            ct_nuevo = ct_xml.replace("</Types>", extras + "</Types>")
        finally:
            zin.close()

        # Reescribir el ZIP con el Content_Types corregido
        salida = _BIO()
        zin = zipfile.ZipFile(_BIO(archivo_bytes), "r")
        zout = zipfile.ZipFile(salida, "w", zipfile.ZIP_DEFLATED)
        try:
            for item in zin.infolist():
                if item.filename == "[Content_Types].xml":
                    zout.writestr(item, ct_nuevo.encode("utf-8"))
                else:
                    zout.writestr(item, zin.read(item.filename))
        finally:
            zin.close()
            zout.close()
        return salida.getvalue()
    except Exception:
        return archivo_bytes


def _preparar_celda_observaciones(ws, fila_inicio=52, fila_fin=57):
    """
    Las observaciones ocupan A{fila_inicio}:C{fila_fin}, pero vienen
    con cada fila mergeada por separado (A52:C52, A53:C53, ...). Para
    que un texto largo con wrap_text se vea bien en una sola celda
    alta, des-mergeamos las N filas y las re-mergeamos como una sola
    gran celda.
    """
    try:
        a_quitar = [f"A{r}:C{r}" for r in range(fila_inicio, fila_fin + 1)]
        for rango in a_quitar:
            if rango in {str(m) for m in ws.merged_cells.ranges}:
                ws.unmerge_cells(rango)
        ws.merge_cells(f"A{fila_inicio}:C{fila_fin}")
    except Exception:
        pass


def _ampliar_filas_piezas(ws, fila_fin_base, extras):
    """
    Inserta `extras` filas nuevas para piezas DESPUES de la ultima fila
    de piezas original (fila_fin_base, por defecto 50). Las filas de
    OBSERVACIONES, mano de obra y totales se desplazan hacia abajo.

    IMPORTANTE: openpyxl `insert_rows()` tiene un bug que NO mueve los
    merges que empiezan en la fila donde se inserta. Para evitarlo,
    primero des-mergeamos todo lo de abajo, insertamos, y los
    re-mergeamos en su nueva posicion.

    Despues de insertar:
      - Replica los merges B:F y G:H para las filas nuevas.
      - Copia estilos (bordes, fuente, alineacion) de la fila modelo.
      - Actualiza las formulas de los totales con los nuevos rangos.
    """
    try:
        from copy import copy as _copy

        FILA_MODELO = 18
        fila_insert = fila_fin_base + 1  # 51

        # 1) Guardar y des-mergear todos los merges que empiezan en
        #    `fila_insert` o despues. Los vamos a recrear desplazados.
        merges_a_desplazar = []
        for m in list(ws.merged_cells.ranges):
            if m.min_row >= fila_insert:
                merges_a_desplazar.append({
                    "min_r": m.min_row, "max_r": m.max_row,
                    "min_c": m.min_col, "max_c": m.max_col,
                })
                ws.unmerge_cells(str(m))

        # 2) Insertar las filas vacias.
        ws.insert_rows(fila_insert, extras)

        # 3) Re-mergear los merges desplazando sus filas.
        for m in merges_a_desplazar:
            try:
                ws.merge_cells(start_row=m["min_r"] + extras,
                               end_row=m["max_r"] + extras,
                               start_column=m["min_c"],
                               end_column=m["max_c"])
            except Exception:
                pass

        # 4) Replicar merges B:F y G:H en las filas nuevas de piezas.
        for i in range(extras):
            f = fila_insert + i
            try:
                ws.merge_cells(start_row=f, start_column=2,
                               end_row=f, end_column=6)   # B:F pieza
                ws.merge_cells(start_row=f, start_column=7,
                               end_row=f, end_column=8)   # G:H precio
            except Exception:
                pass
            # Copiar altura y estilos de la fila modelo.
            try:
                h = ws.row_dimensions[FILA_MODELO].height
                if h:
                    ws.row_dimensions[f].height = h
            except Exception:
                pass
            for col in range(1, 9):
                try:
                    src = ws.cell(row=FILA_MODELO, column=col)
                    dst = ws.cell(row=f, column=col)
                    if src.has_style:
                        dst.font = _copy(src.font)
                        dst.border = _copy(src.border)
                        dst.fill = _copy(src.fill)
                        dst.alignment = _copy(src.alignment)
                        dst.number_format = src.number_format
                        dst.protection = _copy(src.protection)
                except Exception:
                    pass

        # 5) Actualizar formulas de totales con los nuevos indices.
        f_mo_total = 58 + extras       # era 58 (=SUM(G53:H57))
        f_repuestos = 59 + extras      # era 59 (=SUM(G18:H50))
        f_franq = 60 + extras          # era 60 (=+H11)
        f_neto = 61 + extras           # era 61 (=+G58+G59-G60)
        f_mo_ini = 53 + extras         # era 53
        f_mo_fin = 57 + extras         # era 57
        f_pz_fin = fila_fin_base + extras  # era 50

        _escribir(ws, f_mo_total, 7, f"=SUM(G{f_mo_ini}:H{f_mo_fin})")
        _escribir(ws, f_repuestos, 7, f"=SUM(G18:H{f_pz_fin})")
        _escribir(ws, f_franq, 7, "=+H11")
        _escribir(ws, f_neto, 7, f"=+G{f_mo_total}+G{f_repuestos}-G{f_franq}")

        # 6) openpyxl insert_rows no actualiza las referencias de
        #    formulas. Las formulas G=E*F en las filas de MO (ex 53-56)
        #    siguen apuntando a E53*F53 pero deberian apuntar a la
        #    nueva fila. Las re-escribimos para cada fila MO menos la
        #    de "Varios" (que no tiene formula E*F).
        for i in range(4):  # Pintura, Chapa, Mecanica, Tapiceria
            f = f_mo_ini + i
            _escribir(ws, f, 7, f"=+E{f}*F{f}")
    except Exception:
        pass


# ============================================================
#  FUNCION PRINCIPAL
# ============================================================

def _enriquecer_desde_excel(data, texto_peritacion):
    """
    Post-procesamiento sin llamadas a Claude: revisa el texto del
    peritaje en busca de montos que Claude pudo no haber capturado bien,
    y los completa SOLO si el campo correspondiente esta vacio o en 0.

    Hoy busca dos cosas en el bloque del Excel:
      - Total de repuestos (fila "REPUESTOS" con un monto).
      - Items extra de mano de obra (cristaleria/vidrieria/varios/etc)
        que se suman al campo manoObra.varios como monto en pesos.

    No depende del prompt y no agrega memoria ni llamadas.
    """
    if not texto_peritacion:
        return data

    import re

    # Tomar solo el bloque del Excel (no el texto libre pegado, para no
    # confundir con datos sueltos del peritaje en prosa).
    marca = "=== PERITAJE DESDE EXCEL (GRILLA) ==="
    if marca not in texto_peritacion:
        return data
    bloque = texto_peritacion.split(marca, 1)[1]

    def _a_numero(txt):
        # Formatos AR soportados:
        #   "$ 1.708.000,50" -> 1708000.5
        #   "$ 360.000"       -> 360000
        #   "$ 51.000"        -> 51000
        #   "100000"          -> 100000
        #   "100,50"          -> 100.5
        s = re.sub(r"[^\d,.\-]", "", str(txt))
        if not s:
            return 0
        # Si hay coma: es decimal (formato AR). Los puntos son miles.
        if "," in s:
            s = s.replace(".", "").replace(",", ".")
        # Sin coma: si hay 2 o mas puntos, son miles.
        # Si hay 1 punto seguido de exactamente 3 digitos, tambien es miles.
        elif s.count(".") >= 2:
            s = s.replace(".", "")
        elif s.count(".") == 1:
            entera, fraccion = s.split(".")
            if len(fraccion) == 3:
                s = entera + fraccion
        try:
            return float(s)
        except ValueError:
            return 0

    # Listar todas las celdas como pares (texto, monto) en la misma fila.
    # El formato del bloque es: "A1=valor | B1=valor | ..." por linea.
    montos_por_fila = []  # lista de (etiquetas_upper, lista_de_montos)
    for linea in bloque.splitlines():
        if not linea.strip() or linea.startswith("---"):
            continue
        etiquetas = []
        montos = []
        for celda in linea.split(" | "):
            if "=" not in celda:
                continue
            val = celda.split("=", 1)[1].strip()
            if not val:
                continue
            # Heuristica: si tiene digitos y simbolo $ o formato de
            # miles, lo consideramos monto; si no, etiqueta.
            tiene_signo = "$" in val
            tiene_miles = bool(re.search(r"\d{1,3}([.,]\d{3})+", val))
            es_solo_num = bool(re.fullmatch(r"-?\d+([.,]\d+)?", val))
            if tiene_signo or tiene_miles or (es_solo_num and float(_a_numero(val)) >= 1000):
                montos.append(_a_numero(val))
            else:
                etiquetas.append(val.upper())
        if etiquetas and montos:
            montos_por_fila.append((etiquetas, montos))

    # --- 1. Total de repuestos ---
    # Si la suma de precios de piezas a CAMBIAR es 0 y aparece una fila
    # cuya etiqueta dice "REPUESTOS", tomar el monto mayor de esa fila.
    total_piezas = sum((d.get("precio") or 0) for d in data.get("danos", [])
                       if d.get("accion") == "CAMBIAR")
    if total_piezas == 0:
        for etiquetas, montos in montos_por_fila:
            if any("REPUESTO" in e for e in etiquetas):
                data["_totalRepuestosExcel"] = max(montos)
                break

    # --- 2. Mano de obra "varios" (cristaleria, gas, service, etc) ---
    # Si manoObra.varios viene en 0 o como cantidad pequena (<1000),
    # buscar filas con etiquetas conocidas que no sean los 4 conceptos
    # estandar (pintura/chapa/mecanica/tapiceria) y sumar sus subtotales.
    conceptos_estandar = ("PINTURA", "CHAPA", "MECANICA", "TAPICERIA")
    palabras_varios = ("CRISTAL", "VIDRIO", "VIDRIER", "GAS", "SERVICE",
                       "VARIOS", "ALINEACI", "ELECTRIC")
    varios_actual = float(data.get("manoObra", {}).get("varios", 0) or 0)
    if varios_actual < 1000:  # Si es cantidad o esta vacio, recalcular.
        suma = 0.0
        for etiquetas, montos in montos_por_fila:
            # Saltar filas que son de los 4 conceptos estandar.
            if any(c in e for e in etiquetas for c in conceptos_estandar):
                continue
            # Saltar filas de repuestos (ya capturadas arriba).
            if any("REPUESTO" in e for e in etiquetas):
                continue
            # Saltar filas de totales generales / franquicia.
            if any(p in e for e in etiquetas
                   for p in ("TOTAL", "FRANQUICIA", "NETO", "RESUMEN",
                             "IMPREVIS", "INSPECTOR", "RESPONSABLE")):
                continue
            # Aceptar solo si tiene una palabra reconocida como "varios".
            if any(p in e for e in etiquetas for p in palabras_varios):
                # Sumar el SUBTOTAL (monto mayor) de la fila.
                suma += max(montos)
        if suma > 0:
            data["manoObra"]["varios"] = suma

    return data


def procesar(plantilla_bytes, texto_peritacion="", excel_peritacion_bytes=None,
             excel_peritacion_nombre="", cotizar=False, overrides=None):
    """
    Funcion principal: recibe la plantilla y la peritacion,
    devuelve (informe_bytes, data, detalle_cotizacion).

    - texto_peritacion: texto libre de la peritacion (lo entiende Claude).
    - excel_peritacion_bytes: bytes del Excel del peritaje, si fue
      subido. Si está presente, se parsea LOCALMENTE en Python para
      extraer piezas y precios, sin mandar la grilla a Claude (es
      mucho más rápido y no inventa piezas).
    - excel_peritacion_nombre: nombre original del archivo Excel
      (sirve para detectar .xls vs .xlsx).
    - cotizar: parametro mantenido por compatibilidad. La cotizacion
      automatica fue descartada; los precios se cargan manualmente.
    - overrides: dict opcional con campos cargados a mano (ej:
      fechaInspeccion, lugar de inspeccion). Pisan lo inferido por Claude.

    detalle_cotizacion siempre se devuelve como lista vacia.
    """
    # 1) Si vino un Excel, parsear localmente las piezas y mano de obra.
    #    Esto es rapido y confiable (sin invenciones).
    piezas_del_excel = None
    manoObra_del_excel = None
    franquicia_del_excel = None
    observaciones_del_excel = None
    if excel_peritacion_bytes:
        try:
            from parser_peritaje import parsear_peritaje_excel
            resultado = parsear_peritaje_excel(
                excel_peritacion_bytes,
                excel_peritacion_nombre or "x.xlsx")
            if resultado:
                if resultado.get("danos"):
                    piezas_del_excel = resultado["danos"]
                if resultado.get("manoObra_items"):
                    manoObra_del_excel = resultado["manoObra_items"]
                if resultado.get("franquicia"):
                    franquicia_del_excel = resultado["franquicia"]
                if resultado.get("observaciones"):
                    observaciones_del_excel = resultado["observaciones"]
        except Exception:
            pass

    # 2) Si hay piezas del Excel, agregar al texto una lista para que
    #    Claude las use como base y les pueda asignar precios desde la
    #    cotizacion pegada. Sin esto Claude "no ve" las piezas del
    #    peritaje (porque ya no le mandamos la grilla) y no puede
    #    matchear precios.
    texto_para_claude = texto_peritacion or ""
    if piezas_del_excel:
        lineas = ["", "=== PIEZAS DEL PERITAJE (extraidas del Excel) ==="]
        lineas.append("Estas son las piezas del peritaje. USA ESTAS piezas")
        lineas.append("como campo \"danos\" en el JSON, en el mismo orden.")
        lineas.append("Si hay una cotizacion pegada, asignale a cada pieza")
        lineas.append("el precio correspondiente (matcheando por nombre,")
        lineas.append("tolerando abreviaciones y variantes).")
        lineas.append("")
        for i, p in enumerate(piezas_del_excel, 1):
            lineas.append(f"  {i}. [{p['accion']}] {p['pieza']}")
        texto_para_claude = (texto_para_claude + "\n" + "\n".join(lineas)).strip()

    # 3) El texto libre (mas la lista de piezas) va a Claude.
    data = (parsear_texto(texto_para_claude) if texto_para_claude
            else data_vacia())

    # 4) Combinar: si Claude devolvio danos con precios != 0, usarlos.
    #    Si no, caer a las piezas del parser local (sin precios).
    if piezas_del_excel:
        danos_claude = data.get("danos") or []
        # Usar los de Claude si (a) coinciden en cantidad con los del
        # parser (para verificar que respeto la lista) y (b) al menos
        # una tiene precio distinto de 0.
        usar_claude = (
            len(danos_claude) == len(piezas_del_excel)
            and any((d.get("precio") or 0) > 0 for d in danos_claude)
        )
        if usar_claude:
            # Extra: forzar los nombres del parser para evitar que Claude
            # los haya reescrito. Solo tomamos el precio de Claude.
            for i, p in enumerate(piezas_del_excel):
                p["precio"] = danos_claude[i].get("precio", 0) or 0
        data["danos"] = piezas_del_excel

    # 5) Mano de obra y franquicia del Excel: en modo flexible, guardar
    #    los items tal cual vinieron (respetando modificaciones del
    #    perito como "MECANICA Y TAPICERIA" combinado).
    if manoObra_del_excel:
        data["manoObra_items"] = manoObra_del_excel
    if franquicia_del_excel:
        data["franquicia"] = franquicia_del_excel

    # 6) Observaciones del Excel: si el peritaje trae observaciones
    #    (bloque despues del header "OBSERVACIONES"), se combinan con
    #    las que haya devuelto Claude del texto pegado, y despues se
    #    pasan por la reescritura tecnica.
    if observaciones_del_excel:
        obs_previas = (data.get("observaciones") or "").strip()
        if obs_previas:
            texto_obs = obs_previas + ". " + observaciones_del_excel
        else:
            texto_obs = observaciones_del_excel
        try:
            data["observaciones"] = _mejorar_observaciones(texto_obs)
        except Exception:
            data["observaciones"] = texto_obs

    # Los datos cargados a mano tienen prioridad sobre lo inferido.
    if overrides:
        for clave, valor in overrides.items():
            if valor not in (None, ""):
                data[clave] = valor

    informe_bytes = completar_plantilla(plantilla_bytes, data)
    return informe_bytes, data, []
