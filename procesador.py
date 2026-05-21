"""
procesador.py
Motor de procesamiento de informes de tasacion.
- Parsea el texto/Excel de la peritacion.
- Completa la plantilla virgen preservando logos, estilos y formulas.
"""

import re
import time
import zipfile
import unicodedata
import statistics
from io import BytesIO
from urllib.parse import quote
from urllib.request import urlopen, Request
import json

from openpyxl import load_workbook


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
    """Convierte un string a numero, quitando $, puntos y comas."""
    if valor is None:
        return 0
    s = re.sub(r"[^\d]", "", str(valor))
    return int(s) if s else 0


def normalizar_accion(raw):
    """Normaliza la accion a CAMBIAR o REPARAR. 'Pintar' -> 'REPARAR'."""
    a = normalizar(raw)
    if a in ("CAMBIAR", "CAMBIO", "SUSTITUIR", "C"):
        return "CAMBIAR"
    if a in ("REPARAR", "REPARACION", "PINTAR", "PINTURA", "R", "P"):
        return "REPARAR"
    return a


# ============================================================
#  ESTRUCTURA DE DATOS
# ============================================================

def data_vacia():
    return {
        "numeroSiniestro": "", "fechaSiniestro": "", "fechaInspeccion": "",
        "asegurado": "", "marca": "", "modelo": "", "anio": "", "dominio": "",
        "chasis": "", "kilometraje": "", "sumaAsegurada": "", "franquiciaVeh": "",
        "tallerNombre": "", "tallerDireccion": "", "tallerLocalidad": "",
        "danos": [],  # lista de {accion, pieza, precio}
        "manoObra": {
            "pintura": 0, "chapa": 0, "mecanica": 0, "tapiceria": 0, "varios": 0,
            "pinturaValor": 120000, "chapaValor": 120000,
            "mecanicaValor": 50000, "tapiceriaValor": 50000,
        },
        "franquicia": 0,
        "observaciones": "",
    }


# ============================================================
#  PARSER DE TEXTO LIBRE
# ============================================================

# Etiquetas reconocidas: (variantes, campo)
LABEL_DEFS = [
    (["NRO STRO", "NRO SINIESTRO", "NUMERO DE SINIESTRO", "N STRO", "N SINIESTRO"], "numeroSiniestro"),
    (["FECHA STRO", "FECHA DE SINIESTRO", "FECHA SINIESTRO", "F STRO", "F SINIESTRO"], "fechaSiniestro"),
    (["FECHA INSPECCION", "FECHA DE INSPECCION", "FECHA IP", "F IP"], "fechaInspeccion"),
    (["ASEGURADO", "APELLIDO Y NOMBRE"], "asegurado"),
    (["VEHICULO"], "__vehiculo"),
    (["MARCA"], "marca"),
    (["MODELO"], "modelo"),
    (["ANO", "ANIO"], "anio"),
    (["PATENTE", "DOMINIO"], "dominio"),
    (["CHASIS", "N CHASIS", "NRO CHASIS"], "chasis"),
    (["KILOMETRAJE", "KM"], "kilometraje"),
    (["SUMA ASEG", "SUMA ASEGURADA", "S ASEG"], "sumaAsegurada"),
    (["FRANQUICIA A DEDUCIR", "FRANQUICIA"], "franquicia"),
]

# Etiquetas que se ignoran pero sirven de delimitador.
# Todas estas lineas NO deben entrar como observaciones.
IGNORE_LABELS = [
    "FECHA CARGA", "SELLO", "POLIZA", "PROP", "STROS", "PRODUCTOR", "DOM.PAS",
    "CP", "CATEGORIA IVA", "AG. RET CUIT", "CUIT", "NRO JUB", "TIPO JUBILACION",
    "ING BRUTOS", "NRO IB", "JURISDICCION PAS", "SER. SOCIAL", "DOMICILIO",
    "ITEM", "USO", "MOTOR", "DIRECCION", "DIR", "LOCALIDAD",
]


def _construir_patron_labels():
    todas = []
    for variantes, _ in LABEL_DEFS:
        todas.extend(variantes)
    todas.extend(IGNORE_LABELS)
    # Ordenar por longitud descendente para que las mas largas matcheen primero
    todas.sort(key=len, reverse=True)
    return "|".join(re.escape(x) for x in todas)


def _limpiar_valor(valor):
    """
    Limpia un valor de campo: quita marcadores // | y palabras-seccion
    pegadas. NO corta la coma entre digitos (ej. '1,8' es decimal).
    Ej: '8900000 // SUSTITUIR:' -> '8900000'
    """
    if not valor:
        return ""
    valor = re.split(
        r"\s*(?://|\|)\s*|"
        r"\s*,\s+|"
        r"\s+(?=(?:SUSTITUIR|CAMBIAR|PINTAR|REPARAR|REEMPLAZAR|"
        r"OBSERVACIONES?|MANO DE OBRA)\s*:?\s*$)",
        valor, maxsplit=1, flags=re.IGNORECASE)[0]
    return valor.strip()


def parsear_texto(texto):
    """Parsea el texto libre de la peritacion y devuelve el dict de datos."""
    data = data_vacia()
    if not texto or not texto.strip():
        return data

    labels_pattern = _construir_patron_labels()
    lineas = texto.split("\n")
    lineas_norm = [normalizar(l) for l in lineas]

    # Guarda que lineas tienen una etiqueta de cabecera (para excluirlas
    # de las observaciones implicitas mas adelante).
    lineas_con_campo = set()

    # ---- PASO 1: extraer campos de cabecera ----
    for idx, linea_norm in enumerate(lineas_norm):
        if not linea_norm.strip():
            continue
        linea_orig = lineas[idx]

        patron = (r"(" + labels_pattern + r")\s*[:=]\s*(.*?)"
                  r"(?=\s+(?:" + labels_pattern + r")\s*[:=]|$)")
        for m in re.finditer(patron, linea_norm, re.IGNORECASE):
            lineas_con_campo.add(idx)
            label = m.group(1).strip()
            valor_raw = m.group(2)
            inicio = m.start(2)
            valor = linea_orig[inicio:inicio + len(valor_raw)].strip()
            valor = _limpiar_valor(valor)
            if not valor:
                continue

            # Buscar el campo correspondiente
            campo = None
            for variantes, f in LABEL_DEFS:
                if any(normalizar(v) == normalizar(label) for v in variantes):
                    campo = f
                    break
            if campo is None:
                continue

            if campo == "__vehiculo":
                # VEHICULO: CHEVROLET MERIVA GLS 1,8 -> marca + modelo
                if not data["marca"] and not data["modelo"]:
                    partes = valor.split()
                    data["marca"] = partes[0] if partes else ""
                    data["modelo"] = " ".join(partes[1:])
            elif campo == "franquicia":
                if not data["franquicia"]:
                    data["franquicia"] = a_numero(valor)
            else:
                if not data[campo]:
                    data[campo] = valor

    # ---- PASO 1b: nro de siniestro sin etiqueta ----
    # En algunas peritaciones el nro de siniestro aparece solo, como numero
    # suelto en una de las primeras lineas (ej. una linea que dice "6").
    if not data["numeroSiniestro"]:
        for linea in lineas[:5]:
            l = linea.strip()
            if re.fullmatch(r"\d{1,8}", l):
                data["numeroSiniestro"] = l
                break

    # ---- PASO 2: secciones de danos ----
    # Las palabras SUSTITUIR:/PINTAR: pueden estar en su propia linea
    # O al final de otra linea (ej: "SUMA ASEG: 8900000 // SUSTITUIR:").
    sustituir_idx = -1
    pintar_idx = -1
    for i, ln in enumerate(lineas_norm):
        ln = ln.strip()
        if sustituir_idx < 0 and re.search(
                r"(?:^|[\s/|:])(SUSTITUIR|REEMPLAZAR)\s*:?\s*$", ln):
            sustituir_idx = i
        if pintar_idx < 0 and re.search(
                r"(?:^|[\s/|:])(PINTAR)\s*:?\s*$", ln):
            pintar_idx = i

    def es_mano_obra(linea):
        return re.match(
            r"^\d+(?:[.,]\d+)?\s*(?:HS\s+|HORAS?\s+)?"
            r"(PA[NN]OS?|CHAPA|MECANICA|TAPICERIA|PINTURA)",
            normalizar(linea),
        )

    def es_varios(linea):
        return re.match(
            r"^(CARGA DE GAS|VARIOS|OTROS|SERVICE|ADICIONAL)\s*\$?\s*[\d.,]+",
            normalizar(linea),
        )

    def es_campo(linea):
        return re.search(
            r"(NRO STRO|FECHA STRO|FECHA DE|FECHA IP|PATENTE|VEHICULO|CHASIS|"
            r"SUMA ASEG|ASEGURADO|MARCA|MODELO|ANO|KILOMETRAJE|DOMINIO|TALLER|"
            r"FRANQUICIA|LOCALIDAD|DIRECCION|POLIZA|PRODUCTOR|CUIT|ITEM|USO|"
            r"MOTOR|DOMICILIO|CP|CATEGORIA IVA|NRO JUB|TIPO JUBILACION|"
            r"ING BRUTOS|NRO IB|JURISDICCION|SER. SOCIAL|PROP|STROS|"
            r"DOM.PAS|AG. RET)\s*[:=]",
            normalizar(linea),
        )

    def es_seccion(linea):
        ln = normalizar(linea).strip()
        return re.search(
            r"(?:^|[\s/|:])(SUSTITUIR|CAMBIAR|PINTAR|REPARAR|OBSERVACIONES|OBS|"
            r"MANO DE OBRA|COTIZACION|REEMPLAZAR)\s*:?\s*$",
            ln,
        )

    def procesar_bloque_danos(inicio, accion):
        """Lee piezas desde 'inicio' hasta que aparece otra seccion/campo/MO."""
        for i in range(inicio + 1, len(lineas)):
            linea = lineas[i].strip()
            if not linea:
                continue
            if es_seccion(linea):
                break
            if es_mano_obra(linea) or es_varios(linea):
                break
            if es_campo(linea):
                break
            # Es una pieza. Puede tener precio: "Pieza - 12345" o "Pieza $12345"
            m_precio = re.match(r"^(.+?)\s*[-|$]\s*\$?\s*([\d.,]+)\s*$", linea)
            if m_precio:
                pieza = m_precio.group(1).strip()
                precio = a_numero(m_precio.group(2))
            else:
                pieza = linea
                precio = 0
            data["danos"].append({
                "accion": accion,
                "pieza": pieza,
                "precio": precio if accion == "CAMBIAR" else 0,
            })

    if sustituir_idx >= 0:
        procesar_bloque_danos(sustituir_idx, "CAMBIAR")
    if pintar_idx >= 0:
        procesar_bloque_danos(pintar_idx, "REPARAR")

    # ---- PASO 3: mano de obra ----
    for linea in lineas:
        l = linea.strip()
        if not l:
            continue
        m = re.match(
            r"^(\d+(?:[.,]\d+)?)\s*(?:HS\s+|HORAS?\s+)?"
            r"(PA[NN]OS?|CHAPA|MECANICA|TAPICERIA|PINTURA)"
            r"(?:\s*X\s*\$?\s*([\d.,]+))?",
            normalizar(l),
        )
        if m:
            cant = a_numero(m.group(1))
            concepto = m.group(2)
            valor = a_numero(m.group(3)) if m.group(3) else 0
            if "PAN" in concepto or concepto == "PINTURA":
                data["manoObra"]["pintura"] = cant
                if valor:
                    data["manoObra"]["pinturaValor"] = valor
            elif concepto == "CHAPA":
                data["manoObra"]["chapa"] = cant
                if valor:
                    data["manoObra"]["chapaValor"] = valor
            elif concepto.startswith("MEC"):
                data["manoObra"]["mecanica"] = cant
                if valor:
                    data["manoObra"]["mecanicaValor"] = valor
            elif concepto.startswith("TAPIC"):
                data["manoObra"]["tapiceria"] = cant
                if valor:
                    data["manoObra"]["tapiceriaValor"] = valor
            continue
        # Varios / carga de gas
        m_v = re.match(
            r"^(CARGA DE GAS|VARIOS|OTROS|SERVICE|ADICIONAL)\s*\$?\s*([\d.,]+)",
            normalizar(l),
        )
        if m_v:
            data["manoObra"]["varios"] += a_numero(m_v.group(2))

    # ---- PASO 4: observaciones ----
    # Solo captura el bloque que sigue a la etiqueta "OBSERVACIONES".
    # Si no hay etiqueta explicita, toma lineas largas en prosa que NO
    # contengan ninguna etiqueta de campo (para no meter datos de poliza).
    obs = []
    capturando = False
    indices_danos = set()  # lineas que son piezas, para excluirlas

    # Marcar el rango de lineas de los bloques de danos
    for marca_idx in (sustituir_idx, pintar_idx):
        if marca_idx >= 0:
            for i in range(marca_idx, len(lineas)):
                ln = lineas[i].strip()
                if i != marca_idx and (es_seccion(ln) or es_campo(ln)
                                       or es_mano_obra(ln) or es_varios(ln)):
                    break
                indices_danos.add(i)

    for idx, linea in enumerate(lineas):
        l = linea.strip()
        if not l:
            continue
        ln = normalizar(l)
        if re.match(r"^(OBSERVACIONES?|OBS)\s*:?\s*$", ln):
            capturando = True
            continue
        if capturando:
            if es_seccion(l) or es_campo(l):
                capturando = False
            else:
                obs.append(_limpiar_valor(l))
                continue
        # Observaciones implicitas: lineas largas en prosa, sin etiquetas,
        # que no sean piezas ni mano de obra ni esten en lineas con campo.
        if (len(l) > 60
                and idx not in lineas_con_campo
                and idx not in indices_danos
                and not es_campo(l)
                and not es_mano_obra(l)
                and not es_varios(l)
                and not re.match(r"^\d+\s", l)):
            obs.append(_limpiar_valor(l))
    data["observaciones"] = "\n".join(o for o in obs if o)

    return data


# ============================================================
#  PARSER DE EXCEL DE PERITACION
# ============================================================

def parsear_excel_peritacion(file_bytes):
    """Lee un Excel de peritacion y extrae los datos."""
    data = data_vacia()
    wb = load_workbook(BytesIO(file_bytes), data_only=True)

    ws1 = wb.worksheets[0]
    filas = list(ws1.iter_rows(values_only=True))
    for fila in filas:
        for i, celda in enumerate(fila):
            etiqueta = normalizar(celda)
            siguiente = ""
            for j in range(i + 1, len(fila)):
                if fila[j] not in (None, ""):
                    siguiente = str(fila[j])
                    break
            if "NUMERO DE SINIESTRO" in etiqueta or "NRO SINIESTRO" in etiqueta:
                data["numeroSiniestro"] = siguiente
            elif "FECHA DE SINIESTRO" in etiqueta:
                data["fechaSiniestro"] = siguiente
            elif "FECHA DE INSPECC" in etiqueta:
                data["fechaInspeccion"] = siguiente
            elif "APELLIDO Y NOMBRE" in etiqueta:
                data["asegurado"] = siguiente
            elif etiqueta in ("MARCA:", "MARCA"):
                data["marca"] = siguiente
            elif etiqueta in ("MODELO:", "MODELO"):
                data["modelo"] = siguiente
            elif etiqueta in ("ANO:", "ANO"):
                data["anio"] = siguiente
            elif etiqueta in ("DOMINIO:", "DOMINIO", "PATENTE:", "PATENTE"):
                data["dominio"] = siguiente
            elif "CHASIS" in etiqueta:
                data["chasis"] = siguiente
            elif "KILOMETRA" in etiqueta:
                data["kilometraje"] = siguiente
            elif "SUMA ASEG" in etiqueta:
                data["sumaAsegurada"] = siguiente

    if len(wb.worksheets) > 1:
        ws2 = wb.worksheets[1]
        leyendo = False
        for fila in ws2.iter_rows(values_only=True):
            primera = normalizar(fila[0]) if fila and len(fila) > 0 else ""
            segunda = normalizar(fila[1]) if fila and len(fila) > 1 else ""
            if primera in ("ACCION", "ACCION:"):
                leyendo = True
                continue
            if "TOTAL" in primera:
                leyendo = False
                continue
            if leyendo:
                accion = pieza = ""
                precio = 0
                if primera == "X":
                    accion = "CAMBIAR"
                    pieza = str(fila[1] or (fila[2] if len(fila) > 2 else "") or "").strip()
                    precio = a_numero(fila[-1])
                elif segunda == "X":
                    accion = "REPARAR"
                    pieza = str(fila[2] if len(fila) > 2 else "" or "").strip()
                elif primera in ("CAMBIAR", "REPARAR", "PINTAR"):
                    accion = normalizar_accion(primera)
                    pieza = str(fila[1] or "").strip()
                    precio = a_numero(fila[-1]) if accion == "CAMBIAR" else 0
                if accion and pieza:
                    data["danos"].append({"accion": accion, "pieza": pieza, "precio": precio})

    return data


def combinar_datos(data_excel, data_texto):
    """Combina datos de Excel + texto. El texto pisa al Excel donde tenga valor."""
    combinado = dict(data_excel)
    campos = ["numeroSiniestro", "fechaSiniestro", "fechaInspeccion", "asegurado",
              "marca", "modelo", "anio", "dominio", "chasis", "kilometraje",
              "sumaAsegurada", "franquiciaVeh", "tallerNombre", "tallerDireccion",
              "tallerLocalidad", "observaciones"]
    for c in campos:
        if data_texto.get(c):
            combinado[c] = data_texto[c]
    if data_texto.get("danos"):
        combinado["danos"] = list(data_excel.get("danos", [])) + data_texto["danos"]
    for k, v in data_texto.get("manoObra", {}).items():
        if v:
            combinado["manoObra"][k] = v
    if data_texto.get("franquicia"):
        combinado["franquicia"] = data_texto["franquicia"]
    return combinado


# ============================================================
#  PRESERVACION DE IMAGENES (LOGOS) VIA ZIP
# ============================================================

def _reinyectar_imagenes(plantilla_bytes, generado_bytes):
    """
    Copia las imagenes (logos de encabezado y pie) de la plantilla original
    al Excel generado, a nivel ZIP.

    Esto es a prueba de fallos: aunque openpyxl pierda o altere las
    imagenes/drawings al guardar, esta funcion restaura los archivos
    originales exactos (xl/media/* y xl/drawings/*).

    Si la plantilla no tiene imagenes, devuelve el generado sin cambios.
    """
    try:
        zin_plant = zipfile.ZipFile(BytesIO(plantilla_bytes))
        zin_gen = zipfile.ZipFile(BytesIO(generado_bytes))
    except zipfile.BadZipFile:
        return generado_bytes

    # Archivos de imagen / drawing de la plantilla original
    items_img = [n for n in zin_plant.namelist()
                 if n.startswith("xl/media/")
                 or n.startswith("xl/drawings/")]

    # Si la plantilla no tiene imagenes, no hay nada que reinyectar
    if not any(n.startswith("xl/media/") for n in items_img):
        return generado_bytes

    out = BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        # 1. Copiar todo el contenido del generado EXCEPTO media y drawings
        for item in zin_gen.namelist():
            if (item.startswith("xl/media/")
                    or item.startswith("xl/drawings/")):
                continue
            zout.writestr(item, zin_gen.read(item))
        # 2. Copiar media y drawings ORIGINALES de la plantilla
        copiados = set()
        for item in items_img:
            zout.writestr(item, zin_plant.read(item))
            copiados.add(item)
        # 3. Por las dudas, si el generado tenia algun drawing/media extra
        #    que no estaba en la plantilla, no lo copiamos (evita duplicados).

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


def completar_plantilla(plantilla_bytes, data):
    """
    Carga la plantilla virgen, la completa con los datos y devuelve los bytes.
    Preserva logos/imagenes, estilos, formulas y formato de impresion.
    Maneja correctamente las celdas combinadas (merged cells).
    """
    wb = load_workbook(BytesIO(plantilla_bytes))
    hojas = wb.worksheets

    # ---------- HOJA 1: Datos preliminares ----------
    ws1 = hojas[0]

    mapeo_hoja1 = [
        (["NUMERO DE SINIESTRO", "NRO SINIESTRO"], data["numeroSiniestro"]),
        (["FECHA DE SINIESTRO", "FECHA SINIESTRO"], data["fechaSiniestro"]),
        (["FECHA DE INSPECCION", "FECHA INSPECCION"], data["fechaInspeccion"]),
        (["APELLIDO Y NOMBRE"], data["asegurado"]),
        (["MARCA"], data["marca"]),
        (["MODELO"], data["modelo"]),
        (["ANO", "ANIO"], data["anio"]),
        (["DOMINIO", "PATENTE"], data["dominio"]),
        (["CHASIS"], data["chasis"]),
        (["KILOMETRAJE", "KILOMETRA"], data["kilometraje"]),
        (["SUMA ASEGURADA", "SUMA ASEG"], data["sumaAsegurada"]),
    ]
    for etiquetas, valor in mapeo_hoja1:
        if not valor:
            continue
        pos = _buscar_celda_etiqueta(ws1, etiquetas)
        if pos is not None:
            _escribir(ws1, pos[0], pos[1], valor)

    # Franquicia del vehiculo (1ra ocurrencia de "FRANQUICIA").
    # Regla: si no hay valor, se escribe 0 (no se deja vacio).
    franq_veh = a_numero(data.get("franquiciaVeh")) if data.get("franquiciaVeh") else 0
    pos = _buscar_celda_etiqueta(ws1, ["FRANQUICIA"], ocurrencia=1)
    if pos is not None:
        _escribir(ws1, pos[0], pos[1], franq_veh)

    # Franquicia a deducir (2da ocurrencia, "FRANQUICIA:").
    # La Hoja3 la lee desde aqui con una formula. Si no hay valor, va 0.
    franq_ded = a_numero(data.get("franquicia")) if data.get("franquicia") else 0
    pos = _buscar_celda_etiqueta(ws1, ["FRANQUICIA"], ocurrencia=2)
    if pos is not None:
        _escribir(ws1, pos[0], pos[1], franq_ded)

    # ---------- HOJA 2: Descripcion de danos ----------
    if len(hojas) > 1:
        ws2 = hojas[1]
        fila_inicio = None
        col_accion = col_pieza = col_precio = None
        for fila_celdas in ws2.iter_rows():
            for celda in fila_celdas:
                v = normalizar(celda.value)
                if v in ("ACCION", "ACCION:"):
                    fila_inicio = celda.row + 1
                    col_accion = celda.column
                if "PIEZA" in v:
                    col_pieza = celda.column
                if "PRECIO" in v:
                    col_precio = celda.column
            if fila_inicio:
                break

        if fila_inicio:
            col_accion = col_accion or 1
            col_pieza = col_pieza or 2
            col_precio = col_precio or 6
            for i, dano in enumerate(data["danos"]):
                fila = fila_inicio + i
                _escribir(ws2, fila, col_accion, dano["accion"])
                _escribir(ws2, fila, col_pieza, dano["pieza"])
                # Solo las piezas a CAMBIAR llevan precio.
                # Si la pieza es CAMBIAR pero no tiene precio, va 0.
                if dano["accion"] == "CAMBIAR":
                    _escribir(ws2, fila, col_precio, dano.get("precio", 0) or 0)

    # ---------- HOJA 3: Mano de obra, resumen, observaciones ----------
    if len(hojas) > 2:
        ws3 = hojas[2]
        mo = data["manoObra"]

        def _buscar_cols_cant_vu(fila_concepto):
            """Busca las columnas CANT. y V. UNITARIO mirando filas de arriba."""
            col_cant = col_vu = None
            for r in range(max(1, fila_concepto - 6), fila_concepto):
                for c in range(1, 12):
                    hv = normalizar(_valor_celda(ws3, r, c))
                    if "CANT" in hv:
                        col_cant = c
                    if "UNITARIO" in hv or "V. UNIT" in hv:
                        col_vu = c
            return col_cant, col_vu

        def set_mano_obra(nombre_concepto, cantidad, valor_unitario):
            """Busca la fila del concepto y carga cantidad y valor unitario."""
            objetivo = normalizar(nombre_concepto)
            for fila_celdas in ws3.iter_rows():
                for celda in fila_celdas:
                    if normalizar(celda.value) == objetivo:
                        col_cant, col_vu = _buscar_cols_cant_vu(celda.row)
                        if col_cant:
                            _escribir(ws3, celda.row, col_cant, cantidad)
                        if col_vu and valor_unitario:
                            _escribir(ws3, celda.row, col_vu, valor_unitario)
                        return

        set_mano_obra("Pintura", mo["pintura"], mo["pinturaValor"])
        set_mano_obra("Chapa", mo["chapa"], mo["chapaValor"])
        set_mano_obra("Mecanica", mo["mecanica"], mo["mecanicaValor"])
        set_mano_obra("Tapiceria", mo["tapiceria"], mo["tapiceriaValor"])

        # Varios: la fila "Varios" tiene formula SUBTOTAL = CANT * V.UNITARIO.
        # Para no pisar la formula, se carga el monto como V.UNITARIO con CANT=1.
        if mo["varios"]:
            for fila_celdas in ws3.iter_rows():
                for celda in fila_celdas:
                    if normalizar(celda.value) == "VARIOS":
                        col_cant, col_vu = _buscar_cols_cant_vu(celda.row)
                        if col_cant:
                            _escribir(ws3, celda.row, col_cant, 1)
                        if col_vu:
                            _escribir(ws3, celda.row, col_vu, mo["varios"])

        # Franquicia a deducir: NO se escribe si la celda tiene formula
        # (suele ser =+Hoja1!$F$28). El valor real se carga en Hoja1.
        if data.get("franquicia"):
            for fila_celdas in ws3.iter_rows():
                for celda in fila_celdas:
                    if "FRANQUICIA A DEDUCIR" in normalizar(celda.value):
                        for c in range(celda.column + 1, celda.column + 9):
                            val = _valor_celda(ws3, celda.row, c)
                            if isinstance(val, str) and val.startswith("="):
                                continue
                            if val in (None, ""):
                                _escribir(ws3, celda.row, c, data["franquicia"])
                                break

        # Observaciones: la etiqueta "OBSERVACIONES" suele estar en una celda
        # combinada. El texto va en el area (merge) inmediatamente debajo,
        # NUNCA sobre la etiqueta misma.
        if data.get("observaciones"):
            for fila_celdas in ws3.iter_rows():
                hecho = False
                for celda in fila_celdas:
                    if normalizar(celda.value) == "OBSERVACIONES":
                        # Si la etiqueta esta en un merge, escribir en la
                        # primera fila despues de ese rango combinado.
                        rango = _rango_de(ws3, celda.row, celda.column)
                        if rango is not None:
                            fila_dest = rango.max_row + 1
                        else:
                            fila_dest = celda.row + 1
                        _escribir(ws3, fila_dest, celda.column,
                                  data["observaciones"])
                        hecho = True
                        break
                if hecho:
                    break

    # Guardar a bytes
    salida = BytesIO()
    wb.save(salida)
    salida.seek(0)
    generado = salida.read()

    # PRESERVAR LOGOS: reinyectar imagenes originales de la plantilla.
    # Esto garantiza que los logos de encabezado y pie no se pierdan,
    # aunque openpyxl los haya alterado al guardar.
    return _reinyectar_imagenes(plantilla_bytes, generado)


# ============================================================
#  COTIZACION EN MERCADO LIBRE
# ============================================================

def _filtrar_outliers(precios):
    """
    Quita valores atipicos usando el rango intercuartilico (IQR).
    Evita que un precio absurdamente bajo o alto ensucie el promedio.
    """
    if len(precios) < 4:
        if not precios:
            return []
        med = statistics.median(precios)
        return [p for p in precios if 0.2 * med <= p <= 3.0 * med]

    ordenados = sorted(precios)
    n = len(ordenados)
    q1 = ordenados[n // 4]
    q3 = ordenados[(3 * n) // 4]
    iqr = q3 - q1
    limite_inf = q1 - 1.5 * iqr
    limite_sup = q3 + 1.5 * iqr
    filtrados = [p for p in ordenados if limite_inf <= p <= limite_sup]
    return filtrados if filtrados else ordenados


def cotizar_pieza(pieza, marca="", modelo=""):
    """
    Busca una pieza en Mercado Libre Argentina y devuelve un precio promedio
    limpio (sin valores atipicos).
    """
    resultado = {
        "pieza": pieza, "query": "", "precio_sugerido": 0,
        "cantidad_resultados": 0, "cantidad_usados": 0,
        "rango": (0, 0), "error": None,
    }

    partes = [pieza]
    if marca:
        partes.append(marca)
    if modelo:
        partes.append(modelo.split()[0])
    query = " ".join(partes).strip()
    resultado["query"] = query

    url = ("https://api.mercadolibre.com/sites/MLA/search"
           "?q=" + quote(query) + "&condition=new&limit=30")

    try:
        req = Request(url, headers={"User-Agent": "InformeTasacion/1.0"})
        with urlopen(req, timeout=12) as resp:
            datos = json.loads(resp.read().decode("utf-8"))

        items = datos.get("results", [])
        precios = [
            it["price"] for it in items
            if it.get("currency_id") == "ARS" and it.get("price", 0) > 0
        ]
        resultado["cantidad_resultados"] = len(precios)

        if not precios:
            resultado["error"] = "Sin resultados"
            return resultado

        precios_limpios = _filtrar_outliers(precios)
        if not precios_limpios:
            resultado["error"] = "Sin precios validos tras filtrar"
            return resultado

        promedio = int(round(sum(precios_limpios) / len(precios_limpios)))
        resultado["precio_sugerido"] = promedio
        resultado["cantidad_usados"] = len(precios_limpios)
        resultado["rango"] = (min(precios_limpios), max(precios_limpios))

    except Exception as e:
        resultado["error"] = str(e)

    return resultado


def cotizar_danos(data, solo_cambiar=True):
    """
    Cotiza todas las piezas a CAMBIAR de un informe.
    Modifica data["danos"] cargando el precio sugerido en cada pieza.
    Devuelve una lista con el detalle de cada cotizacion.
    """
    detalle = []
    marca = data.get("marca", "")
    modelo = data.get("modelo", "")

    for dano in data["danos"]:
        if solo_cambiar and dano["accion"] != "CAMBIAR":
            continue
        if dano.get("precio", 0) > 0:
            detalle.append({
                "pieza": dano["pieza"],
                "precio_sugerido": dano["precio"],
                "fuente": "cargado manualmente",
            })
            continue

        cot = cotizar_pieza(dano["pieza"], marca, modelo)
        if cot["precio_sugerido"] > 0:
            dano["precio"] = cot["precio_sugerido"]
        detalle.append({
            "pieza": dano["pieza"],
            "precio_sugerido": cot["precio_sugerido"],
            "cantidad_resultados": cot["cantidad_resultados"],
            "cantidad_usados": cot["cantidad_usados"],
            "rango": cot["rango"],
            "fuente": "Mercado Libre" if cot["precio_sugerido"] > 0 else "sin cotizacion",
            "error": cot["error"],
        })
        time.sleep(0.25)

    return detalle


# ============================================================
#  FUNCION PRINCIPAL
# ============================================================

def procesar(plantilla_bytes, texto_peritacion="", excel_peritacion_bytes=None,
             cotizar=False):
    """
    Funcion principal: recibe la plantilla y la peritacion,
    devuelve (informe_bytes, data, detalle_cotizacion).
    """
    data_texto = parsear_texto(texto_peritacion) if texto_peritacion else data_vacia()

    if excel_peritacion_bytes:
        data_excel = parsear_excel_peritacion(excel_peritacion_bytes)
        data = combinar_datos(data_excel, data_texto)
    else:
        data = data_texto

    detalle_cotizacion = []
    if cotizar:
        detalle_cotizacion = cotizar_danos(data)

    informe_bytes = completar_plantilla(plantilla_bytes, data)
    return informe_bytes, data, detalle_cotizacion
