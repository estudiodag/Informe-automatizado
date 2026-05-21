"""
procesador.py
Motor de procesamiento de informes de tasación.
- Parsea el texto/Excel de la peritación.
- Completa la plantilla virgen preservando logos, estilos y fórmulas.
"""

import re
import time
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
    """Pasa a mayúsculas y quita tildes."""
    if texto is None:
        return ""
    texto = str(texto).upper()
    texto = unicodedata.normalize("NFD", texto)
    texto = "".join(c for c in texto if unicodedata.category(c) != "Mn")
    return texto.strip()


def a_numero(valor):
    """Convierte un string a número, quitando $, puntos y comas."""
    if valor is None:
        return 0
    s = re.sub(r"[^\d]", "", str(valor))
    return int(s) if s else 0


def normalizar_accion(raw):
    """Normaliza la acción a CAMBIAR o REPARAR. 'Pintar' -> 'REPARAR'."""
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
    (["ANO"], "anio"),
    (["PATENTE", "DOMINIO"], "dominio"),
    (["CHASIS", "N CHASIS", "NRO CHASIS"], "chasis"),
    (["KILOMETRAJE", "KM"], "kilometraje"),
    (["SUMA ASEG", "SUMA ASEGURADA", "S ASEG"], "sumaAsegurada"),
    (["FRANQUICIA A DEDUCIR", "FRANQUICIA"], "franquicia"),
    (["DIRECCION", "DIR"], "tallerDireccion"),
    (["LOCALIDAD"], "tallerLocalidad"),
]

# Etiquetas que se ignoran pero sirven de delimitador
IGNORE_LABELS = [
    "FECHA CARGA", "SELLO", "POLIZA", "PROP", "STROS", "PRODUCTOR", "DOM.PAS",
    "CP", "CATEGORIA IVA", "AG. RET CUIT", "CUIT", "NRO JUB", "TIPO JUBILACION",
    "ING BRUTOS", "NRO IB", "JURISDICCION PAS", "SER. SOCIAL", "DOMICILIO",
    "ITEM", "USO", "MOTOR",
]


def _construir_patron_labels():
    todas = []
    for variantes, _ in LABEL_DEFS:
        todas.extend(variantes)
    todas.extend(IGNORE_LABELS)
    # Ordenar por longitud descendente para que las más largas matcheen primero
    todas.sort(key=len, reverse=True)
    return "|".join(re.escape(x) for x in todas)


def parsear_texto(texto):
    """Parsea el texto libre de la peritación y devuelve el dict de datos."""
    data = data_vacia()
    if not texto or not texto.strip():
        return data

    labels_pattern = _construir_patron_labels()
    lineas = texto.split("\n")
    lineas_norm = [normalizar(l) for l in lineas]

    # ---- PASO 1: extraer campos de cabecera ----
    for idx, linea_norm in enumerate(lineas_norm):
        if not linea_norm.strip():
            continue
        linea_orig = lineas[idx]

        patron = r"(" + labels_pattern + r")\s*[:=]\s*(.*?)(?=\s+(?:" + labels_pattern + r")\s*[:=]|$)"
        for m in re.finditer(patron, linea_norm, re.IGNORECASE):
            label = m.group(1).strip()
            valor_raw = m.group(2)
            # Calcular posición del valor en la línea original
            inicio = m.start(2)
            valor = linea_orig[inicio:inicio + len(valor_raw)].strip()
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

    # ---- PASO 2: secciones de daños y mano de obra ----
    sustituir_idx = -1
    pintar_idx = -1
    for i, ln in enumerate(lineas_norm):
        ln = ln.strip()
        if sustituir_idx < 0 and re.match(r"^(SUSTITUIR|CAMBIAR|REEMPLAZAR)\s*:?\s*$", ln):
            sustituir_idx = i
        if pintar_idx < 0 and re.match(r"^(PINTAR|REPARAR)\s*:?\s*$", ln):
            pintar_idx = i

    def es_mano_obra(linea):
        return re.match(
            r"^\d+(?:[.,]\d+)?\s*(?:HS\s+|HORAS?\s+)?(PA[NN]OS?|CHAPA|MECANICA|TAPICERIA|PINTURA)",
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
            r"FRANQUICIA|LOCALIDAD|DIRECCION|POLIZA|PRODUCTOR|CUIT|ITEM|USO)\s*[:=]",
            normalizar(linea),
        )

    def es_seccion(linea):
        ln = normalizar(linea).strip()
        return re.match(
            r"^(SUSTITUIR|CAMBIAR|PINTAR|REPARAR|OBSERVACIONES|OBS|MANO DE OBRA|"
            r"COTIZACION|REEMPLAZAR)\s*:?\s*$",
            ln,
        )

    def procesar_bloque_danos(inicio, accion):
        """Lee piezas desde 'inicio' hasta que aparece otra sección/campo/MO."""
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
    # Toma líneas largas que no son campos ni piezas ni MO.
    obs = []
    capturando = False
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
                obs.append(l)
                continue
        # Observaciones implícitas: líneas largas sin etiqueta
        if (len(l) > 50 and not es_campo(l) and not es_mano_obra(l)
                and not re.match(r"^\d+\s", l)):
            obs.append(l)
    data["observaciones"] = "\n".join(obs)

    return data


# ============================================================
#  PARSER DE EXCEL DE PERITACIÓN
# ============================================================

def parsear_excel_peritacion(file_bytes):
    """Lee un Excel de peritación y extrae los datos."""
    data = data_vacia()
    wb = load_workbook(BytesIO(file_bytes), data_only=True)

    # Hoja 1: datos generales
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

    # Hoja 2: daños (detecta X en col A=CAMBIAR, col B=REPARAR)
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
#  COMPLETAR LA PLANTILLA (PRESERVA LOGOS Y FORMATO)
# ============================================================

def _rango_de(ws, fila, columna):
    """Devuelve el CellRange combinado que contiene a (fila, columna),
    o None si la celda no es parte de ningún merge."""
    for rango in ws.merged_cells.ranges:
        if (rango.min_row <= fila <= rango.max_row and
                rango.min_col <= columna <= rango.max_col):
            return rango
    return None


def _valor_celda(ws, fila, columna):
    """
    Lee el valor de una celda de forma segura.
    Si la celda es parte de un rango combinado, el valor real está
    en la celda ancla (esquina superior izquierda). Las MergedCell
    no-ancla siempre tienen value=None.
    """
    rango = _rango_de(ws, fila, columna)
    if rango is not None:
        return ws.cell(row=rango.min_row, column=rango.min_col).value
    return ws.cell(row=fila, column=columna).value


def _escribir(ws, fila, columna, valor):
    """
    Escribe un valor en una celda manejando celdas combinadas.

    FIX del error "'MergedCell' object attribute 'value' is read-only":
    si la celda destino cae dentro de un rango combinado, se DESHACE
    el merge antes de escribir. Así la celda deja de ser una MergedCell
    inmutable y el value pasa a ser escribible. Luego se vuelve a
    combinar el rango para preservar el formato visual de la plantilla.
    """
    rango = _rango_de(ws, fila, columna)
    if rango is not None:
        rango_str = str(rango)
        anc_row, anc_col = rango.min_row, rango.min_col
        # Deshacer el merge -> todas las celdas del rango pasan a ser escribibles
        ws.unmerge_cells(rango_str)
        # El valor de un rango combinado SIEMPRE va en la celda ancla
        ws.cell(row=anc_row, column=anc_col).value = valor
        # Volver a combinar para conservar el aspecto original de la plantilla
        ws.merge_cells(rango_str)
    else:
        ws.cell(row=fila, column=columna).value = valor


def _buscar_celda_etiqueta(ws, textos_buscados, ocurrencia=1):
    """
    Busca una celda que contenga una etiqueta y devuelve la coordenada
    (fila, columna) de la celda donde hay que escribir el valor
    (la siguiente celda vacía a la derecha en la misma fila).
    Devuelve None si no encuentra la etiqueta.
    """
    encontradas = 0
    for fila_celdas in ws.iter_rows():
        for celda in fila_celdas:
            valor = normalizar(celda.value)
            if not valor:
                continue
            for buscado in textos_buscados:
                bn = normalizar(buscado)
                if bn and (bn in valor or valor == bn):
                    encontradas += 1
                    if encontradas == ocurrencia:
                        fila = celda.row
                        col = celda.column
                        # Buscar la siguiente celda vacía a la derecha.
                        # Se usa _valor_celda para leer correctamente
                        # incluso si el destino es una celda combinada.
                        for c in range(col + 1, col + 9):
                            if _valor_celda(ws, fila, c) in (None, ""):
                                return (fila, c)
                        return (fila, col + 1)
    return None


def completar_plantilla(plantilla_bytes, data):
    """
    Carga la plantilla virgen y la completa con los datos.
    openpyxl PRESERVA: logos/imágenes, estilos, fórmulas, formato de impresión.
    Maneja correctamente las celdas combinadas (merged cells).
    Devuelve los bytes del Excel completado.
    """
    wb = load_workbook(BytesIO(plantilla_bytes))  # conserva fórmulas y formato
    hojas = wb.worksheets

    # ---------- HOJA 1: Datos preliminares ----------
    ws1 = hojas[0]

    mapeo_hoja1 = [
        (["NUMERO DE SINIESTRO", "NRO SINIESTRO"], data["numeroSiniestro"]),
        (["FECHA DE SINIESTRO", "FECHA SINIESTRO"], data["fechaSiniestro"]),
        (["FECHA DE INSPECC", "FECHA INSPECC"], data["fechaInspeccion"]),
        (["APELLIDO Y NOMBRE"], data["asegurado"]),
        (["MARCA"], data["marca"]),
        (["MODELO"], data["modelo"]),
        (["ANO"], data["anio"]),
        (["DOMINIO", "PATENTE"], data["dominio"]),
        (["CHASIS"], data["chasis"]),
        (["KILOMETRA"], data["kilometraje"]),
        (["SUMA ASEG"], data["sumaAsegurada"]),
        (["NOMBRE"], data["tallerNombre"]),
        (["DIRECCION"], data["tallerDireccion"]),
        (["LOCALIDAD"], data["tallerLocalidad"]),
    ]
    for etiquetas, valor in mapeo_hoja1:
        if not valor:
            continue
        pos = _buscar_celda_etiqueta(ws1, etiquetas)
        if pos is not None:
            _escribir(ws1, pos[0], pos[1], valor)

    # Franquicia del vehículo (primera ocurrencia de "FRANQUICIA")
    if data.get("franquiciaVeh"):
        pos = _buscar_celda_etiqueta(ws1, ["FRANQUICIA"], ocurrencia=1)
        if pos is not None:
            _escribir(ws1, pos[0], pos[1], data["franquiciaVeh"])

    # ---------- HOJA 2: Descripción de daños ----------
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
            col_precio = col_precio or 5
            for i, dano in enumerate(data["danos"]):
                fila = fila_inicio + i
                _escribir(ws2, fila, col_accion, dano["accion"])
                _escribir(ws2, fila, col_pieza, dano["pieza"])
                if dano["accion"] == "CAMBIAR" and dano["precio"] > 0:
                    _escribir(ws2, fila, col_precio, dano["precio"])

    # ---------- HOJA 3: Mano de obra, resumen, observaciones ----------
    if len(hojas) > 2:
        ws3 = hojas[2]
        mo = data["manoObra"]

        def set_mano_obra(nombre_concepto, cantidad, valor_unitario):
            """Busca la fila del concepto y carga cantidad y valor unitario."""
            for fila_celdas in ws3.iter_rows():
                for celda in fila_celdas:
                    v = normalizar(celda.value)
                    if v == normalizar(nombre_concepto):
                        col_cant = None
                        col_vu = None
                        for r in range(max(1, celda.row - 4), celda.row):
                            for c in range(1, 12):
                                hv = normalizar(_valor_celda(ws3, r, c))
                                if "CANT" in hv:
                                    col_cant = c
                                if "UNITARIO" in hv or "V. UNIT" in hv:
                                    col_vu = c
                        if col_cant:
                            _escribir(ws3, celda.row, col_cant, cantidad)
                        if col_vu and valor_unitario:
                            _escribir(ws3, celda.row, col_vu, valor_unitario)
                        return

        set_mano_obra("Pintura", mo["pintura"], mo["pinturaValor"])
        set_mano_obra("Chapa", mo["chapa"], mo["chapaValor"])
        set_mano_obra("Mecanica", mo["mecanica"], mo["mecanicaValor"])
        set_mano_obra("Tapiceria", mo["tapiceria"], mo["tapiceriaValor"])

        # Varios: va directo en SUBTOTAL
        for fila_celdas in ws3.iter_rows():
            for celda in fila_celdas:
                if normalizar(celda.value) == "VARIOS":
                    for r in range(max(1, celda.row - 6), celda.row):
                        for c in range(1, 12):
                            if "SUBTOTAL" in normalizar(_valor_celda(ws3, r, c)):
                                _escribir(ws3, celda.row, c, mo["varios"])
                                break

        # Franquicia a deducir
        if data.get("franquicia"):
            for fila_celdas in ws3.iter_rows():
                for celda in fila_celdas:
                    if "FRANQUICIA A DEDUCIR" in normalizar(celda.value):
                        for c in range(celda.column + 1, celda.column + 9):
                            if _valor_celda(ws3, celda.row, c) in (None, ""):
                                _escribir(ws3, celda.row, c, data["franquicia"])
                                break

        # Observaciones
        if data.get("observaciones"):
            for fila_celdas in ws3.iter_rows():
                for celda in fila_celdas:
                    if normalizar(celda.value) == "OBSERVACIONES":
                        _escribir(ws3, celda.row + 1, celda.column, data["observaciones"])
                        break

    # Guardar a bytes
    salida = BytesIO()
    wb.save(salida)
    salida.seek(0)
    return salida.read()


# ============================================================
#  COTIZACIÓN EN MERCADO LIBRE
# ============================================================

def _filtrar_outliers(precios):
    """
    Quita valores atípicos usando el método del rango intercuartílico (IQR).
    Esto evita que un precio absurdamente bajo o alto ensucie el promedio.

    Ejemplo: si los precios son [70k, 75k, 80k, 82k, 85k, 5k, 900k]
    -> descarta 5k y 900k, promedia el resto.
    """
    if len(precios) < 4:
        # Con pocos datos no se puede calcular IQR de forma confiable.
        # Filtramos lo evidente: descartar valores < 20% o > 300% de la mediana.
        if not precios:
            return []
        med = statistics.median(precios)
        return [p for p in precios if 0.2 * med <= p <= 3.0 * med]

    ordenados = sorted(precios)
    n = len(ordenados)
    # Cuartiles
    q1 = ordenados[n // 4]
    q3 = ordenados[(3 * n) // 4]
    iqr = q3 - q1
    # Límites: valores fuera de [Q1 - 1.5*IQR, Q3 + 1.5*IQR] son atípicos
    limite_inf = q1 - 1.5 * iqr
    limite_sup = q3 + 1.5 * iqr
    filtrados = [p for p in ordenados if limite_inf <= p <= limite_sup]
    return filtrados if filtrados else ordenados


def cotizar_pieza(pieza, marca="", modelo=""):
    """
    Busca una pieza en Mercado Libre Argentina y devuelve un precio promedio
    "limpio" (sin valores atípicos).

    Devuelve un dict:
      {
        "pieza": str,
        "query": str,
        "precio_sugerido": int,   # promedio filtrado, 0 si no hubo resultados
        "cantidad_resultados": int,
        "cantidad_usados": int,   # cuántos se usaron para el promedio
        "rango": (min, max),      # rango de precios usados
        "error": str or None
      }
    """
    resultado = {
        "pieza": pieza,
        "query": "",
        "precio_sugerido": 0,
        "cantidad_resultados": 0,
        "cantidad_usados": 0,
        "rango": (0, 0),
        "error": None,
    }

    # Construir la búsqueda: pieza + marca + modelo para más precisión
    partes = [pieza]
    if marca:
        partes.append(marca)
    if modelo:
        # Solo la primera palabra del modelo (ej: "Corolla" de "Corolla XLI 1.8")
        partes.append(modelo.split()[0])
    query = " ".join(partes).strip()
    resultado["query"] = query

    url = (
        "https://api.mercadolibre.com/sites/MLA/search"
        "?q=" + quote(query) +
        "&condition=new&limit=30"
    )

    try:
        req = Request(url, headers={"User-Agent": "InformeTasacion/1.0"})
        with urlopen(req, timeout=12) as resp:
            datos = json.loads(resp.read().decode("utf-8"))

        items = datos.get("results", [])
        # Solo precios en pesos argentinos y mayores a 0
        precios = [
            it["price"] for it in items
            if it.get("currency_id") == "ARS" and it.get("price", 0) > 0
        ]
        resultado["cantidad_resultados"] = len(precios)

        if not precios:
            resultado["error"] = "Sin resultados"
            return resultado

        # Filtrar valores atípicos
        precios_limpios = _filtrar_outliers(precios)
        if not precios_limpios:
            resultado["error"] = "Sin precios válidos tras filtrar"
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
    Devuelve una lista con el detalle de cada cotización (para mostrar al usuario).
    """
    detalle = []
    marca = data.get("marca", "")
    modelo = data.get("modelo", "")

    for dano in data["danos"]:
        # Solo cotizar piezas a CAMBIAR (las REPARAR no llevan precio de repuesto)
        if solo_cambiar and dano["accion"] != "CAMBIAR":
            continue
        # Si ya tiene precio cargado, no lo pisamos
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
            "fuente": "Mercado Libre" if cot["precio_sugerido"] > 0 else "sin cotización",
            "error": cot["error"],
        })

        # Pequeña pausa para no saturar la API de Mercado Libre
        time.sleep(0.25)

    return detalle


# ============================================================
#  FUNCIÓN PRINCIPAL
# ============================================================

def procesar(plantilla_bytes, texto_peritacion="", excel_peritacion_bytes=None,
             cotizar=False):
    """
    Función principal: recibe la plantilla y la peritación,
    devuelve los bytes del informe completado.

    Parámetros:
      plantilla_bytes        -- bytes del Excel virgen
      texto_peritacion       -- texto libre de la peritación
      excel_peritacion_bytes -- bytes de un Excel de peritación (opcional)
      cotizar                -- si True, busca precios en Mercado Libre

    Devuelve:
      (informe_bytes, data, detalle_cotizacion)
    """
    # Parsear la peritación
    data_texto = parsear_texto(texto_peritacion) if texto_peritacion else data_vacia()

    if excel_peritacion_bytes:
        data_excel = parsear_excel_peritacion(excel_peritacion_bytes)
        data = combinar_datos(data_excel, data_texto)
    else:
        data = data_texto

    # Cotizar en Mercado Libre si el usuario lo pidió
    detalle_cotizacion = []
    if cotizar:
        detalle_cotizacion = cotizar_danos(data)

    # Completar la plantilla
    informe_bytes = completar_plantilla(plantilla_bytes, data)
    return informe_bytes, data, detalle_cotizacion
