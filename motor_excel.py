from __future__ import annotations

import copy
import hashlib
import io
import re
import zipfile
from collections import OrderedDict, defaultdict
from typing import Any
from xml.etree import ElementTree as ET

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

NS = {"a": MAIN_NS, "r": REL_NS, "pr": PKG_REL_NS}
ET.register_namespace("", MAIN_NS)
ET.register_namespace("r", REL_NS)

NAMESPACES_ESTANDAR_EXCEL = {
    "": MAIN_NS,
    "r": REL_NS,
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "x14": "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main",
    "x14ac": "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac",
    "x15": "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main",
    "x16r2": "http://schemas.microsoft.com/office/spreadsheetml/2015/02/main",
    "xr": "http://schemas.microsoft.com/office/spreadsheetml/2014/revision",
    "xr2": "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2",
    "xr3": "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3",
    "xr6": "http://schemas.microsoft.com/office/spreadsheetml/2016/revision6",
    "xr10": "http://schemas.microsoft.com/office/spreadsheetml/2016/revision10",
}


def _registrar_namespaces_estandar_excel() -> None:
    """Usa los prefijos oficiales incluso si se carga un Excel exportado por una versión antigua."""
    for prefijo, uri in NAMESPACES_ESTANDAR_EXCEL.items():
        try:
            ET.register_namespace(prefijo, uri)
        except ValueError:
            pass


CRITERIOS = OrderedDict(
    [
        ("GANA", {"color": "VERDE", "fill_rgb": "FF92D050"}),
        ("NO GANA", {"color": "ROJO", "fill_rgb": "FFFF0000"}),
        ("NESTUM EN LA EXHIBICION", {"color": "ROJO", "fill_rgb": "FFFF0000"}),
        ("NESTOGENO EN LA EXHIBICION", {"color": "ROJO", "fill_rgb": "FFFF0000"}),
        ("NESTUM Y NESTOGENO EN LA EXHIBICION", {"color": "ROJO", "fill_rgb": "FFFF0000"}),
        ("NO TIENE 3 O MAS CATEGORIAS", {"color": "ROJO", "fill_rgb": "FFFF0000"}),
        ("POCOS PRODUCTOS", {"color": "AMARILLO", "fill_rgb": "FFFFFF00"}),
        ("NO ES LA MISMA TIENDA", {"color": "AMARILLO", "fill_rgb": "FFFFFF00"}),
    ]
)

COLUMNAS_REQUERIDAS = {
    "C": "IDCliente",
    "S": "Febrero Fotos",
    "T": "Marzo 2026",
    "U": "Abril 2026",
    "W": "ANOTACIONES",
}

COLUMNAS_TABLA = [
    "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "S", "T", "U", "V", "W"
]

MESES_FOTO = [
    ("S", "Febrero"),
    ("T", "Marzo"),
    ("U", "Abril"),
]


def qn(tag: str) -> str:
    return f"{{{MAIN_NS}}}{tag}"


def _registrar_namespaces_originales(xml_contenido: bytes) -> list[tuple[str, str]]:
    """Registra los prefijos originales antes de volver a serializar XML de Excel."""
    namespaces: list[tuple[str, str]] = []
    vistos: set[tuple[str, str]] = set()
    for _, dato in ET.iterparse(io.BytesIO(xml_contenido), events=("start-ns",)):
        prefijo, uri = dato
        prefijo = prefijo or ""
        clave = (prefijo, uri)
        if clave in vistos:
            continue
        vistos.add(clave)
        namespaces.append(clave)
        try:
            ET.register_namespace(prefijo, uri)
        except ValueError:
            # Algunos prefijos reservados no pueden registrarse; ElementTree
            # mantiene igualmente una declaración válida cuando sean usados.
            pass
    return namespaces


def _declaraciones_raiz_originales(xml_contenido: bytes, etiqueta_raiz: bytes) -> list[tuple[str, str]]:
    """Obtiene únicamente los xmlns declarados en la etiqueta raíz original."""
    inicio = xml_contenido.find(b"<" + etiqueta_raiz)
    if inicio == -1:
        return []
    fin = xml_contenido.find(b">", inicio)
    apertura = xml_contenido[inicio:fin]
    patron = re.compile(rb'\sxmlns(?::([A-Za-z_][A-Za-z0-9_.-]*))?="([^"]+)"')
    resultado: list[tuple[str, str]] = []
    for coincidencia in patron.finditer(apertura):
        prefijo = coincidencia.group(1).decode("utf-8") if coincidencia.group(1) else ""
        uri = coincidencia.group(2).decode("utf-8")
        resultado.append((prefijo, uri))
    return resultado


def _restaurar_declaraciones_raiz(
    xml_serializado: bytes, etiqueta_raiz: bytes, declaraciones: list[tuple[str, str]]
) -> bytes:
    """Restaura xmlns omitidos por ElementTree pero referenciados por mc:Ignorable."""
    inicio = xml_serializado.find(b"<" + etiqueta_raiz)
    if inicio == -1:
        return xml_serializado
    fin = xml_serializado.find(b">", inicio)
    apertura = xml_serializado[inicio:fin]
    faltantes: list[bytes] = []
    for prefijo, uri in declaraciones:
        atributo = b'xmlns="' if not prefijo else f'xmlns:{prefijo}="'.encode("utf-8")
        if atributo not in apertura:
            nombre = "xmlns" if not prefijo else f"xmlns:{prefijo}"
            faltantes.append(f' {nombre}="{uri}"'.encode("utf-8"))
    if faltantes:
        xml_serializado = xml_serializado[:fin] + b"".join(faltantes) + xml_serializado[fin:]
    return xml_serializado


def _garantizar_namespaces_ignorables(xml_serializado: bytes, etiqueta_raiz: bytes) -> bytes:
    """Agrega declaraciones para prefijos mencionados en mc:Ignorable si faltan."""
    inicio = xml_serializado.find(b"<" + etiqueta_raiz)
    if inicio == -1:
        return xml_serializado
    fin = xml_serializado.find(b">", inicio)
    apertura = xml_serializado[inicio:fin]
    coincidencia = re.search(rb'(?:mc|[A-Za-z_][A-Za-z0-9_.-]*):Ignorable="([^"]+)"', apertura)
    if not coincidencia:
        return xml_serializado
    faltantes: list[bytes] = []
    for prefijo_bytes in coincidencia.group(1).split():
        prefijo = prefijo_bytes.decode("utf-8")
        uri = NAMESPACES_ESTANDAR_EXCEL.get(prefijo)
        if uri and f'xmlns:{prefijo}="'.encode("utf-8") not in apertura:
            faltantes.append(f' xmlns:{prefijo}="{uri}"'.encode("utf-8"))
    if faltantes:
        xml_serializado = xml_serializado[:fin] + b"".join(faltantes) + xml_serializado[fin:]
    return xml_serializado


def letra_a_numero(letras: str) -> int:
    numero = 0
    for letra in letras.upper():
        numero = numero * 26 + ord(letra) - 64
    return numero


def columna_de_referencia(referencia: str) -> str:
    encontrado = re.match(r"([A-Z]+)", referencia or "")
    return encontrado.group(1) if encontrado else ""


def _leer_shared_strings(archivo_zip: zipfile.ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in archivo_zip.namelist():
        return []
    raiz = ET.fromstring(archivo_zip.read("xl/sharedStrings.xml"))
    resultado: list[str] = []
    for si in raiz.findall("a:si", NS):
        resultado.append("".join(t.text or "" for t in si.findall(".//a:t", NS)))
    return resultado


def _valor_celda(celda: ET.Element, shared_strings: list[str]) -> str:
    formula = celda.find("a:f", NS)
    if formula is not None:
        return "=" + (formula.text or "")
    tipo = celda.attrib.get("t")
    if tipo == "inlineStr":
        return "".join(t.text or "" for t in celda.findall(".//a:t", NS))
    valor = celda.find("a:v", NS)
    if valor is None or valor.text is None:
        return ""
    if tipo == "s":
        try:
            return shared_strings[int(valor.text)]
        except (IndexError, ValueError):
            return ""
    return valor.text


def _mapa_fila(fila: ET.Element, shared_strings: list[str]) -> dict[str, str]:
    return {
        columna_de_referencia(celda.attrib.get("r", "")): _valor_celda(celda, shared_strings)
        for celda in fila.findall("a:c", NS)
    }


def _ruta_hoja(archivo_zip: zipfile.ZipFile, nombre_hoja: str) -> str:
    workbook = ET.fromstring(archivo_zip.read("xl/workbook.xml"))
    relaciones = ET.fromstring(archivo_zip.read("xl/_rels/workbook.xml.rels"))
    mapa_relaciones = {
        relacion.attrib["Id"]: relacion.attrib["Target"]
        for relacion in relaciones.findall("pr:Relationship", NS)
    }
    for hoja in workbook.find("a:sheets", NS) or []:
        if hoja.attrib.get("name") == nombre_hoja:
            relacion_id = hoja.attrib.get(f"{{{REL_NS}}}id")
            destino = mapa_relaciones.get(relacion_id or "", "")
            return destino if destino.startswith("xl/") else "xl/" + destino.lstrip("/")
    raise ValueError(f"No se encontró la hoja '{nombre_hoja}' en el archivo.")


def _es_url_foto(valor: str) -> bool:
    valor = (valor or "").strip()
    return valor.lower().startswith(("http://", "https://"))



COLORES_EXCEL_REFERENCIA = {
    "VERDE": {
        "FF92D050",  # verde exacto solicitado
        "FFC6EFCE",  # verde claro de validaciones anteriores
        "FF00B050",
        "FF70AD47",
        "FF00FF00",
    },
    "ROJO": {
        "FFFF0000",  # rojo exacto solicitado
        "FFFFC7CE",  # rojo claro de validaciones anteriores
        "FFC00000",
    },
    "AMARILLO": {
        "FFFFFF00",  # amarillo exacto solicitado
        "FFFFEB9C",  # amarillo claro de validaciones anteriores
        "FFFFC000",
        "FFFFD966",
    },
}


def _normalizar_rgb_excel(valor: str | None) -> str:
    valor = (valor or "").replace("#", "").upper().strip()
    if len(valor) == 6:
        valor = "FF" + valor
    return valor


def _clasificar_rgb_excel(valor: str | None) -> str:
    rgb = _normalizar_rgb_excel(valor)
    for nombre_color, referencias in COLORES_EXCEL_REFERENCIA.items():
        if rgb in referencias:
            return nombre_color
    return ""


def _mapa_estilos_a_colores(raiz_estilos: ET.Element) -> dict[int, str]:
    """Devuelve un mapa style_index -> VERDE/ROJO/AMARILLO cuando el estilo tiene fill reconocible."""
    resultado: dict[int, str] = {}
    fills = raiz_estilos.find("a:fills", NS)
    cell_xfs = raiz_estilos.find("a:cellXfs", NS)
    if fills is None or cell_xfs is None:
        return resultado

    fills_lista = list(fills)
    for indice_estilo, xf in enumerate(list(cell_xfs)):
        try:
            fill_id = int(xf.attrib.get("fillId", "0"))
        except ValueError:
            continue
        if fill_id < 0 or fill_id >= len(fills_lista):
            continue
        fill = fills_lista[fill_id]
        fg = fill.find(".//a:fgColor", NS)
        if fg is None:
            continue
        color = _clasificar_rgb_excel(fg.attrib.get("rgb"))
        if color:
            resultado[indice_estilo] = color
    return resultado


def _colores_fila_excel(fila: ET.Element, mapa_estilos: dict[int, str]) -> list[str]:
    colores: set[str] = set()
    for celda in fila.findall("a:c", NS):
        columna = columna_de_referencia(celda.attrib.get("r", ""))
        if not columna:
            continue
        if not (letra_a_numero("A") <= letra_a_numero(columna) <= letra_a_numero("W")):
            continue
        try:
            estilo = int(celda.attrib.get("s", "0"))
        except ValueError:
            estilo = 0
        color = mapa_estilos.get(estilo)
        if color:
            colores.add(color)
    return sorted(colores)


def _color_principal(colores: list[str]) -> str:
    if not colores:
        return "SIN COLOR"
    if len(colores) == 1:
        return colores[0]
    return "MIXTO"

def analizar_excel(contenido: bytes) -> dict[str, Any]:
    """Lee el Excel, detecta pendientes, validados y colores existentes para filtros de revisión."""
    with zipfile.ZipFile(io.BytesIO(contenido), "r") as archivo_zip:
        shared_strings = _leer_shared_strings(archivo_zip)
        ruta_consolidado = _ruta_hoja(archivo_zip, "Consolidado")
        raiz_hoja = ET.fromstring(archivo_zip.read(ruta_consolidado))
        raiz_estilos = ET.fromstring(archivo_zip.read("xl/styles.xml")) if "xl/styles.xml" in archivo_zip.namelist() else ET.Element(qn("styleSheet"))
        mapa_estilos_colores = _mapa_estilos_a_colores(raiz_estilos)

        filas = raiz_hoja.findall(".//a:sheetData/a:row", NS)
        if not filas:
            raise ValueError("La hoja Consolidado está vacía.")

        encabezados_por_columna = _mapa_fila(filas[0], shared_strings)
        faltantes = [
            f"{columna} ({nombre})"
            for columna, nombre in COLUMNAS_REQUERIDAS.items()
            if not encabezados_por_columna.get(columna, "").strip()
        ]
        if faltantes:
            raise ValueError(
                "No se localizaron las columnas requeridas en la hoja Consolidado: "
                + ", ".join(faltantes)
            )

        registros_por_id: dict[str, list[dict[str, Any]]] = defaultdict(list)
        for fila in filas[1:]:
            numero_fila = int(fila.attrib.get("r", "0"))
            datos = _mapa_fila(fila, shared_strings)
            id_cliente = datos.get("C", "").strip()
            if id_cliente:
                registros_por_id[id_cliente].append(
                    {
                        "fila_excel": numero_fila,
                        "datos": datos,
                        "colores_excel": _colores_fila_excel(fila, mapa_estilos_colores),
                    }
                )

        tiendas_todas: list[dict[str, Any]] = []
        tiendas_pendientes: list[dict[str, Any]] = []
        tiendas_omitidas: list[dict[str, Any]] = []
        total_con_fotos = 0

        conteo_colores = {"VERDE": 0, "ROJO": 0, "AMARILLO": 0, "MIXTO": 0, "SIN COLOR": 0}

        for id_cliente, registros in registros_por_id.items():
            representante = next(
                (
                    registro
                    for registro in registros
                    if all(_es_url_foto(registro["datos"].get(col, "")) for col in ("S", "T", "U"))
                ),
                None,
            )
            if representante is None:
                continue

            total_con_fotos += 1

            observaciones_previas = sorted(
                {
                    registro["datos"].get("V", "").strip()
                    for registro in registros
                    if registro["datos"].get("V", "").strip()
                }
            )
            anotaciones_previas = sorted(
                {
                    registro["datos"].get("W", "").strip()
                    for registro in registros
                    if registro["datos"].get("W", "").strip()
                }
            )
            colores_id = sorted({color for registro in registros for color in registro.get("colores_excel", [])})
            color_excel = _color_principal(colores_id)
            conteo_colores[color_excel] = conteo_colores.get(color_excel, 0) + 1
            validada_previa = bool(observaciones_previas or anotaciones_previas)

            tienda = {
                "id_cliente": id_cliente,
                "nombre_cliente": representante["datos"].get("E", "")
                or representante["datos"].get("D", ""),
                "regional": representante["datos"].get("F", ""),
                "filas_excel": [registro["fila_excel"] for registro in registros],
                "representante": representante["datos"],
                "validada_previa": validada_previa,
                "observaciones_previas": observaciones_previas,
                "anotaciones_previas": anotaciones_previas,
                "color_excel": color_excel,
                "colores_excel": colores_id,
                "filas_tabla": [
                    {
                        "Fila Excel": registro["fila_excel"],
                        "Color Excel": ", ".join(registro.get("colores_excel", [])) or "SIN COLOR",
                        **{
                            encabezados_por_columna.get(columna, columna): registro["datos"].get(columna, "")
                            for columna in COLUMNAS_TABLA
                        },
                    }
                    for registro in registros
                ],
            }
            tiendas_todas.append(tienda)

            if validada_previa:
                tiendas_omitidas.append(
                    {
                        "id_cliente": id_cliente,
                        "filas_excel": [registro["fila_excel"] for registro in registros],
                        "tiene_observacion": bool(observaciones_previas),
                        "tiene_anotacion": bool(anotaciones_previas),
                        "observaciones": observaciones_previas,
                        "anotaciones": anotaciones_previas,
                        "color_excel": color_excel,
                        "colores_excel": colores_id,
                    }
                )
            else:
                tiendas_pendientes.append(tienda)

        tiendas_todas.sort(key=lambda item: min(item["filas_excel"]))
        tiendas_pendientes.sort(key=lambda item: min(item["filas_excel"]))
        tiendas_omitidas.sort(key=lambda item: min(item["filas_excel"]))
        total_filas_objetivo = sum(len(tienda["filas_excel"]) for tienda in tiendas_pendientes)

        return {
            "hash_archivo": hashlib.sha256(contenido).hexdigest(),
            "ruta_consolidado": ruta_consolidado,
            "encabezados": encabezados_por_columna,
            "tiendas": tiendas_pendientes,
            "tiendas_todas": tiendas_todas,
            "total_tiendas": len(tiendas_pendientes),
            "total_tiendas_todas": len(tiendas_todas),
            "total_filas_objetivo": total_filas_objetivo,
            "total_con_fotos": total_con_fotos,
            "tiendas_omitidas": tiendas_omitidas,
            "total_tiendas_omitidas": len(tiendas_omitidas),
            "omitidas_con_observacion": sum(1 for t in tiendas_omitidas if t["tiene_observacion"]),
            "omitidas_con_anotacion": sum(1 for t in tiendas_omitidas if t["tiene_anotacion"]),
            "conteo_colores": conteo_colores,
            "decisiones_precargadas": {},
        }


def _obtener_o_crear_celda(fila: ET.Element, columna: str) -> ET.Element:
    numero_fila = fila.attrib["r"]
    referencia = f"{columna}{numero_fila}"
    celdas = fila.findall("a:c", NS)
    for celda in celdas:
        if celda.attrib.get("r") == referencia:
            return celda

    nueva_celda = ET.Element(qn("c"), {"r": referencia})
    numero_columna = letra_a_numero(columna)
    posicion = len(celdas)
    for indice, celda in enumerate(celdas):
        if letra_a_numero(columna_de_referencia(celda.attrib.get("r", ""))) > numero_columna:
            posicion = indice
            break
    fila.insert(posicion, nueva_celda)
    return nueva_celda


def _escribir_texto_inline(celda: ET.Element, texto: str) -> None:
    estilo = celda.attrib.get("s")
    referencia = celda.attrib["r"]
    celda.clear()
    celda.attrib["r"] = referencia
    if estilo is not None:
        celda.attrib["s"] = estilo
    celda.attrib["t"] = "inlineStr"
    inline = ET.SubElement(celda, qn("is"))
    nodo_texto = ET.SubElement(inline, qn("t"))
    nodo_texto.text = texto


def _agregar_fills_colores(raiz_estilos: ET.Element) -> dict[str, int]:
    fills = raiz_estilos.find("a:fills", NS)
    if fills is None:
        raise ValueError("El archivo no contiene definición de estilos/fills.")

    colores_necesarios = {detalle["fill_rgb"] for detalle in CRITERIOS.values()}
    fill_por_color: dict[str, int] = {}

    for indice, fill in enumerate(list(fills)):
        fg = fill.find(".//a:fgColor", NS)
        if fg is not None and fg.attrib.get("rgb") in colores_necesarios:
            fill_por_color[fg.attrib["rgb"]] = indice

    for color in colores_necesarios:
        if color in fill_por_color:
            continue
        nuevo_fill = ET.SubElement(fills, qn("fill"))
        patron = ET.SubElement(nuevo_fill, qn("patternFill"), {"patternType": "solid"})
        ET.SubElement(patron, qn("fgColor"), {"rgb": color})
        ET.SubElement(patron, qn("bgColor"), {"indexed": "64"})
        fill_por_color[color] = len(list(fills)) - 1

    fills.attrib["count"] = str(len(list(fills)))
    return fill_por_color


def _crear_estilos_coloreados(
    raiz_estilos: ET.Element, estilos_requeridos: set[tuple[int, str]]
) -> dict[tuple[int, str], int]:
    cell_xfs = raiz_estilos.find("a:cellXfs", NS)
    if cell_xfs is None:
        raise ValueError("El archivo no contiene definición de estilos/cellXfs.")

    fill_por_color = _agregar_fills_colores(raiz_estilos)
    xfs_originales = list(cell_xfs)
    mapa_estilo: dict[tuple[int, str], int] = {}

    for estilo_origen, color in sorted(estilos_requeridos):
        if estilo_origen >= len(xfs_originales):
            estilo_origen = 0
        nuevo_xf = copy.deepcopy(xfs_originales[estilo_origen])
        nuevo_xf.attrib["fillId"] = str(fill_por_color[color])
        nuevo_xf.attrib["applyFill"] = "1"
        cell_xfs.append(nuevo_xf)
        mapa_estilo[(estilo_origen, color)] = len(list(cell_xfs)) - 1

    cell_xfs.attrib["count"] = str(len(list(cell_xfs)))
    return mapa_estilo


def exportar_excel(
    contenido_original: bytes, analisis: dict[str, Any], decisiones: dict[str, str]
) -> bytes:
    """Exporta una copia del Excel aplicando anotación en W y color de A:W a las filas del ID."""
    decisiones_validas = {
        id_cliente: criterio
        for id_cliente, criterio in decisiones.items()
        if criterio in CRITERIOS
    }
    if not decisiones_validas:
        raise ValueError("No hay decisiones guardadas para exportar.")

    tiendas_exportables = analisis.get("tiendas_todas") or analisis.get("tiendas") or []
    filas_por_id = {
        tienda["id_cliente"]: tienda["filas_excel"] for tienda in tiendas_exportables
    }

    entrada = io.BytesIO(contenido_original)
    salida = io.BytesIO()

    with zipfile.ZipFile(entrada, "r") as origen:
        ruta_hoja = analisis["ruta_consolidado"]
        xml_hoja_original = origen.read(ruta_hoja)
        xml_estilos_original = origen.read("xl/styles.xml")

        # Conservar los prefijos XML originales del libro evita que Excel
        # marque el archivo exportado como dañado al abrirlo.
        _registrar_namespaces_originales(xml_hoja_original)
        _registrar_namespaces_originales(xml_estilos_original)
        # Una exportación antigua podía traer ns1/ns2 en vez de mc/x14ac/xr.
        # Forzamos los prefijos oficiales para que la salida quede válida.
        _registrar_namespaces_estandar_excel()
        declaraciones_hoja = _declaraciones_raiz_originales(xml_hoja_original, b"worksheet")
        declaraciones_estilos = _declaraciones_raiz_originales(xml_estilos_original, b"styleSheet")

        raiz_hoja = ET.fromstring(xml_hoja_original)
        raiz_estilos = ET.fromstring(xml_estilos_original)

        filas_xml = {
            int(fila.attrib.get("r", "0")): fila
            for fila in raiz_hoja.findall(".//a:sheetData/a:row", NS)
        }

        estilos_requeridos: set[tuple[int, str]] = set()
        actualizaciones: list[tuple[ET.Element, str, int]] = []
        for id_cliente, criterio in decisiones_validas.items():
            color = CRITERIOS[criterio]["fill_rgb"]
            for numero_fila in filas_por_id.get(id_cliente, []):
                fila = filas_xml.get(numero_fila)
                if fila is None:
                    continue
                celda_anotacion = _obtener_o_crear_celda(fila, "W")
                _escribir_texto_inline(celda_anotacion, criterio)
                for numero_columna in range(letra_a_numero("A"), letra_a_numero("W") + 1):
                    columna = ""
                    valor = numero_columna
                    while valor:
                        valor, resto = divmod(valor - 1, 26)
                        columna = chr(65 + resto) + columna
                    celda = _obtener_o_crear_celda(fila, columna)
                    estilo_origen = int(celda.attrib.get("s", "0"))
                    estilos_requeridos.add((estilo_origen, color))
                    actualizaciones.append((celda, color, estilo_origen))

        estilos_nuevos = _crear_estilos_coloreados(raiz_estilos, estilos_requeridos)
        for celda, color, estilo_origen in actualizaciones:
            celda.attrib["s"] = str(estilos_nuevos[(estilo_origen, color)])

        hoja_modificada = ET.tostring(raiz_hoja, encoding="utf-8", xml_declaration=True)
        estilos_modificados = ET.tostring(raiz_estilos, encoding="utf-8", xml_declaration=True)
        hoja_modificada = _restaurar_declaraciones_raiz(
            hoja_modificada, b"worksheet", declaraciones_hoja
        )
        estilos_modificados = _restaurar_declaraciones_raiz(
            estilos_modificados, b"styleSheet", declaraciones_estilos
        )
        hoja_modificada = _garantizar_namespaces_ignorables(hoja_modificada, b"worksheet")
        estilos_modificados = _garantizar_namespaces_ignorables(estilos_modificados, b"styleSheet")

        with zipfile.ZipFile(salida, "w", compression=zipfile.ZIP_DEFLATED) as destino:
            for informacion in origen.infolist():
                if informacion.filename == ruta_hoja:
                    destino.writestr(informacion, hoja_modificada)
                elif informacion.filename == "xl/styles.xml":
                    destino.writestr(informacion, estilos_modificados)
                else:
                    destino.writestr(informacion, origen.read(informacion.filename))

    return salida.getvalue()
