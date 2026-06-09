from __future__ import annotations

from datetime import datetime
from pathlib import Path
import json

import streamlit as st

from cache_imagenes import estadisticas_cache, obtener_imagen_cacheada, precargar_imagenes
from memoria import borrar_sesion, cargar_sesion, guardar_sesion, preparar_respaldo_json
from motor_excel import CRITERIOS, MESES_FOTO, analizar_excel, exportar_excel

# ============================================================
# CONFIGURACIÓN GENERAL
# ============================================================

st.set_page_config(
    page_title="Validación de Tiendas Ancla",
    page_icon="💠",
    layout="wide",
)

st.markdown(
    """
    <style>
      :root {
        --bg: #f5f7fb;
        --card: #ffffff;
        --border: #e4e7ec;
        --text-main: #101828;
        --text-muted: #667085;
        --green: #92D050;
        --yellow: #FFFF00;
        --blue: #2563eb;
        --control-blue: #2563EB;
        --control-blue-dark: #1D4ED8;
        --control-blue-light: #38BDF8;
        --control-blue-soft: rgba(37, 99, 235, 0.18);
        --control-track: rgba(148, 163, 184, 0.35);
      }

      .block-container {
        padding-top: 1.2rem;
        padding-bottom: 2.2rem;
        max-width: 1500px;
      }

      .app-header {
        background: white;
        color: black;
        padding: 1.25rem 1.45rem;
        border-radius: 18px;
        margin-bottom: 1rem;
        border: 1px solid rgba(255,255,255,0.12);
      }

      .app-title {
        font-size: 1.75rem;
        font-weight: 800;
        margin-bottom: .25rem;
      }

      .app-subtitle {
        color: rgba(0, 0, 0, 1);
        font-size: .98rem;
      }

      .pill-row {
        display: flex;
        flex-wrap: wrap;
        gap: .45rem;
        margin-top: .9rem;
      }

      .pill {
        background: rgba(255,255,255,.13);
        border: 1px solid rgba(255,255,255,.22);
        color: white;
        border-radius: 999px;
        padding: .3rem .65rem;
        font-size: .82rem;
      }

      .metric-card {
        background: var(--card);
        border: 1px solid var(--border);
        border-radius: 16px;
        padding: .95rem 1rem;
        box-shadow: 0 1px 2px rgba(16, 24, 40, .05);
        height: 100%;
      }

      .metric-label {
        color: var(--text-muted);
        font-size: .82rem;
        font-weight: 600;
        margin-bottom: .35rem;
      }

      .metric-value {
        color: var(--text-main);
        font-size: 1.65rem;
        font-weight: 800;
        line-height: 1.15;
      }

      .metric-help {
        color: var(--text-muted);
        font-size: .75rem;
        margin-top: .25rem;
      }

      .client-card {
        background: var(--card);
        border: 1px solid var(--border);
        border-radius: 18px;
        padding: 1rem 1.15rem;
        box-shadow: 0 1px 2px rgba(16, 24, 40, .05);
        margin-top: .9rem;
        margin-bottom: .9rem;
      }

      .section-title {
        font-size: 1.05rem;
        font-weight: 800;
        color: var(--text-main);
        margin-top: .35rem;
        margin-bottom: .55rem;
      }

      .client-id {
        font-size: 1.35rem;
        font-weight: 850;
        color: var(--text-main);
        margin-bottom: .25rem;
      }

      .client-meta {
        color: var(--text-muted);
        font-size: .92rem;
      }

      .status-ok {
        display: inline-block;
        background: #ecfdf3;
        color: #027a48;
        border: 1px solid #abefc6;
        border-radius: 999px;
        padding: .22rem .55rem;
        font-size: .78rem;
        font-weight: 700;
      }

      .status-warn {
        display: inline-block;
        background: #fffaeb;
        color: #b54708;
        border: 1px solid #fedf89;
        border-radius: 999px;
        padding: .22rem .55rem;
        font-size: .78rem;
        font-weight: 700;
      }

      .image-label {
        font-weight: 800;
        font-size: .98rem;
        color: #101828;
        margin-bottom: .35rem;
      }

      .decision-panel {
        background: #ffffff;
        border: 1px solid var(--border);
        border-radius: 18px;
        padding: 1rem 1.1rem;
        margin-top: .85rem;
        box-shadow: 0 1px 2px rgba(16, 24, 40, .05);
      }

      .decision-group-title {
        font-size: .82rem;
        text-transform: uppercase;
        letter-spacing: .05em;
        color: #667085;
        font-weight: 800;
        margin-bottom: .35rem;
      }

      .footer-note {
        color: #667085;
        font-size: .82rem;
      }

      div[data-testid="stSidebar"] {
        background: #f8fafc;
      }

      div[data-testid="stDataFrame"] {
        border-radius: 12px;
      }

      /* Botones principales y secundarios */
        .stButton button,
        .stDownloadButton button,
        .stLinkButton a {
        background: linear-gradient(var(--control-blue-light)) !important;
        color: #FFFFFF !important;
        border: 1px solid rgba(147, 197, 253, 0.45) !important;
        border-radius: 14px !important;
        min-height: 2.75rem !important;
        font-weight: 850 !important;
        box-shadow: 0 8px 20px rgba(37, 99, 235, 0.22) !important;
        }

        /* Hover de botones */
        .stButton button:hover,
        .stDownloadButton button:hover,
        .stLinkButton a:hover {
        background: linear-gradient(
            135deg,
            #1D4ED8 0%,
            #0284C7 100%
        ) !important;
        color: #FFFFFF !important;
        border-color: rgba(191, 219, 254, 0.75) !important;
        transform: translateY(-1px);
        box-shadow: 0 12px 28px rgba(37, 99, 235, 0.32) !important;
        }

        /* Botón deshabilitado */
        .stButton button:disabled,
        .stDownloadButton button:disabled {
        background: rgba(148, 163, 184, 0.35) !important;
        color: rgba(255, 255, 255, 0.72) !important;
        border-color: rgba(148, 163, 184, 0.35) !important;
        box-shadow: none !important;
        }

        /* Slider: bolita */
        div[data-testid="stSlider"] div[role="slider"] {
        background-color: var(--control-blue) !important;
        border-color: var(--control-blue) !important;
        box-shadow: 0 0 0 4px var(--control-blue-soft) !important;
        }

        /* Slider: números y textos */
        div[data-testid="stSlider"] label,
        div[data-testid="stSlider"] span,
        div[data-testid="stSlider"] p {
        color: var(--app-text) !important;
        }

        /* Toggle activo */
        div[data-testid="stToggle"] button[aria-checked="true"],
        div[data-testid="stToggle"] div[aria-checked="true"] {
        background-color: var(--control-blue) !important;
        border-color: var(--control-blue) !important;
        }

        /* Toggle inactivo */
        div[data-testid="stToggle"] button[aria-checked="false"],
        div[data-testid="stToggle"] div[aria-checked="false"] {
        background-color: var(--control-track) !important;
        border-color: var(--control-track) !important;
        }

        /* Radio seleccionado */
        div[data-testid="stRadio"] [role="radio"][aria-checked="true"] {
        border-color: var(--control-blue) !important;
        background-color: var(--control-blue) !important;
        }

        /* Checkbox seleccionado */
        div[data-testid="stCheckbox"] [aria-checked="true"] {
        border-color: var(--control-blue) !important;
        background-color: var(--control-blue) !important;
        }

        /* Etiquetas seleccionadas del multiselect */
        div[data-baseweb="tag"] {
        background: var(--control-blue) !important;
        color: #FFFFFF !important;
        border-radius: 10px !important;
        }

        div[data-baseweb="tag"] span {
        color: #FFFFFF !important;
        }

        /* Quitar fondo azul raro en textos al seleccionar/focus */
        div[data-testid="stRadio"] label,
        div[data-testid="stRadio"] label *,
        div[data-testid="stToggle"] label,
        div[data-testid="stToggle"] label *,
        div[data-testid="stCheckbox"] label,
        div[data-testid="stCheckbox"] label * {
        background: transparent !important;
        }
                                              
    </style>                                 
    """,
    unsafe_allow_html=True,
)


# ============================================================
# FUNCIONES CON CACHÉ
# ============================================================

@st.cache_data(show_spinner=False)
def procesar_archivo(contenido: bytes):
    return analizar_excel(contenido)


@st.cache_data(show_spinner=False, ttl=1800, max_entries=30)
def cargar_imagen(url: str) -> bytes:
    return obtener_imagen_cacheada(url)


def fmt_numero(valor: int | float) -> str:
    return f"{valor:,}".replace(",", ".")


def fmt_porcentaje(valor: float) -> str:
    return f"{valor:.1f}%".replace(".", ",")


def urls_de_tienda(tienda: dict) -> list[str]:
    return [
        tienda["representante"].get(columna_excel, "")
        for columna_excel, _ in MESES_FOTO
        if tienda["representante"].get(columna_excel, "")
    ]


def render_metric_card(label: str, value: str, help_text: str = ""):
    st.markdown(
        f"""
        <div class="metric-card">
          <div class="metric-label">{label}</div>
          <div class="metric-value">{value}</div>
          <div class="metric-help">{help_text}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def inicializar_estado(contenido: bytes, analisis: dict, nombre_archivo: str):
    if st.session_state.get("hash_archivo") != analisis["hash_archivo"]:
        sesion = cargar_sesion(analisis["hash_archivo"])
        tiendas_base_memoria = analisis.get("tiendas_todas") or analisis.get("tiendas") or []
        ids_disponibles_memoria = {tienda["id_cliente"] for tienda in tiendas_base_memoria}

        decisiones_recuperadas = {
            id_cliente: criterio
            for id_cliente, criterio in sesion.get("decisiones", {}).items()
            if id_cliente in ids_disponibles_memoria and criterio in CRITERIOS
        }

        st.session_state.hash_archivo = analisis["hash_archivo"]
        st.session_state.nombre_archivo = nombre_archivo
        st.session_state.contenido_excel = contenido
        st.session_state.analisis = analisis
        st.session_state.decisiones = decisiones_recuperadas
        st.session_state.sesion_recuperada = sesion
        st.session_state.id_actual = (analisis.get("tiendas") or analisis.get("tiendas_todas") or [{}])[0].get("id_cliente")
        st.session_state.selector_token = 0
        st.session_state.export_bytes = None
        st.session_state.export_count = 0
        st.session_state.ultimo_guardado = sesion.get("ultima_actualizacion") if sesion else None
        st.session_state.mensaje_guardado = ""


def persistir_decisiones(ultimo_id_cliente: str | None = None):
    guardar_sesion(
        st.session_state.hash_archivo,
        st.session_state.nombre_archivo,
        st.session_state.decisiones,
        st.session_state.analisis["total_tiendas"],
        ultimo_id_cliente=ultimo_id_cliente,
    )
    ahora = datetime.now().strftime("%H:%M:%S")
    st.session_state.ultimo_guardado = ahora
    st.session_state.mensaje_guardado = f"Guardado automático: {ultimo_id_cliente or 'registro'} · {ahora}"
    st.session_state.export_bytes = None
    st.session_state.export_count = 0


def seleccionar_siguiente_pendiente(tiendas_en_rango: list[dict]):
    siguientes = [
        tienda for tienda in tiendas_en_rango
        if tienda["id_cliente"] not in st.session_state.decisiones
    ]
    st.session_state.id_actual = siguientes[0]["id_cliente"] if siguientes else None
    st.session_state.selector_token = st.session_state.get("selector_token", 0) + 1


def guardar_decision(criterio: str, tienda_actual: dict, tiendas_en_rango: list[dict], avanzar: bool = True):
    st.session_state.decisiones[tienda_actual["id_cliente"]] = criterio
    persistir_decisiones(ultimo_id_cliente=tienda_actual["id_cliente"])
    if avanzar:
        seleccionar_siguiente_pendiente(tiendas_en_rango)
    else:
        st.session_state.selector_token = st.session_state.get("selector_token", 0) + 1
    st.rerun()


# ============================================================
# CARGA DEL ARCHIVO
# ============================================================

st.markdown(
    """
    <div class="app-header">
      <div class="app-title">Sistema de validación de Tiendas Ancla</div>
      <div class="app-subtitle">Revisión de exhibiciones</div>
    </div>
    """,
    unsafe_allow_html=True,
)

archivo = st.file_uploader(
    "Sube el archivo Excel base o el último Excel recuperado",
    type=["xlsx"],
    help="El sistema leerá la hoja Consolidado y trabajará con las tiendas pendientes que tienen fotos de febrero, marzo y abril.",
)

if archivo is None:
    st.info("Carga el Excel para iniciar la revisión.")
    st.stop()

contenido_excel = archivo.getvalue()

try:
    with st.spinner("Leyendo el archivo y agrupando las tiendas..."):
        analisis = procesar_archivo(contenido_excel)
except Exception as error:
    st.error(f"No fue posible leer el archivo: {error}")
    st.stop()

inicializar_estado(contenido_excel, analisis, archivo.name)
decisiones = st.session_state.decisiones

# ------------------------------------------------------------
# FILTRO INICIAL DE REVISIÓN
# ------------------------------------------------------------

st.markdown('<div class="section-title">Configuración de revisión</div>', unsafe_allow_html=True)

tiendas_todas = analisis.get("tiendas_todas") or analisis.get("tiendas") or []
tiendas_pendientes = analisis.get("tiendas") or []
conteo_colores = analisis.get("conteo_colores", {})

cf1, cf2 = st.columns([1.2, 1])
with cf1:
    modo_revision = st.radio(
        "¿Qué registros quieres revisar?",
        [
            "Pendientes nuevas",
            "Filtrar por color del Excel",
            "Todas las tiendas con fotos",
        ],
        horizontal=True,
        help=(
            "Pendientes nuevas excluye lo que ya tenga OBSERVACIÓN o ANOTACIONES. "
            "Filtrar por color permite revisar registros ya pintados en rojo, verde o amarillo."
        ),
    )

with cf2:
    colores_seleccionados = []
    if modo_revision == "Filtrar por color del Excel":
        colores_seleccionados = st.multiselect(
            "Colores a revisar",
            ["ROJO", "VERDE", "AMARILLO", "MIXTO", "SIN COLOR"],
            default=["ROJO", "VERDE", "AMARILLO"],
            help="El filtro detecta colores aplicados en las filas del Excel entre las columnas A y W.",
        )
    else:
        st.caption(
            "Colores detectados: "
            f"Rojo {conteo_colores.get('ROJO', 0)} · "
            f"Verde {conteo_colores.get('VERDE', 0)} · "
            f"Amarillo {conteo_colores.get('AMARILLO', 0)}"
        )

if modo_revision == "Pendientes nuevas":
    tiendas = tiendas_pendientes
elif modo_revision == "Filtrar por color del Excel":
    colores_set = set(colores_seleccionados)
    tiendas = [
        tienda for tienda in tiendas_todas
        if colores_set and (set(tienda.get("colores_excel", [])) & colores_set or tienda.get("color_excel") in colores_set)
    ]
else:
    tiendas = tiendas_todas

if not tiendas:
    st.warning("No se encontraron tiendas para el filtro seleccionado.")
    st.stop()

ids_revision_archivo = {t["id_cliente"] for t in tiendas}
if st.session_state.get("id_actual") not in ids_revision_archivo:
    st.session_state.id_actual = tiendas[0]["id_cliente"]
    st.session_state.selector_token = st.session_state.get("selector_token", 0) + 1

revisadas = len([id_cliente for id_cliente in decisiones if id_cliente in ids_revision_archivo])
total_revision = len(tiendas)
pendientes = total_revision - revisadas
avance = (revisadas / total_revision * 100) if total_revision else 100


# ============================================================
# MÉTRICAS PRINCIPALES
# ============================================================

m1, m2, m3, m4 = st.columns(4)
with m1:
    render_metric_card("Registros del filtro", fmt_numero(total_revision), modo_revision)
with m2:
    render_metric_card("Validadas omitidas", fmt_numero(analisis["total_tiendas_omitidas"]), "Con OBSERVACIÓN o ANOTACIONES")
with m3:
    render_metric_card("Calificadas ahora", fmt_numero(revisadas), "Guardadas en memoria local")
with m4:
    render_metric_card("Avance actual", fmt_porcentaje(avance), f"Faltan {fmt_numero(pendientes)}")

st.progress(avance / 100)

if st.session_state.get("mensaje_guardado"):
    st.success(st.session_state.mensaje_guardado)

sesion_recuperada = st.session_state.get("sesion_recuperada") or {}
if sesion_recuperada:
    total_recuperadas = len(st.session_state.decisiones)
    ultima = sesion_recuperada.get("ultima_actualizacion", "sin fecha")
    if total_recuperadas:
        st.info(f"Memoria recuperada: {total_recuperadas} decisiones restauradas. Última actualización: {ultima}.")
    else:
        st.info("Se encontró una memoria anterior, pero sus decisiones ya estaban escritas en el Excel o no correspondían a tiendas pendientes.")

if analisis["total_tiendas_omitidas"]:
    st.caption(
        f'Se omitieron {analisis["total_tiendas_omitidas"]} tiendas porque ya tienen contenido en OBSERVACIÓN o ANOTACIONES.'
    )


# ============================================================
# BARRA LATERAL
# ============================================================

with st.sidebar:
    st.header("Panel de control")

    st.subheader("Archivo")
    st.caption(f"**Archivo:** {archivo.name}")
    st.caption(f"**Hash:** {analisis['hash_archivo'][:12]}...")
    if st.session_state.get("ultimo_guardado"):
        st.caption(f"**Último guardado:** {st.session_state.ultimo_guardado}")

    st.divider()

    st.subheader("Navegación")
    solo_pendientes = st.toggle("Mostrar solo pendientes de esta sesión", value=True)
    precargar = st.toggle("Precargar próximas imágenes", value=False)

    indice_inicio, indice_fin = st.slider(
        "Rango de tiendas a revisar",
        min_value=1,
        max_value=total_revision,
        value=(1, total_revision),
    )

    tiendas_en_rango = tiendas[indice_inicio - 1:indice_fin]

    if solo_pendientes:
        disponibles = [tienda for tienda in tiendas_en_rango if tienda["id_cliente"] not in decisiones]
    else:
        disponibles = tiendas_en_rango

    if not disponibles:
        st.success("No quedan tiendas pendientes en el rango seleccionado.")
        disponibles = tiendas_en_rango

    opciones = {
        f'{tienda["id_cliente"]} | {tienda["nombre_cliente"]}': tienda
        for tienda in disponibles
    }

    etiquetas = list(opciones.keys())
    etiqueta_por_id = {tienda["id_cliente"]: etiqueta for etiqueta, tienda in opciones.items()}
    id_actual = st.session_state.get("id_actual")

    indice_selector = (
        etiquetas.index(etiqueta_por_id[id_actual])
        if id_actual in etiqueta_por_id
        else 0
    )

    seleccion_label = st.selectbox(
        "Buscar o seleccionar tienda",
        options=etiquetas,
        index=indice_selector,
        key=f'selector_tienda_{st.session_state.get("selector_token", 0)}',
    )

    tienda_actual = opciones[seleccion_label]
    st.session_state.id_actual = tienda_actual["id_cliente"]

    st.divider()

    st.subheader("Memoria local")
    st.caption("En Streamlit Cloud la memoria local es de apoyo. Para máxima seguridad, descarga y conserva el respaldo JSON.")

    respaldo_subido = st.file_uploader("Cargar respaldo JSON", type=["json"], key="respaldo_json_upload")
    if respaldo_subido is not None:
        try:
            datos_respaldo = json.loads(respaldo_subido.getvalue().decode("utf-8"))
            decisiones_json = datos_respaldo.get("decisiones", {})
            ids_validos = {tienda["id_cliente"] for tienda in (analisis.get("tiendas_todas") or analisis.get("tiendas") or [])}
            decisiones_importadas = {
                id_cliente: criterio
                for id_cliente, criterio in decisiones_json.items()
                if id_cliente in ids_validos and criterio in CRITERIOS
            }
            if decisiones_importadas:
                st.session_state.decisiones.update(decisiones_importadas)
                persistir_decisiones(ultimo_id_cliente="respaldo_json")
                st.success(f"Se importaron {len(decisiones_importadas)} decisiones del respaldo.")
                st.rerun()
            else:
                st.warning("El respaldo no contiene decisiones válidas para este Excel.")
        except Exception as error:
            st.error(f"No fue posible leer el respaldo JSON: {error}")

    if decisiones:
        respaldo_json = preparar_respaldo_json(
            st.session_state.hash_archivo,
            st.session_state.nombre_archivo,
            decisiones,
        )
        st.download_button(
            "Descargar respaldo JSON",
            data=respaldo_json,
            file_name=Path(archivo.name).stem + "_respaldo_decisiones.json",
            mime="application/json",
            use_container_width=True,
        )

    with st.expander("Borrar memoria de este Excel"):
        confirmar_borrado = st.checkbox("Confirmo que quiero borrar la memoria local")
        if st.button("Borrar memoria local", disabled=not confirmar_borrado, use_container_width=True):
            borrar_sesion(st.session_state.hash_archivo)
            st.session_state.decisiones = {}
            st.session_state.sesion_recuperada = {}
            st.session_state.export_bytes = None
            st.session_state.export_count = 0
            st.session_state.selector_token = st.session_state.get("selector_token", 0) + 1
            st.rerun()

    st.divider()

    st.subheader("Exportación")
    st.caption("El Excel se genera solo cuando presionas el botón para mejorar rendimiento.")

    nombre_salida = Path(archivo.name).stem + "_revisado.xlsx"

    if decisiones:
        if st.button("Preparar Excel actualizado", type="primary", use_container_width=True):
            try:
                with st.spinner("Generando Excel actualizado..."):
                    st.session_state.export_bytes = exportar_excel(
                        st.session_state.contenido_excel,
                        analisis,
                        decisiones,
                    )
                    st.session_state.export_count = len(decisiones)
            except Exception as error:
                st.error(f"No fue posible preparar la descarga: {error}")

        if st.session_state.get("export_bytes"):
            if st.session_state.get("export_count") == len(decisiones):
                st.download_button(
                    "Descargar Excel actualizado",
                    data=st.session_state.export_bytes,
                    file_name=nombre_salida,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
            else:
                st.warning("Hay nuevas decisiones después de la última generación. Prepara el Excel nuevamente.")
    else:
        st.caption("Aún no hay decisiones nuevas para exportar.")

    st.divider()

    st.subheader("Rendimiento")
    cache_info = estadisticas_cache()
    mb_cache = cache_info["bytes"] / (1024 * 1024)
    st.caption(f'Imágenes en caché: {cache_info["archivos"]}/{cache_info.get("limite_archivos", "?")} archivos · {mb_cache:.1f}/{cache_info.get("limite_mb", "?")} MB')


# ============================================================
# TARJETA DEL CLIENTE
# ============================================================

decision_actual = decisiones.get(tienda_actual["id_cliente"])
estado_html = (
    f'<span class="status-ok">Ya calificada: {decision_actual}</span>'
    if decision_actual
    else '<span class="status-warn">Pendiente de calificación</span>'
)

st.markdown(
    f"""
    <div class="client-card">
      <div style="display:flex; justify-content:space-between; gap:1rem; flex-wrap:wrap;">
        <div>
          <div class="client-id">IDCliente: {tienda_actual["id_cliente"]}</div>
          <div class="client-meta">
            <b>Cliente:</b> {tienda_actual["nombre_cliente"] or "Sin nombre"} ·
            <b>Regional:</b> {tienda_actual["regional"] or "Sin regional"} ·
            <b>Filas Excel:</b> {", ".join(str(fila) for fila in tienda_actual["filas_excel"])} ·
            <b>Color Excel:</b> {tienda_actual.get("color_excel", "SIN COLOR")}
          </div>
        </div>
        <div>{estado_html}</div>
      </div>
    </div>
    """,
    unsafe_allow_html=True,
)


if tienda_actual.get("observaciones_previas") or tienda_actual.get("anotaciones_previas"):
    st.caption(
        "Validación previa detectada: "
        f"OBSERVACIÓN={'; '.join(tienda_actual.get('observaciones_previas', [])) or 'vacío'} · "
        f"ANOTACIONES={'; '.join(tienda_actual.get('anotaciones_previas', [])) or 'vacío'}"
    )

# ============================================================
# REGISTROS ASOCIADOS
# ============================================================

with st.expander("Ver filas asociadas del Excel"):
    columnas_mostrar = [
        "Fila Excel",
        "Color Excel",
        "Medición",
        "IDCliente",
        "Nombre Cliente 1",
        "Nombre Cliente 2",
        "NRegional",
        "NSubcanal",
        "OBSERVACIÓN:  NESTUM NO SE P",
        "ANOTACIONES",
    ]

    tabla = []
    for fila in tienda_actual["filas_tabla"]:
        fila_visible = {columna: fila.get(columna, "") for columna in columnas_mostrar}
        tabla.append(fila_visible)

    st.dataframe(tabla, use_container_width=True, hide_index=True)


# ============================================================
# FOTOGRAFÍAS
# ============================================================

st.markdown('<div class="section-title">Fotografías de validación</div>', unsafe_allow_html=True)

columnas_imagen = st.columns(3)
for columna_visual, (columna_excel, mes) in zip(columnas_imagen, MESES_FOTO):
    url = tienda_actual["representante"].get(columna_excel, "")
    with columna_visual:
        st.markdown(f'<div class="image-label">{mes}</div>', unsafe_allow_html=True)
        try:
            imagen = cargar_imagen(url)
            st.image(imagen, use_container_width=True)
        except Exception as error:
            st.warning("No fue posible cargar la imagen en vista previa.")
            st.caption(str(error))
        st.link_button(f"Abrir foto de {mes}", url, use_container_width=True)

if precargar:
    try:
        ids_disponibles = [tienda["id_cliente"] for tienda in disponibles]
        posicion_actual = ids_disponibles.index(tienda_actual["id_cliente"])
        siguientes_precarga = disponibles[posicion_actual + 1: posicion_actual + 3]
        urls_precarga = []
        for tienda in siguientes_precarga:
            urls_precarga.extend(urls_de_tienda(tienda))
        if urls_precarga:
            precargar_imagenes(urls_precarga, max_workers=2, max_urls=3)
    except Exception:
        pass


# ============================================================
# DECISIONES RÁPIDAS
# ============================================================

st.markdown(
    """
    <div class="decision-panel">
      <div class="section-title">Decisión rápida</div>
      <div class="footer-note">
        Al presionar un criterio, la decisión se guarda en memoria local y la aplicación avanza a la siguiente tienda pendiente.
      </div>
    </div>
    """,
    unsafe_allow_html=True,
)

st.markdown('<div class="decision-group-title">Aprobación</div>', unsafe_allow_html=True)
c_gana, c_guardar_sin, c_quitar = st.columns([1, 1, 1])
with c_gana:
    if st.button("GANA", type="primary", use_container_width=True):
        guardar_decision("GANA", tienda_actual, tiendas_en_rango, avanzar=True)

with c_guardar_sin:
    if st.button("Guardar GANA sin avanzar", use_container_width=True):
        guardar_decision("GANA", tienda_actual, tiendas_en_rango, avanzar=False)

with c_quitar:
    if st.button("Quitar clasificación", use_container_width=True):
        decisiones.pop(tienda_actual["id_cliente"], None)
        st.session_state.decisiones = decisiones
        persistir_decisiones(ultimo_id_cliente=tienda_actual["id_cliente"])
        st.session_state.selector_token = st.session_state.get("selector_token", 0) + 1
        st.rerun()

st.markdown('<div class="decision-group-title">Rechazo / No gana</div>', unsafe_allow_html=True)
r1, r2, r3 = st.columns(3)
with r1:
    if st.button("NO GANA", use_container_width=True):
        guardar_decision("NO GANA", tienda_actual, tiendas_en_rango, avanzar=True)
with r2:
    if st.button("NESTUM EN LA EXHIBICION", use_container_width=True):
        guardar_decision("NESTUM EN LA EXHIBICION", tienda_actual, tiendas_en_rango, avanzar=True)
with r3:
    if st.button("NESTOGENO EN LA EXHIBICION", use_container_width=True):
        guardar_decision("NESTOGENO EN LA EXHIBICION", tienda_actual, tiendas_en_rango, avanzar=True)

r4, r5, _ = st.columns(3)
with r4:
    if st.button("NESTUM Y NESTOGENO", use_container_width=True):
        guardar_decision("NESTUM Y NESTOGENO EN LA EXHIBICION", tienda_actual, tiendas_en_rango, avanzar=True)
with r5:
    if st.button("NO TIENE 3 O MÁS CATEGORÍAS", use_container_width=True):
        guardar_decision("NO TIENE 3 O MAS CATEGORIAS", tienda_actual, tiendas_en_rango, avanzar=True)

st.markdown('<div class="decision-group-title">Enviar a revisión</div>', unsafe_allow_html=True)
a1, a2, _ = st.columns(3)
with a1:
    if st.button("POCOS PRODUCTOS", use_container_width=True):
        guardar_decision("POCOS PRODUCTOS", tienda_actual, tiendas_en_rango, avanzar=True)


with st.expander("Modo corrección manual"):
    criterios = list(CRITERIOS.keys())
    indice_inicial = criterios.index(decision_actual) if decision_actual in criterios else 0

    criterio_manual = st.selectbox(
        "Selecciona un criterio y decide si quieres avanzar o no:",
        criterios,
        index=indice_inicial,
    )

    cm1, cm2 = st.columns(2)
    with cm1:
        if st.button("Guardar criterio manual y avanzar", use_container_width=True):
            guardar_decision(criterio_manual, tienda_actual, tiendas_en_rango, avanzar=True)
    with cm2:
        if st.button("Guardar criterio manual sin avanzar", use_container_width=True):
            guardar_decision(criterio_manual, tienda_actual, tiendas_en_rango, avanzar=False)


# ============================================================
# INFORMACIÓN FINAL
# ============================================================

with st.expander("¿Qué se modifica al descargar el Excel?"):
    st.markdown(
        """
        - Se conserva el archivo cargado y se genera una **copia nueva**.
        - La decisión seleccionada se escribe en la columna **W — ANOTACIONES**.
        - Las tres filas del mismo **IDCliente** se colorean desde la columna **A hasta W**.
        - La columna **V — OBSERVACIÓN** se mantiene sin reemplazar.
        - Las demás hojas, fórmulas y enlaces permanecen en el archivo exportado.
        - Las decisiones se guardan automáticamente en `data/sesiones`.
        - Las imágenes descargadas se guardan de forma limitada en `data/cache_imagenes` para no saturar Streamlit Cloud.
        - Puedes filtrar por registros pintados en rojo, verde o amarillo desde la configuración de revisión.
        """
    )
