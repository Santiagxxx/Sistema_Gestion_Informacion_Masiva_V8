from __future__ import annotations

import hashlib
import os
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from typing import Iterable

import requests

BASE_DIR = Path(__file__).resolve().parent
CACHE_DIR = BASE_DIR / "data" / "cache_imagenes"
CACHE_DIR.mkdir(parents=True, exist_ok=True)

# Límites pensados para Streamlit Community Cloud.
# Evitan que el caché local crezca demasiado si se revisan muchas tiendas.
MAX_CACHE_FILES = 90
MAX_CACHE_MB = 350


def _extension_desde_url(url: str) -> str:
    limpio = (url or "").split("?")[0].lower()
    for extension in (".jpg", ".jpeg", ".png", ".webp", ".gif"):
        if limpio.endswith(extension):
            return extension
    return ".img"


def ruta_imagen(url: str) -> Path:
    nombre = hashlib.sha256((url or "").encode("utf-8")).hexdigest()
    return CACHE_DIR / f"{nombre}{_extension_desde_url(url)}"


def _podar_cache() -> None:
    archivos = [ruta for ruta in CACHE_DIR.glob("*") if ruta.is_file() and not ruta.name.endswith(".tmp")]
    if not archivos:
        return

    total_bytes = sum(ruta.stat().st_size for ruta in archivos)
    limite_bytes = MAX_CACHE_MB * 1024 * 1024

    if len(archivos) <= MAX_CACHE_FILES and total_bytes <= limite_bytes:
        return

    archivos.sort(key=lambda ruta: ruta.stat().st_mtime)
    while archivos and (len(archivos) > MAX_CACHE_FILES or total_bytes > limite_bytes):
        ruta = archivos.pop(0)
        try:
            size = ruta.stat().st_size
            ruta.unlink()
            total_bytes -= size
        except FileNotFoundError:
            pass


def obtener_imagen_cacheada(url: str, timeout: int = 20) -> bytes:
    if not url:
        raise ValueError("La URL de la imagen está vacía.")

    ruta = ruta_imagen(url)
    if ruta.exists() and ruta.stat().st_size > 0:
        try:
            os.utime(ruta, None)
        except Exception:
            pass
        return ruta.read_bytes()

    respuesta = requests.get(url, timeout=timeout)
    respuesta.raise_for_status()
    contenido = respuesta.content

    ruta_temporal = ruta.with_suffix(ruta.suffix + ".tmp")
    with ruta_temporal.open("wb") as archivo:
        archivo.write(contenido)
    os.replace(ruta_temporal, ruta)
    _podar_cache()

    return contenido


def precargar_imagenes(urls: Iterable[str], max_workers: int = 2, max_urls: int = 3) -> dict[str, str]:
    urls_limpias = [url for url in dict.fromkeys(urls) if url][:max_urls]
    resultado: dict[str, str] = {}

    pendientes = [
        url for url in urls_limpias
        if not (ruta_imagen(url).exists() and ruta_imagen(url).stat().st_size > 0)
    ]

    if not pendientes:
        return {url: "cache" for url in urls_limpias}

    with ThreadPoolExecutor(max_workers=max_workers) as ejecutor:
        futuros = {ejecutor.submit(obtener_imagen_cacheada, url): url for url in pendientes}
        for futuro in as_completed(futuros):
            url = futuros[futuro]
            try:
                futuro.result()
                resultado[url] = "descargada"
            except Exception as error:
                resultado[url] = f"error: {error}"

    for url in urls_limpias:
        resultado.setdefault(url, "cache")

    return resultado


def estadisticas_cache() -> dict[str, int]:
    archivos = [ruta for ruta in CACHE_DIR.glob("*") if ruta.is_file() and not ruta.name.endswith(".tmp")]
    return {
        "archivos": len(archivos),
        "bytes": sum(ruta.stat().st_size for ruta in archivos),
        "limite_archivos": MAX_CACHE_FILES,
        "limite_mb": MAX_CACHE_MB,
    }
