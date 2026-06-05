from __future__ import annotations

import json
import os
from datetime import datetime
from pathlib import Path
from typing import Any

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
SESIONES_DIR = DATA_DIR / "sesiones"
BACKUPS_DIR = DATA_DIR / "backups"

for carpeta in (DATA_DIR, SESIONES_DIR, BACKUPS_DIR):
    carpeta.mkdir(parents=True, exist_ok=True)


def _hash_corto(hash_archivo: str) -> str:
    return (hash_archivo or "sin_hash")[:16]


def ruta_sesion(hash_archivo: str) -> Path:
    return SESIONES_DIR / f"{_hash_corto(hash_archivo)}_decisiones.json"


def cargar_sesion(hash_archivo: str) -> dict[str, Any]:
    ruta = ruta_sesion(hash_archivo)
    if not ruta.exists():
        return {}

    try:
        with ruta.open("r", encoding="utf-8") as archivo:
            datos = json.load(archivo)
        if not isinstance(datos, dict):
            return {}
        if not isinstance(datos.get("decisiones", {}), dict):
            datos["decisiones"] = {}
        return datos
    except Exception:
        return {}


def guardar_sesion(
    hash_archivo: str,
    nombre_archivo: str,
    decisiones: dict[str, str],
    total_pendientes_archivo: int,
    ultimo_id_cliente: str | None = None,
) -> Path:
    ruta = ruta_sesion(hash_archivo)
    ruta_temporal = ruta.with_suffix(".tmp")

    sesion_anterior = cargar_sesion(hash_archivo)
    fecha_inicio = sesion_anterior.get("fecha_inicio") or datetime.now().isoformat(timespec="seconds")

    datos = {
        "version_memoria": "v6",
        "hash_archivo": hash_archivo,
        "nombre_archivo": nombre_archivo,
        "fecha_inicio": fecha_inicio,
        "ultima_actualizacion": datetime.now().isoformat(timespec="seconds"),
        "total_pendientes_archivo": total_pendientes_archivo,
        "total_decisiones": len(decisiones),
        "ultimo_id_cliente": ultimo_id_cliente,
        "decisiones": dict(decisiones),
    }

    with ruta_temporal.open("w", encoding="utf-8") as archivo:
        json.dump(datos, archivo, ensure_ascii=False, indent=2)

    os.replace(ruta_temporal, ruta)
    return ruta


def borrar_sesion(hash_archivo: str) -> bool:
    ruta = ruta_sesion(hash_archivo)
    if ruta.exists():
        ruta.unlink()
        return True
    return False


def preparar_respaldo_json(hash_archivo: str, nombre_archivo: str, decisiones: dict[str, str]) -> bytes:
    datos = {
        "version_memoria": "v6",
        "hash_archivo": hash_archivo,
        "nombre_archivo": nombre_archivo,
        "fecha_descarga_respaldo": datetime.now().isoformat(timespec="seconds"),
        "total_decisiones": len(decisiones),
        "decisiones": dict(decisiones),
    }
    return json.dumps(datos, ensure_ascii=False, indent=2).encode("utf-8")
