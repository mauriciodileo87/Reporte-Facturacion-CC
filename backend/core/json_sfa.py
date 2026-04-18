# json_sfa.py
# Lógica para leer el archivo JSON/TXT SFA y armar el resumen.

import json
import re
from utils_sfa import normalizar_codigo

# Ejemplo de clave:
# "0005@0000001432@Z005": { "cnt": 867, "importe": 816800000
PATRON_SFA = (
    r'"(\d{4})@(\d{10})@([A-Z0-9]{4})"\s*:\s*{\s*"cnt"\s*:\s*\d+\s*,\s*"importe"\s*:\s*(\d+)'
)
REGEX_SFA = re.compile(PATRON_SFA)


def _agregar_item(resumen: dict, clave_compuesta: str, importe_centavos: int) -> None:
    """Agrega un item al resumen final respetando el formato esperado."""
    partes = clave_compuesta.split("@")
    if len(partes) != 3:
        return

    juego_raw, sorteo_raw, concepto_raw = partes
    juego = normalizar_codigo(juego_raw)
    sorteo = normalizar_codigo(sorteo_raw)
    concepto = concepto_raw.strip().upper()

    if not sorteo or not concepto:
        return

    clave = (juego, sorteo, concepto)
    resumen[clave] = resumen.get(clave, 0.0) + (importe_centavos / 100.0)


def _leer_json_estructurado(contenido: str) -> dict | None:
    """Intenta parsear el contenido como JSON real."""
    try:
        data = json.loads(contenido)
    except json.JSONDecodeError:
        return None

    if not isinstance(data, dict):
        return None

    resumen = {}
    for clave_compuesta, payload in data.items():
        if not isinstance(payload, dict):
            continue

        importe = payload.get("importe")
        if isinstance(importe, str):
            if not importe.isdigit():
                continue
            importe_centavos = int(importe)
        elif isinstance(importe, int):
            importe_centavos = importe
        else:
            continue

        _agregar_item(resumen, str(clave_compuesta), importe_centavos)

    return resumen


def _leer_por_regex(contenido: str) -> dict:
    """Fallback para archivos no JSON estricto (formato histórico)."""
    resumen = {}

    for m in REGEX_SFA.finditer(contenido):
        juego_raw = m.group(1)
        sorteo_raw = m.group(2)
        concepto_raw = m.group(3)
        importe_centavos_str = m.group(4)

        try:
            importe_centavos = int(importe_centavos_str)
        except ValueError:
            continue

        _agregar_item(resumen, f"{juego_raw}@{sorteo_raw}@{concepto_raw}", importe_centavos)

    return resumen


def leer_json_sfa(ruta_json: str) -> dict:
    """
    Lee el JSON/TXT SFA y devuelve un dict:
    { (juego, sorteo, concepto): importe_en_pesos }
    """
    with open(ruta_json, "r", encoding="utf-8", errors="replace") as f:
        contenido = f.read()

    resumen = _leer_json_estructurado(contenido)
    if resumen is not None:
        return resumen

    return _leer_por_regex(contenido)
