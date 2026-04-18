# utils_sfa.py
# Funciones comunes para SFA / Reporte.

def normalizar_codigo(cadena: str) -> str:
    """
    Quita espacios y ceros a la izquierda si es numérico.
    '0005' -> '5', '0000001432' -> '1432'
    Si no se puede convertir a int, devuelve el string recortado.
    """
    cadena = cadena.strip()
    if cadena == "":
        return ""
    try:
        return str(int(cadena))
    except ValueError:
        return cadena


def formato_pesos(importe: float) -> str:
    """
    Devuelve el importe formateado con símbolo de pesos.
    Ej: 1234.5 -> '$ 1,234.50'
    """
    return f"$ {importe:,.2f}"


def ordenar_resumen(resumen: dict):
    """
    Ordena un dict de resúmenes:
    { (juego, sorteo, concepto): importe }
    y devuelve lista de items ordenados.
    """
    def orden_clave(item):
        (juego, sorteo, concepto), _importe = item
        try:
            jn = int(juego)
        except ValueError:
            jn = 999999
        try:
            sn = int(sorteo)
        except ValueError:
            sn = 999999
        return (jn, sn, concepto)

    return sorted(resumen.items(), key=orden_clave)


def filtrar_resumen(resumen: dict, filtros: dict):
    """
    Aplica filtros por columnas sobre un resumen.
    filtros: dict con claves posibles:
        "juego", "sorteo", "concepto", "importe"
    El filtro es "contiene" (no exacto).
    """
    if not filtros:
        return resumen.copy()

    filtro_juego = filtros.get("juego", "").strip()
    filtro_sorteo = filtros.get("sorteo", "").strip()
    filtro_concepto = filtros.get("concepto", "").strip().upper()
    filtro_importe = filtros.get("importe", "").strip()

    filtrado = {}

    for (j, s, c), imp in resumen.items():
        if filtro_juego and filtro_juego not in j:
            continue

        if filtro_sorteo and filtro_sorteo not in s:
            continue

        if filtro_concepto and filtro_concepto not in c.upper():
            continue

        if filtro_importe:
            imp_str = formato_pesos(imp)
            if filtro_importe not in imp_str:
                continue

        filtrado[(j, s, c)] = imp

    return filtrado
