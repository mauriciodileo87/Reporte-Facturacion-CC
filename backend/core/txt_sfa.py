# txt_sfa.py
# Lógica para leer el archivo TXT SFA y armar el resumen.

from utils_sfa import normalizar_codigo


def leer_txt_sfa(ruta_txt: str) -> dict:
    """
    Lee el TXT SFA y devuelve un dict:
    { (juego, sorteo, concepto): importe_en_pesos }

    Usa el layout:
    offsets: 0,8,12,22,30,34,35,50,58,66,70,78,88
    """
    resumen = {}

    with open(ruta_txt, "r", encoding="latin-1", errors="replace") as f:
        for linea in f:
            if len(linea) < 50:
                continue

            # campos
            juego = linea[8:12]
            sorteo = linea[12:22]
            concepto = linea[30:34]
            tipo_mov = linea[34:35].strip().upper()  # D / C
            importe_str = linea[35:50].strip()       # 15 dígitos, centavos

            if not importe_str.isdigit():
                continue

            importe = int(importe_str) / 100.0  # a pesos

            if tipo_mov not in ("D", "C"):
                continue

            juego_norm = normalizar_codigo(juego)
            sorteo_norm = normalizar_codigo(sorteo)
            concepto_norm = concepto.strip().upper()
            if sorteo_norm == "" or concepto_norm == "":
                continue

            clave = (juego_norm, sorteo_norm, concepto_norm)
            resumen[clave] = resumen.get(clave, 0.0) + importe

    return resumen