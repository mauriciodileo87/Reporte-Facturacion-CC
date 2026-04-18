# facuni_sfa.py
# Parser FACUNI TXT ancho fijo + resumen con nombres de concepto

ANCHO_COLUMNAS = [2, 5, 1, 4, 10, 10, 8, 1, 15, 4, 8]

# ✅ Mapeo de códigos -> nombres (tomado de tu VBA)
MAPEO_CONCEPTOS = {
    34: "TELEKINO",
    750: "Dev. tickets mal imp.",
    632: "Sobrante de caja",
    630: "Seg. med. c/o coop.",
    118: "HIPODROMO LA PLATA",
    117: "HIPODROMO SAN ISIDRO",
    102: "HIPODROMO PALERMO",
    622: "Loto Match",
    467: "Quini 6 segunda",
    70: "La Grande De La Ciudad",
    742: "Quini 6",
    32: "Quiniela de la ciudad",
    42: "Loto Plus tradicional",
    617: "Loto 5 plus",
    677: "Brinco",
    138: "Comision posnet",
    636: "Seg. integr c/o coop.",
    910: "LAS VEGAS",
    544: "Integr. capital coop",
    452: "Quiniela Poceada de la ciudad",
    491: "Multa",
}


def _split_fixed_width(linea: str, anchos: list[int]) -> list[str]:
    out = []
    i = 0
    for w in anchos:
        out.append(linea[i:i + w])
        i += w
    return out


def _solo_digitos(s) -> str:
    s = "" if s is None else str(s)
    return "".join(ch for ch in s if ch.isdigit())


def _parse_importe(valor_raw) -> float | None:
    """
    VBA:
    - últimos 2 dígitos = decimales
    - EXACTAMENTE 2 decimales
    - soporta '-' al inicio o al final (trailing minus)
    - si supera 15 dígitos, toma los últimos 15
    """
    if valor_raw is None:
        return None

    txt = str(valor_raw).strip()
    if not txt:
        return None

    es_neg = False
    if txt.startswith("-"):
        es_neg = True
        txt = txt[1:].strip()
    if txt.endswith("-"):
        es_neg = True
        txt = txt[:-1].strip()

    txt = _solo_digitos(txt)
    if not txt:
        return None

    if len(txt) > 15:
        txt = txt[-15:]

    num = int(txt) / 100.0
    if es_neg:
        num = -num

    return round(num, 2)


def _parse_concepto_cta_cte(valor_raw) -> int | None:
    s = _solo_digitos(valor_raw)
    if not s:
        return None
    try:
        return int(s)
    except ValueError:
        return None


def _nombre_concepto(codigo: int) -> str:
    """
    Devuelve el nombre si está mapeado.
    Si no está, devuelve "CODIGO - SIN NOMBRE" para que lo detectes fácil.
    """
    return MAPEO_CONCEPTOS.get(codigo, f"{codigo} - SIN NOMBRE")


def importar_facuni_txt(path_txt: str):
    """
    Devuelve:
      detalle: list[dict] (NO lo vas a mostrar, pero lo devolvemos por compatibilidad)
      resumen: list[dict] con Concepto (nombre), Importe, Tipo
      total_afectaciones: float
    """

    # === Detalle (interno) ===
    detalle = []
    with open(path_txt, "r", encoding="utf-8", errors="replace") as f:
        for linea in f:
            linea = linea.rstrip("\n")
            if not linea.strip():
                continue

            cols = _split_fixed_width(linea, ANCHO_COLUMNAS)

            provincia = (cols[0] or "").strip()            # A
            concepto_sap = (cols[3] or "").strip()         # D
            tipo = (cols[7] or "").strip().upper()         # H (D/C)
            importe = _parse_importe(cols[8])              # I
            concepto_cta_cte = _parse_concepto_cta_cte(cols[9])  # J

            detalle.append({
                "Provincia": provincia,
                "Concepto SAP": concepto_sap,
                "Tipo": tipo,
                "Importe": importe,
                "Concepto Cta Cte": concepto_cta_cte,
            })

    # === Resumen (SUM(D) - SUM(C)) ===
    acumulado = {}
    for fila in detalle:
        codigo = fila.get("Concepto Cta Cte")
        tipo = (fila.get("Tipo") or "").strip().upper()
        importe = fila.get("Importe")

        if codigo is None or importe is None:
            continue

        acumulado.setdefault(codigo, 0.0)

        if tipo == "D":
            acumulado[codigo] += float(importe)
        elif tipo == "C":
            acumulado[codigo] -= float(importe)

    resumen = []
    for codigo in sorted(acumulado.keys()):
        imp = round(acumulado[codigo], 2)
        if imp < 0:
            t = "C"
        elif imp > 0:
            t = "D"
        else:
            t = ""

        resumen.append({
            "Concepto": _nombre_concepto(codigo),  # ✅ ahora sale el NOMBRE
            "Importe": imp,
            "Tipo": t
        })

    total_afectaciones = round(sum(r["Importe"] for r in resumen), 2) if resumen else 0.0
    return detalle, resumen, total_afectaciones