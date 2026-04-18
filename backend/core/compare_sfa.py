# compare_sfa.py
# Lógica para comparar resumen TXT vs JSON.

def comparar_resumenes(txt_resumen: dict, json_resumen: dict, tolerancia=0.01):
    """
    Compara dos resúmenes:
    txt_resumen, json_resumen: { (juego, sorteo, concepto): importe }

    Devuelve lista de diferencias:
    [(juego, sorteo, concepto, importe_txt_abs, importe_json), ...]

    Sólo incluye diferencias con |txt - json| > tolerancia.
    """
    diferencias = []
    todas = set(txt_resumen.keys()) | set(json_resumen.keys())

    for clave in todas:
        imp_txt = abs(txt_resumen.get(clave, 0.0))   # usamos valor absoluto como en Excel
        imp_json = json_resumen.get(clave, 0.0)
        diff = round(imp_txt - imp_json, 2)

        if abs(diff) > tolerancia:
            juego, sorteo, concepto = clave
            diferencias.append((juego, sorteo, concepto, imp_txt, imp_json))

    # Ordenar por juego, sorteo, concepto
    def orden(item):
        j, s, c, _, _ = item
        try:
            jn = int(j)
        except ValueError:
            jn = 999999
        try:
            sn = int(s)
        except ValueError:
            sn = 999999
        return (jn, sn, c)

    diferencias.sort(key=orden)
    return diferencias