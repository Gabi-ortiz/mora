"""Lee el PDF "Consulta del Mayor" de FBS (Mayor.rpt) y devuelve los asientos de la cuenta.

Las columnas se identifican por posición horizontal (el PDF no trae tabla):
Asiento | Fecha | Referencia | Comentario | Debe | Haber | Saldo | Centro de Costo
"""
import re
from collections import Counter
from datetime import datetime

import pymupdf

X_REF, X_COM, X_IMPORTES, X_CC = 105, 190, 316, 512
X_DEBE_HABER = 381   # fin de la columna Debe (~379); Haber termina ~433
X_SALDO = 436


def _num(t):
    t = t.strip()
    neg = t.startswith("(") or t.endswith(")")
    v = float(t.strip("()").replace(".", "").replace(",", "."))
    return -v if neg else v


def _lineas(page):
    for b in page.get_text("dict")["blocks"]:
        for l in b.get("lines", []):
            t = "".join(s["text"] for s in l["spans"]).strip()
            if t:
                yield l["bbox"], t


def leer_mayor(path):
    doc = pymupdf.open(path)
    info, asientos = {}, []
    for page in doc:
        lineas = sorted(_lineas(page), key=lambda x: (x[0][1], x[0][0]))
        inicios = [(bb[1], t) for bb, t in lineas if bb[0] < 30 and re.fullmatch(r"\d{6,8}", t)]
        filas = [{"asiento": t, "y": y, "ref": [], "com": [], "cc": []} for y, t in inicios]
        for bb, t in lineas:
            x0, y, x1 = bb[0], bb[1], bb[2]
            if y > page.rect.height - 30:  # pie de página (Mayor.rpt / Hoja n / m)
                continue
            if t.startswith("Cuenta:"):
                info["cuenta"] = t.split(":", 1)[1].strip()
            if not filas or y < filas[0]["y"] - 3:
                continue
            f = max((f for f in filas if f["y"] - 3 <= y), key=lambda f: f["y"])
            if 65 < x0 < X_REF:
                f["fecha"] = datetime.strptime(t, "%d/%m/%y").date()
            elif X_REF <= x0 < X_COM - 5:
                f["ref"].append(t)
            elif X_COM - 5 <= x0 < X_IMPORTES:
                f["com"].append(t)
            elif x0 >= X_IMPORTES and x1 < X_SALDO and re.fullmatch(r"[\d.,]+", t):
                f["debe" if x1 < X_DEBE_HABER else "haber"] = _num(t)
            elif x1 >= X_SALDO and x0 < X_CC and re.fullmatch(r"\(?[\d.,]+\)?", t):
                f["saldo"] = _num(t)
            elif x0 >= X_CC:
                f["cc"].append(t)
        for f in filas:
            asientos.append({
                "asiento": f["asiento"], "fecha": f.get("fecha"), "referencia": " ".join(f["ref"]),
                "comentario": re.sub(r"\s+", " ", " ".join(f["com"])).strip(),
                "debe": f.get("debe", 0.0), "haber": f.get("haber", 0.0), "saldo": f.get("saldo"),
                "centro_costo": " ".join(f["cc"]),
            })
    # El PDF a veces recorta el primer dígito del saldo: se toma el saldo inicial más repetido
    # (saldo de la fila menos el acumulado) y se recalcula el saldo corrido.
    acum, candidatos = 0.0, Counter()
    for a in asientos:
        acum += a["debe"] - a["haber"]
        if a["saldo"] is not None:
            candidatos[round(a["saldo"] - acum, 2)] += 1
    saldo = info["saldo_inicial"] = candidatos.most_common(1)[0][0] if candidatos else 0.0
    for a in asientos:
        saldo = round(saldo + a["debe"] - a["haber"], 2)
        a["saldo_pdf"], a["saldo"] = a["saldo"], saldo
    info["saldo_final"] = saldo
    info["filas_saldo_distinto"] = sum(1 for a in asientos if a["saldo_pdf"] is not None and abs(a["saldo_pdf"] - a["saldo"]) > 0.005)
    return info, asientos
