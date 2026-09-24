"""Normaliza y clasifica un extracto Banco Macro y genera una planilla de trabajo.

Formatos soportados:
- .csv  "Resumen" (Fecha;Referencia;Codigo Causal;Concepto;Debito;Credito;Saldo)
- .xlsx "Últimos Movimientos" de Macro Empresas (Fecha, Nro. de Referencia, Causal, Concepto, Importe con signo, Saldo)

Uso: python procesar_extracto.py extracto.(csv|xlsx) salida.xlsx
"""
import csv
import re
import sys
from collections import defaultdict
from datetime import datetime

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

from reglas_macro import BANCO, EMPRESAS_GRUPO, REGLAS

CUIT_RE = re.compile(r"\b(20|23|24|27|30|33|34)(\d{8})(\d)\b")
NUM_FMT = '#,##0.00;[Red]-#,##0.00'
HEADER_FILL = PatternFill("solid", fgColor="1F4E78")


PREGUNTAS = [
    ("Empresas del grupo", "Pasar la lista de CUIT de todas las empresas del grupo: los movimientos con ellas se marcan como internos."),
    ("Transferencias propias (3912 E-set M/T, 2018, 2026 BTOB, 3862 'CCDO DIST T')", "¿A qué cuentas van? En agosto hay $1.811M en E-set M/T y varias figuran en 'Cheques no registrados'. ¿Se registran en FBS contra la otra cuenta bancaria?"),
    ("Pagos a proveedores (2027 lotes, 3913 E-set D/T)", "¿FBS tiene el detalle por proveedor de cada lote Datanet / E-set? Sin eso no se pueden cruzar uno a uno."),
    ("IMP. AFIP $1.770 (3696, 4196)", "Débitos de $1.770 c/u (vistos en CHEXA). ¿Qué son? ¿Se registran uno por uno o agrupados?"),
    ("Número de Operación (4333, 4334)", "Créditos 'NNN - Numero de Operacion'. ¿Son cobranzas por convenio de recaudación? ¿Cómo se identifica al cliente?"),
    ("DL1 (366)", "Créditos '<CUIT> <NOMBRE> DL1'. ¿Qué operatoria es?"),
    ("PP- N/C Deuda Pública (1952)", "¿Qué corresponde a 'PP- N/C DEUD. PUBLICA-P.PREVIO'? En Turin agosto suman $95,5M."),
    ("Crédito 1232 (referencia de 30 dígitos)", "Crédito de $124.049.934,39 del 25/08 sin descripción: figura en 'Depósitos no registrados'. ¿Qué es?"),
    ("Banco del Sol (4061)", "'Transf. Presta BANCO DEL SOL': ¿préstamos, cobranzas financiadas o transferencias propias?"),
    ("Cobranzas clientes", "¿El recibo de FBS (RC-X-...) guarda el CUIT/DNI del cliente? Permitiría el cruce automático por CUIT + importe + fecha."),
    ("Columnas extra en la conciliación", "En cada sector hay una segunda columna de importes que no suma al total (ej. 'Canje de Valor', cheques 'RM-...'). ¿Qué significa?"),
    ("Partidas antiguas", "'Depósitos no registrados' arrastra 243 partidas anteriores a agosto 2026 (desde 2023). ¿Se van a regularizar o quedan como históricas?"),
    ("Datos FBS", "Exportación del mayor de las cuentas E y O de Banco Macro del mes (Excel/CSV) con fecha, comprobante, detalle, debe y haber."),
]


def num(s):
    return float(s.replace(".", "").replace(",", ".")) if s.strip() else 0.0


def cuit_valido(c):
    pesos = [5, 4, 3, 2, 7, 6, 5, 4, 3, 2]
    r = 11 - sum(int(d) * p for d, p in zip(c[:10], pesos)) % 11
    return int(c[10]) == {11: 0, 10: 9}.get(r, r)


def extraer_cuit(concepto):
    for m in CUIT_RE.finditer(concepto):
        c = "".join(m.groups())
        if cuit_valido(c):
            return c
    return ""


def extraer_nombre(concepto):
    for pat in (r"ING TRANSF:(.+?)-\d{11}", r"^TRANSF (.+?) \d{11}", r"TEF DATANET \w+ (.+?)\s+\d{11}",
                r"^CCERR (.+?) \d{11}", r"^TRMIN (.+?) \d{11}", r"^TRPRO (.+?) \d{11}",
                r"^\d{11} (.+?) DL1", r"Transf\. \w+ (.+?)\s+\d{11}"):
        m = re.search(pat, concepto)
        if m:
            return re.sub(r"\s+", " ", m.group(1)).strip(" /,")
    return ""


def clasificar(causal, concepto, importe, cuit):
    if cuit in EMPRESAS_GRUPO:
        return ("TRANSFERENCIA ENTRE CUENTAS PROPIAS / GRUPO", f"Movimiento con {EMPRESAS_GRUPO[cuit]}",
                "Transferencia interna")
    sentido = "D" if importe < 0 else "C"
    for cod, sen, patron, cat, sub, fbs in REGLAS:
        if cod == causal and sen in (None, sentido) and (patron is None or re.search(patron, concepto)):
            return cat, sub, fbs
    return "A IDENTIFICAR", "Código causal nuevo", ""


def _mov(fecha, ref, causal, concepto, debito, credito, saldo):
    concepto = concepto.strip()
    importe = credito - debito
    cuit = extraer_cuit(concepto)
    cat, sub, fbs = clasificar(causal, concepto, importe, cuit)
    return {"fecha": fecha, "ref": ref, "causal": causal, "concepto": concepto, "debito": debito, "credito": credito,
            "importe": importe, "saldo": saldo, "cat": cat, "sub": sub, "fbs": fbs,
            "cuit": cuit, "nombre": extraer_nombre(concepto)}


def leer_csv(path):
    with open(path, encoding="utf-8-sig", newline="") as f:
        rows = list(csv.DictReader(f, delimiter=";"))
    return {}, [_mov(datetime.strptime(r["Fecha"], "%d/%m/%Y").date(), r["Referencia"], r["Codigo Causal"],
                     r["Concepto"], num(r["Debito"]), num(r["Credito"]), num(r["Saldo"]))
                for r in rows]


def leer_xlsx(path):
    ws = load_workbook(path, data_only=True).active
    info, movs, cols = {}, [], None
    for row in ws.iter_rows(values_only=True):
        if row[0] == "Fecha":
            cols = {v: i for i, v in enumerate(row) if v}
        elif cols and isinstance(row[0], datetime):
            imp = float(row[cols["Importe"]])
            movs.append(_mov(row[0].date(), str(row[cols["Nro. de Referencia"]]), str(row[cols["Causal"]]),
                             row[cols["Concepto"]], max(-imp, 0.0), max(imp, 0.0), float(row[cols["Saldo"]])))
        elif isinstance(row[0], str) and row[0].startswith("Empresa:"):
            info["Empresa"] = row[0].split(":", 1)[1].strip()
        elif row[0] == "Número":
            info["Cuenta"] = row[2]
    return info, movs


def leer(path):
    info, movs = (leer_xlsx if path.lower().endswith(".xlsx") else leer_csv)(path)
    movs.reverse()  # el banco lo entrega del más nuevo al más viejo
    return info, movs


def gastos_bancarios(movs):
    """Resumen para el asiento mensual de gastos e impuestos bancarios."""
    g = defaultdict(float)
    for m in movs:
        t, d = m["concepto"].upper(), m["debito"] - m["credito"]
        if m["causal"] in ("1684", "1685"):
            g["imp"] += d
        elif m["causal"] == "1297":
            g["Percepciones IIBB (SIRCREB)"] += d
        elif m["causal"] == "4145":
            g["Retenciones IIBB rentas financieras"] += d
        elif m["cat"] == "GASTOS BANCARIOS":
            g["Percepciones IVA" if "PERCEP" in t or "IVA_PER" in t else "IVA 21%" if "IVA" in t else "Comisiones"] += d
    filas = [[k, g[k]] for k in ("Comisiones", "IVA 21%", "Percepciones IVA", "Percepciones IIBB (SIRCREB)",
                                 "Retenciones IIBB rentas financieras")]
    return filas + [["Imp. créditos y débitos (total)", g["imp"]], ["  Crédito computable 33%", g["imp"] * 0.33],
                    ["  Imp. créditos y débitos (gasto 67%)", g["imp"] * 0.67]]


def hoja(wb, titulo, headers, filas, anchos, formatos=None):
    ws = wb.create_sheet(titulo)
    ws.append(headers)
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = HEADER_FILL
        c.alignment = Alignment(wrap_text=True, vertical="center")
    for f in filas:
        ws.append(f)
    for i, w in enumerate(anchos, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    for col, fmt in (formatos or {}).items():
        for cell in ws[col][1:]:
            cell.number_format = fmt
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    return ws


def generar(info, movs, salida):
    wb = Workbook()
    wb.remove(wb.active)

    # Resumen y control de saldos
    ini = movs[0]["saldo"] - movs[0]["importe"]
    tot_c = sum(m["credito"] for m in movs)
    tot_d = sum(m["debito"] for m in movs)
    fin = movs[-1]["saldo"]
    quiebres = sum(1 for a, b in zip(movs, movs[1:]) if abs(a["saldo"] + b["importe"] - b["saldo"]) > 0.005)
    por_cat = defaultdict(lambda: [0, 0.0, 0.0])
    for m in movs:
        x = por_cat[m["cat"]]
        x[0] += 1; x[1] += m["credito"]; x[2] += m["debito"]
    ws = hoja(wb, "Resumen", ["Concepto", "Valor"], [
        ["Banco", BANCO],
        *[[k, v] for k, v in info.items()],
        ["Período", f"{movs[0]['fecha']:%d/%m/%Y} al {movs[-1]['fecha']:%d/%m/%Y}"],
        ["Cantidad de movimientos", len(movs)],
        ["Saldo inicial (calculado)", ini],
        ["(+) Créditos", tot_c],
        ["(-) Débitos", tot_d],
        ["Saldo final según extracto", fin],
        ["Diferencia control (debe ser 0)", round(ini + tot_c - tot_d - fin, 2)],
        ["Quiebres de saldo entre filas", quiebres],
        [], ["Categoría", "Cant.", "Créditos", "Débitos", "Neto"],
    ] + [[k, v[0], v[1], v[2], v[1] - v[2]] for k, v in sorted(por_cat.items(), key=lambda kv: -(kv[1][1] + kv[1][2]))]
      + [[], ["Gastos e impuestos bancarios (para asiento)", "Importe"]] + gastos_bancarios(movs),
        [42, 20, 20, 20, 20])
    for row in ws.iter_rows(min_row=2):
        if row[0].value in ("Categoría", "Gastos e impuestos bancarios (para asiento)"):
            for c in row:
                c.font = Font(bold=True)
        for c in row[1:]:
            if isinstance(c.value, float):
                c.number_format = NUM_FMT

    # Catálogo de conceptos (una fila por regla, con estadísticas del mes)
    stats = defaultdict(lambda: [0, 0.0, 0.0, []])
    for m in movs:
        s = stats[(m["causal"], m["sub"])]
        s[0] += 1; s[1] += m["debito"]; s[2] += m["credito"]
        if len(s[3]) < 3 and m["concepto"] not in s[3]:
            s[3].append(m["concepto"])
    filas, vistos = [], set()
    for cod, sen, patron, cat, sub, fbs in REGLAS:
        s = stats.get((cod, sub))
        if not s or (cod, sub) in vistos:
            continue
        vistos.add((cod, sub))
        filas.append([BANCO, int(cod), {"D": "Débito", "C": "Crédito"}.get(sen, "Ambos"), patron or "(todos)", cat, sub,
                      s[0], s[1], s[2], " | ".join(s[3]), fbs, "A validar", ""])
    for (cod, sub), s in stats.items():
        if (cod, sub) not in vistos:
            cat = next(m["cat"] for m in movs if (m["causal"], m["sub"]) == (cod, sub))
            filas.append([BANCO, int(cod), "Ambos", "(todos)", cat, sub, s[0], s[1], s[2], " | ".join(s[3]),
                          "", "A validar", ""])
    hoja(wb, "Catalogo_Conceptos", ["Banco", "Cód. causal", "Sentido", "Patrón en concepto", "Categoría",
                                    "Descripción propuesta", "Cant. mes", "Débitos mes", "Créditos mes",
                                    "Ejemplos del extracto", "Registro esperado en FBS", "Estado validación",
                                    "Comentario administración"],
         filas, [8, 10, 9, 16, 34, 42, 9, 18, 18, 70, 34, 14, 40], {"H": NUM_FMT, "I": NUM_FMT})

    # Movimientos normalizados
    hoja(wb, "Movimientos", ["Fecha", "Referencia", "Cód. causal", "Concepto", "Débito", "Crédito", "Importe (+/-)",
                             "Saldo", "Categoría", "Descripción", "CUIT contraparte", "Nombre contraparte",
                             "Estado conciliación", "Comprobante FBS", "Observaciones"],
         [[m["fecha"], m["ref"], int(m["causal"]), m["concepto"], m["debito"] or None, m["credito"] or None,
           m["importe"], m["saldo"], m["cat"], m["sub"], m["cuit"], m["nombre"], "Pendiente", "", ""] for m in movs],
         [11, 13, 9, 48, 16, 16, 16, 18, 30, 34, 14, 28, 16, 18, 30],
         {"A": "dd/mm/yyyy", "E": NUM_FMT, "F": NUM_FMT, "G": NUM_FMT, "H": NUM_FMT})

    hoja(wb, "Preguntas_Admin", ["#", "Tema", "Pregunta", "Respuesta administración"],
         [[i, t, q, ""] for i, (t, q) in enumerate(PREGUNTAS, 1)], [4, 34, 100, 50])

    wb.save(salida)


if __name__ == "__main__":
    generar(*leer(sys.argv[1]), sys.argv[2])
