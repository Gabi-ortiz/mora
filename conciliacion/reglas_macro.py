"""Catálogo de conceptos del extracto Banco Macro (formato CSV de Resumen).

Cada regla: (codigo_causal, sentido, patron_concepto_regex, categoria, subcategoria, registro_fbs_esperado).
- sentido: "D" débito, "C" crédito o None (ambos). Un mismo código puede significar cosas distintas según el signo.
- patron: regex sobre el concepto, o None para todo el código causal.
Se evalúan en orden; gana la primera que coincide. Antes de las reglas, si la contraparte es una empresa
del grupo (EMPRESAS_GRUPO) el movimiento se clasifica como transferencia interna.
Todas las reglas nacen con estado "A validar" hasta que administración las confirme.
"""

BANCO = "MACRO"

# CUIT -> razón social de las empresas del grupo (completar con el resto)
EMPRESAS_GRUPO = {
    "30669400167": "TURIN SA",
    "30681026386": "CHEXA SA",
}

# Categorías comunes a todos los bancos (base para expandir al resto)
CATEGORIAS = [
    "COBRANZA CLIENTES",
    "COBRANZA EMPRESAS / ASEGURADORAS",
    "PAGO PROVEEDORES",
    "SUELDOS",
    "TRANSFERENCIA ENTRE CUENTAS PROPIAS / GRUPO",
    "INVERSIONES (FCI)",
    "PLAN DE AHORRO / TERMINAL",
    "IMPUESTOS Y RETENCIONES",
    "GASTOS BANCARIOS",
    "DEBITOS AUTOMATICOS",
    "DEPOSITOS EN EFECTIVO",
    "COBRANZA TARJETAS",
    "CHEQUES",
    "FINANCIACION / PRESTAMOS",
    "A IDENTIFICAR",
]

REGLAS = [
    # --- Transferencias recibidas de terceros ---
    ("493", None, None, "COBRANZA CLIENTES", "Transferencia recibida (otros bancos)", "Recibo / cobranza cliente"),
    ("893", None, None, "COBRANZA CLIENTES", "Transferencia recibida (otros bancos)", "Recibo / cobranza cliente"),
    ("4397", None, None, "COBRANZA CLIENTES", "Transferencia recibida (solo CUIT)", "Recibo / cobranza cliente"),
    ("3816", None, None, "COBRANZA CLIENTES", "Transferencia recibida (otros bancos)", "Recibo / cobranza cliente"),
    ("4543", None, None, "COBRANZA CLIENTES", "Transferencia recibida identificada (nombre + CUIT)", "Recibo / cobranza cliente"),
    ("4544", None, None, "COBRANZA CLIENTES", "Transferencia recibida identificada (nombre + CUIT)", "Recibo / cobranza cliente"),
    ("4080", None, None, "COBRANZA CLIENTES", "Transferencia inmediata / billetera virtual", "Recibo / cobranza cliente"),
    ("4081", None, None, "COBRANZA CLIENTES", "Transferencia inmediata / billetera virtual", "Recibo / cobranza cliente"),
    ("4091", None, None, "COBRANZA CLIENTES", "Transferencia inmediata / billetera virtual", "Recibo / cobranza cliente"),
    ("4013", None, None, "COBRANZA CLIENTES", "CREDIN (crédito inmediato)", "Recibo / cobranza cliente"),
    ("4016", None, None, "COBRANZA CLIENTES", "CREDIN (crédito inmediato)", "Recibo / cobranza cliente"),
    ("2010", None, None, "COBRANZA CLIENTES", "Transferencia minorista (TRMIN)", "Recibo / cobranza cliente"),
    ("366", None, None, "A IDENTIFICAR", "Crédito 'DL1' con CUIT persona", "¿Cobranza cliente? a validar"),
    ("4333", None, None, "A IDENTIFICAR", "Crédito 'Número de Operación'", "¿Cobranza por convenio de recaudación?"),
    ("4334", None, None, "A IDENTIFICAR", "Crédito 'Número de Operación'", "¿Cobranza por convenio de recaudación?"),
    ("4218", None, None, "A IDENTIFICAR", "Transferencia 'APC Aportes de capital'", "a validar"),
    ("4220", None, None, "A IDENTIFICAR", "Transferencia 'APC Aportes de capital'", "a validar"),
    # --- Datanet (home banking empresas) ---
    ("2027", None, r"PROVEEDORES", "PAGO PROVEEDORES", "Transferencia Datanet a proveedores", "Orden de pago proveedor"),
    ("2027", None, r"TEF DATANET PR", "COBRANZA EMPRESAS / ASEGURADORAS", "TEF Datanet recibida de empresa", "Recibo / cobranza"),
    ("4015", None, r"EGRESO", "PAGO PROVEEDORES", "Egreso transferencia inmediata", "Orden de pago proveedor"),
    ("3862", "C", None, "COBRANZA EMPRESAS / ASEGURADORAS", "Circuito cerrado (CCERR)", "Recibo / cobranza"),
    ("3956", None, None, "COBRANZA EMPRESAS / ASEGURADORAS", "Circuito cerrado (CCERR) San Cristóbal", "Recibo / cobranza"),
    ("2024", None, None, "IMPUESTOS Y RETENCIONES", "Pago AFIP/ARCA vía Datanet", "Pago impuesto"),
    # --- Entre cuentas propias / grupo (CUIT 30681026386 = CHEXA SA) ---
    ("2018", None, None, "TRANSFERENCIA ENTRE CUENTAS PROPIAS / GRUPO", "TEF Datanet misma titularidad (M/T)", "Transferencia interna"),
    ("2026", None, None, "TRANSFERENCIA ENTRE CUENTAS PROPIAS / GRUPO", "TEF Datanet BTOB", "Transferencia interna (a validar)"),
    ("3808", None, None, "TRANSFERENCIA ENTRE CUENTAS PROPIAS / GRUPO", "Transferencia desde CHEXA SA", "Transferencia interna"),
    ("4254", None, None, "TRANSFERENCIA ENTRE CUENTAS PROPIAS / GRUPO", "Transferencia inmediata CHEXA SA", "Transferencia interna"),
    ("4267", None, None, "TRANSFERENCIA ENTRE CUENTAS PROPIAS / GRUPO", "Transferencia inmediata a CHEXA SA", "Transferencia interna"),
    # --- Inversiones ---
    ("1411", None, None, "INVERSIONES (FCI)", "Suscripción FCI", "Movimiento inversión"),
    ("4361", None, None, "INVERSIONES (FCI)", "Suscripción FCI", "Movimiento inversión"),
    ("1413", None, None, "INVERSIONES (FCI)", "Rescate FCI", "Movimiento inversión"),
    # --- Terminal / plan de ahorro ---
    ("4196", None, r"PLAN AHORRO", "PLAN DE AHORRO / TERMINAL", "Débito Plan de Ahorro Chevrolet", "Pago a terminal"),
    ("1445", None, None, "PLAN DE AHORRO / TERMINAL", "Transferencia recibida de Chevrolet (TRPRO)", "Cobro a terminal"),
    ("1952", None, None, "A IDENTIFICAR", "PP- N/C Deuda Pública - P.Previo", "a validar"),
    # --- Impuestos ---
    ("1684", None, None, "IMPUESTOS Y RETENCIONES", "Impuesto Ley 25.413 s/débitos", "Asiento impuesto (automático)"),
    ("1685", None, None, "IMPUESTOS Y RETENCIONES", "Impuesto Ley 25.413 s/créditos", "Asiento impuesto (automático)"),
    ("1297", None, None, "IMPUESTOS Y RETENCIONES", "Retención IIBB SIRCREB", "Asiento retención (automático)"),
    ("4145", None, None, "IMPUESTOS Y RETENCIONES", "Retención IIBB Córdoba renta financiera", "Asiento retención (automático)"),
    ("4196", None, r"IMP\. AFIP", "IMPUESTOS Y RETENCIONES", "IMP. AFIP ($1.770 c/u)", "a validar"),
    ("3696", None, None, "IMPUESTOS Y RETENCIONES", "IMP. AFIP ($1.770 c/u)", "a validar"),
    ("23", None, r"ARCA|AFIP", "IMPUESTOS Y RETENCIONES", "Débito directo ARCA", "Pago impuesto"),
    # --- Gastos bancarios ---
    ("531", None, None, "GASTOS BANCARIOS", "Comisión transferencia Datanet + IVA", "Asiento gasto bancario (automático)"),
    ("1210", None, None, "GASTOS BANCARIOS", "Comisión convenio recaudación 61068", "Asiento gasto bancario (automático)"),
    ("1211", None, None, "GASTOS BANCARIOS", "IVA / percepción convenio recaudación 61068", "Asiento gasto bancario (automático)"),
    ("17", None, None, "GASTOS BANCARIOS", "Comisión resumen cuenta + IVA", "Asiento gasto bancario (automático)"),
    ("33", None, None, "GASTOS BANCARIOS", "Comisión mantenimiento cuenta + IVA", "Asiento gasto bancario (automático)"),
    ("3872", None, None, "GASTOS BANCARIOS", "Comisión transferencia", "Asiento gasto bancario (automático)"),
    ("1450", None, None, "GASTOS BANCARIOS", "Comisión transferencia", "Asiento gasto bancario (automático)"),
    # --- Sueldos / débitos automáticos / efectivo / financiación ---
    ("1693", None, None, "SUELDOS", "Pago de remuneraciones", "Liquidación de sueldos"),
    ("23", None, None, "DEBITOS AUTOMATICOS", "Débito automático (seguros, peajes, cámara)", "Pago servicio / seguro"),
    ("1735", None, None, "DEPOSITOS EN EFECTIVO", "Interdepósito en efectivo", "Depósito de caja"),
    ("4061", None, None, "FINANCIACION / PRESTAMOS", "Transf. 'Presta' Banco del Sol", "a validar"),
    ("4066", None, None, "A IDENTIFICAR", "Transf. Varios VAZCAR SAS", "a validar"),
    # --- Agregados con el extracto de Turin (agosto 2026) ---
    ("3862", "D", None, "TRANSFERENCIA ENTRE CUENTAS PROPIAS / GRUPO", "Transf. MacrOnline 'CCDO DIST T'", "Transferencia interna"),
    ("3912", None, None, "TRANSFERENCIA ENTRE CUENTAS PROPIAS / GRUPO", "Transf. MacrOnline E-set M/T (misma titularidad)", "Transferencia interna"),
    ("3913", None, None, "PAGO PROVEEDORES", "Transf. MacrOnline E-set D/T (distinta titularidad)", "Orden de pago"),
    ("4085", "D", None, "PAGO PROVEEDORES", "Transferencia inmediata enviada", "Orden de pago / devolución"),
    ("4250", None, None, "TRANSFERENCIA ENTRE CUENTAS PROPIAS / GRUPO", "CREDIN propio", "Transferencia interna"),
    ("4253", None, None, "TRANSFERENCIA ENTRE CUENTAS PROPIAS / GRUPO", "Transferencia inmediata propia", "Transferencia interna"),
    ("183",  "C", None, "COBRANZA TARJETAS", "Liquidación Payway (tarjetas)", "Liquidación de tarjeta"),
    ("183",  "D", None, "COBRANZA TARJETAS", "Débito liquidación Payway (contracargo/ajuste)", "Liquidación de tarjeta"),
    ("14",   None, None, "CHEQUES", "Acreditación cheques depositados (remesas)", "Depósito de cheques de terceros"),
    ("85",   None, None, "CHEQUES", "Cheque propio pagado por cámara", "Cheque emitido"),
    ("2837", None, None, "CHEQUES", "Cheque propio pagado por canje interno", "Cheque emitido"),
    ("1232", None, None, "A IDENTIFICAR", "Crédito con referencia numérica (30 dígitos)", "a validar"),
    ("16",   None, r"IVA", "GASTOS BANCARIOS", "IVA comisión valores al cobro", "Asiento gasto bancario (automático)"),
    ("16",   None, None, "GASTOS BANCARIOS", "Comisión administración valores al cobro", "Asiento gasto bancario (automático)"),
    ("1724", None, None, "GASTOS BANCARIOS", "Comisión depósito cheque otra sucursal + IVA", "Asiento gasto bancario (automático)"),
    ("574",  None, None, "GASTOS BANCARIOS", "Comisión cheque pagado por clearing + IVA", "Asiento gasto bancario (automático)"),
    ("3914", None, None, "GASTOS BANCARIOS", "Comisión transferencia MacrOnline", "Asiento gasto bancario (automático)"),
    ("3734", None, None, "GASTOS BANCARIOS", "Comisión débito pago remuneraciones", "Asiento gasto bancario (automático)"),
    ("3205", None, None, "GASTOS BANCARIOS", "IVA comisión", "Asiento gasto bancario (automático)"),
    ("1212", None, None, "GASTOS BANCARIOS", "Mantenimiento caja de seguridad", "Asiento gasto bancario (automático)"),
    ("1214", None, None, "GASTOS BANCARIOS", "IVA caja de seguridad", "Asiento gasto bancario (automático)"),
    ("1215", None, None, "GASTOS BANCARIOS", "Percepción IVA caja de seguridad", "Asiento gasto bancario (automático)"),
    # --- Agregados con julio 2026 (Turin) ---
    ("5",    None, None, "GASTOS BANCARIOS", "Intereses adelanto en cuenta corriente + IVA", "Asiento gasto bancario (automático)"),
    ("1802", None, None, "GASTOS BANCARIOS", "Comisión cheque consulta + IVA", "Asiento gasto bancario (automático)"),
    ("1479", None, None, "IMPUESTOS Y RETENCIONES", "Impuesto de sellos Córdoba (DGR)", "Asiento gasto bancario (automático)"),
    ("1972", None, None, "IMPUESTOS Y RETENCIONES", "Retención IIBB Tucumán", "Asiento gasto bancario (automático)"),
    ("4196", None, r"PLAN", "PLAN DE AHORRO / TERMINAL", "Débito plan de ahorro (Chevrolet, Fiat, etc.)", "Pago a terminal"),
]
