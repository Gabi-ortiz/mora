"""Catálogo de conceptos del extracto Banco Macro (formato CSV de Resumen).

Cada regla: (codigo_causal, patron_concepto_regex | None, categoria, subcategoria, registro_fbs_esperado).
Se evalúan en orden; gana la primera que coincide. Si el patrón es None, aplica a todo el código causal.
Todas las reglas nacen con estado "A validar" hasta que administración las confirme.
"""

BANCO = "MACRO"

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
    "FINANCIACION / PRESTAMOS",
    "A IDENTIFICAR",
]

REGLAS = [
    # --- Transferencias recibidas de terceros ---
    ("493",  None, "COBRANZA CLIENTES", "Transferencia recibida (otros bancos)", "Recibo / cobranza cliente"),
    ("893",  None, "COBRANZA CLIENTES", "Transferencia recibida (otros bancos)", "Recibo / cobranza cliente"),
    ("4397", None, "COBRANZA CLIENTES", "Transferencia recibida (solo CUIT)", "Recibo / cobranza cliente"),
    ("3816", None, "COBRANZA CLIENTES", "Transferencia recibida (otros bancos)", "Recibo / cobranza cliente"),
    ("4543", None, "COBRANZA CLIENTES", "Transferencia recibida identificada (nombre + CUIT)", "Recibo / cobranza cliente"),
    ("4544", None, "COBRANZA CLIENTES", "Transferencia recibida identificada (nombre + CUIT)", "Recibo / cobranza cliente"),
    ("4080", None, "COBRANZA CLIENTES", "Transferencia inmediata / billetera virtual", "Recibo / cobranza cliente"),
    ("4081", None, "COBRANZA CLIENTES", "Transferencia inmediata / billetera virtual", "Recibo / cobranza cliente"),
    ("4091", None, "COBRANZA CLIENTES", "Transferencia inmediata / billetera virtual", "Recibo / cobranza cliente"),
    ("4013", None, "COBRANZA CLIENTES", "CREDIN (crédito inmediato)", "Recibo / cobranza cliente"),
    ("4016", None, "COBRANZA CLIENTES", "CREDIN (crédito inmediato)", "Recibo / cobranza cliente"),
    ("2010", None, "COBRANZA CLIENTES", "Transferencia minorista (TRMIN)", "Recibo / cobranza cliente"),
    ("366",  None, "A IDENTIFICAR", "Crédito 'DL1' con CUIT persona", "¿Cobranza cliente? a validar"),
    ("4333", None, "A IDENTIFICAR", "Crédito 'Número de Operación'", "¿Cobranza por convenio de recaudación?"),
    ("4334", None, "A IDENTIFICAR", "Crédito 'Número de Operación'", "¿Cobranza por convenio de recaudación?"),
    ("4218", None, "A IDENTIFICAR", "Transferencia 'APC Aportes de capital'", "a validar"),
    ("4220", None, "A IDENTIFICAR", "Transferencia 'APC Aportes de capital'", "a validar"),
    # --- Datanet (home banking empresas) ---
    ("2027", r"PROVEEDORES", "PAGO PROVEEDORES", "Transferencia Datanet a proveedores", "Orden de pago proveedor"),
    ("2027", r"TEF DATANET PR", "COBRANZA EMPRESAS / ASEGURADORAS", "TEF Datanet recibida de empresa", "Recibo / cobranza"),
    ("4015", r"EGRESO", "PAGO PROVEEDORES", "Egreso transferencia inmediata", "Orden de pago proveedor"),
    ("3862", None, "COBRANZA EMPRESAS / ASEGURADORAS", "Circuito cerrado (CCERR)", "Recibo / cobranza"),
    ("3956", None, "COBRANZA EMPRESAS / ASEGURADORAS", "Circuito cerrado (CCERR) San Cristóbal", "Recibo / cobranza"),
    ("2024", None, "IMPUESTOS Y RETENCIONES", "Pago AFIP/ARCA vía Datanet", "Pago impuesto"),
    # --- Entre cuentas propias / grupo (CUIT 30681026386 = CHEXA SA) ---
    ("2018", None, "TRANSFERENCIA ENTRE CUENTAS PROPIAS / GRUPO", "TEF Datanet misma titularidad (M/T)", "Transferencia interna"),
    ("2026", None, "TRANSFERENCIA ENTRE CUENTAS PROPIAS / GRUPO", "TEF Datanet BTOB", "Transferencia interna (a validar)"),
    ("3808", None, "TRANSFERENCIA ENTRE CUENTAS PROPIAS / GRUPO", "Transferencia desde CHEXA SA", "Transferencia interna"),
    ("4254", None, "TRANSFERENCIA ENTRE CUENTAS PROPIAS / GRUPO", "Transferencia inmediata CHEXA SA", "Transferencia interna"),
    ("4267", None, "TRANSFERENCIA ENTRE CUENTAS PROPIAS / GRUPO", "Transferencia inmediata a CHEXA SA", "Transferencia interna"),
    # --- Inversiones ---
    ("1411", None, "INVERSIONES (FCI)", "Suscripción FCI", "Movimiento inversión"),
    ("4361", None, "INVERSIONES (FCI)", "Suscripción FCI", "Movimiento inversión"),
    ("1413", None, "INVERSIONES (FCI)", "Rescate FCI", "Movimiento inversión"),
    # --- Terminal / plan de ahorro ---
    ("4196", r"PLAN AHORRO", "PLAN DE AHORRO / TERMINAL", "Débito Plan de Ahorro Chevrolet", "Pago a terminal"),
    ("1445", None, "PLAN DE AHORRO / TERMINAL", "Transferencia recibida de Chevrolet (TRPRO)", "Cobro a terminal"),
    ("1952", None, "A IDENTIFICAR", "PP- N/C Deuda Pública - P.Previo", "a validar"),
    # --- Impuestos ---
    ("1684", None, "IMPUESTOS Y RETENCIONES", "Impuesto Ley 25.413 s/débitos", "Asiento impuesto (automático)"),
    ("1685", None, "IMPUESTOS Y RETENCIONES", "Impuesto Ley 25.413 s/créditos", "Asiento impuesto (automático)"),
    ("1297", None, "IMPUESTOS Y RETENCIONES", "Retención IIBB SIRCREB", "Asiento retención (automático)"),
    ("4145", None, "IMPUESTOS Y RETENCIONES", "Retención IIBB Córdoba renta financiera", "Asiento retención (automático)"),
    ("4196", r"IMP\. AFIP", "IMPUESTOS Y RETENCIONES", "IMP. AFIP ($1.770 c/u)", "a validar"),
    ("3696", None, "IMPUESTOS Y RETENCIONES", "IMP. AFIP ($1.770 c/u)", "a validar"),
    ("23",   r"ARCA|AFIP", "IMPUESTOS Y RETENCIONES", "Débito directo ARCA", "Pago impuesto"),
    # --- Gastos bancarios ---
    ("531",  None, "GASTOS BANCARIOS", "Comisión transferencia Datanet + IVA", "Asiento gasto bancario (automático)"),
    ("1210", None, "GASTOS BANCARIOS", "Comisión convenio recaudación 61068", "Asiento gasto bancario (automático)"),
    ("1211", None, "GASTOS BANCARIOS", "IVA / percepción convenio recaudación 61068", "Asiento gasto bancario (automático)"),
    ("17",   None, "GASTOS BANCARIOS", "Comisión resumen cuenta + IVA", "Asiento gasto bancario (automático)"),
    ("33",   None, "GASTOS BANCARIOS", "Comisión mantenimiento cuenta + IVA", "Asiento gasto bancario (automático)"),
    ("3872", None, "GASTOS BANCARIOS", "Comisión transferencia", "Asiento gasto bancario (automático)"),
    ("1450", None, "GASTOS BANCARIOS", "Comisión transferencia", "Asiento gasto bancario (automático)"),
    # --- Sueldos / débitos automáticos / efectivo / financiación ---
    ("1693", None, "SUELDOS", "Pago de remuneraciones", "Liquidación de sueldos"),
    ("23",   None, "DEBITOS AUTOMATICOS", "Débito automático (seguros, peajes, cámara)", "Pago servicio / seguro"),
    ("1735", None, "DEPOSITOS EN EFECTIVO", "Interdepósito en efectivo", "Depósito de caja"),
    ("4061", None, "FINANCIACION / PRESTAMOS", "Transf. 'Presta' Banco del Sol", "a validar"),
    ("4066", None, "A IDENTIFICAR", "Transf. Varios VAZCAR SAS", "a validar"),
]
