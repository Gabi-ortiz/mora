# Propuesta: norma de registración de movimientos bancarios en FBS

Objetivo: que cada movimiento del extracto se pueda cruzar automáticamente con su registro en FBS.
El cruce busca en este orden: **1) CUIT/CUIL → 2) referencia del banco → 3) fecha → 4) importe**.
Cuanto más arriba se resuelva, más confiable es (el importe solo se repite mucho).

## Reglas generales

1. **Un registro por movimiento del banco.** Si un cliente paga con 2 transferencias, el recibo lleva 2 valores
   (uno por transferencia) con su importe exacto. Si una transferencia paga 2 comprobantes, se imputa en un recibo.
2. **Fecha = fecha de acreditación/débito en el extracto**, no la fecha de carga.
3. **Importe exacto**, sin redondeos (hoy hay diferencias de $0,01, ej. rescate FCI 74.228.577,96 vs 74.228.577,95).
4. **Comentario con formato fijo**, siempre en el mismo orden:
   `TIPO-COMPROBANTE / CUIT 20123456789 / REF 123456789 / NOMBRE`
   - CUIT de 11 dígitos, sin guiones ni espacios (hoy aparece como `TR30 50004045 5`, `SX 32458203`, `TR14401150`).
   - REF = columna *Referencia* del extracto (número de operación / cheque / solicitud).
5. **Cuentas E y O**: se registra en O y se pasa a E al confirmarlo contra el extracto (o al revés, pero siempre igual).
   El pase no debe cambiar fecha, importe ni comentario del registro original.

## Por tipo de movimiento

| Movimiento | Qué registrar | Cómo se cruza |
|---|---|---|
| Cobranza por transferencia / CREDIN / billetera | Recibo con CUIT del **ordenante** (quien transfiere) + REF | CUIT |
| Cobranza de empresas / aseguradoras (Datanet, CCERR) | Recibo con CUIT de la empresa | CUIT |
| Pago a proveedor (Datanet PROVEEDORES, E-set D/T) | Orden de pago con CUIT del proveedor + nº de operación o lote | CUIT / lote |
| Cheques propios (cámara, canje interno) | Nº de cheque (ya se hace: `Nº 64274941`) | Referencia ✔ |
| Cheques de terceros depositados (remesas) | Nº de cheque / boleta de depósito | Referencia |
| Transferencias entre cuentas propias (E-set M/T, Datanet MT, Pasefondo) | Comprobante "PASE" con banco origen/destino + nº de operación, **el mismo día** | Referencia / fecha |
| FCI (suscripción / rescate) | Nº de solicitud del extracto (`Liq.Susc 13112230`, `Sol.Resc 13132337`) | Referencia |
| Tarjetas (Payway, Naranja, etc.) | Liquidación con nº de establecimiento/liquidación del extracto (`PAGO43466226`) e importe neto acreditado | Referencia |
| Venta / descuento de cheques | Neto acreditado o bruto + gastos, con la referencia del crédito | Combinación |
| Sueldos | Un registro por débito de lote del banco, o lote identificado con la fecha del débito | Lote del día |
| Impuestos (AFIP/ARCA vía Datanet) | Un registro por VEP pagado, con nº de VEP | Referencia / lote |
| Gastos e impuestos bancarios | **Un asiento mensual** "gastos bancarios MM/AAAA" (ya se hace y cierra exacto) | Total mensual ✔ |

## Qué se detectó en agosto 2026 (Turin, Macro)

- **$1.146M de débitos del mes sin registrar**: $1.144M son transferencias a cuentas propias (MacrOnline E-set M/T)
  y una suscripción de FCI. Son las más fáciles de normalizar.
- **Cobranzas**: sin CUIT normalizado, 99 cruces salieron por coincidencia de apellido y 119 solo por importe (quedan como "Sugerido", a revisar).
- **Posible duplicado**: `RM-390702` y `RM-390995` (ambos "BANCO MACRO Nº 20082026", $39.353.190,80) contra un solo
  pago AFIP en el banco.
- **Partidas antiguas**: el sector *Depósitos no registrados* arrastra 219 partidas anteriores a agosto 2026
  (desde mayo 2024), por $220M.
