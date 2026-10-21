# Liquidación FIAT - PDA (`liquidacion.gs`)

Script del archivo **"Incentivos pedidos FIAT"** (separado del archivo de pedidos). Calcula, para cualquier mes,
las líneas de la liquidación de Fiat y las compara con lo liquidado:

| Línea | Cálculo |
|---|---|
| Adelanto Margen | Pedidos netos del mes **sin REU en el mismo mes**: precio público s/IVA del modelo pedido (lista del mes del pedido) × 14% × 60% |
| Recupero Adelanto Margen | Unidades con REU en el mes cuyo pedido es de un mes anterior: −14% × 60% a precio del mes del REU |
| Márgen Por Pedido (60%) | Unidades con REU en el mes: 14% × 60% a precio del mes del REU |
| Margen por Entrega (40%) | Unidades con REU en el mes: 14% × 40% a precio del mes de entrega |
| Bonus Pedidos Aceptados | 2A (% categoría × público c/IVA del pedido) + 2B (% categoría × CE s/IVA NC1 × pedidos adicionales) |
| 2C por familia | Inc pedido del modelo pedido × CE s/IVA del modelo base de la familia (sin Cronos 90/10 B90 desde el grupo 18002) |

No se calculan: fondo de contingencia, incentivos por cuota (mora) y retribución a la suscripción.

## Fuentes (solo lectura)
- **Pedidos - PDA**: BASE, PRECIOS, OBJETIVOS e INCENTIVOS.
- **Entregas PDA**: `Base` (FBS) y `SGA`. El mes del REU es el mes de la *Fecha Ingreso* de SGA. Si una unidad
  entregada (FBS) todavía no está en SGA, queda como **PENDIENTE REU**.

## Hojas
- **RESUMEN**: elegir el mes en C3. Muestra cada línea (calculado, liquidado Fiat, diferencia y unidades), el 2C
  por familia y las entregas pendientes de REU.
- **PARAMETROS**: una fila por mes (cumplió suscripciones y patentamientos, margen, % adelanto, grupo y concepto B90).
  Los meses nuevos se agregan solos copiando el mes anterior; hay que revisarlos con las señales del mes.
- **FAMILIAS**: modelo base del 2C e IVA (10,5% para Fiorino, Strada, Toro y Titano).
- **MODELOS**: equivalencia entre la versión de FBS y el código. Las versiones nuevas se agregan solas con el
  código vacío y hay que completarlo.
- **LIQ_FIAT**: lo que liquidó Fiat (se carga a mano: mes, concepto, importe; para el 2C, la familia en la columna E).
- **PEDIDOS_LIQ / ENTREGAS_LIQ / MESES**: detalle armado por el script (no editar; se rehacen al actualizar).

Precios: si falta el precio de un modelo en un mes, las entregas usan el del mes más cercano y lo avisan en la Nota.
Para los pedidos (adelanto, 2A y 2C), el precio tiene que estar cargado en el mes.

## Instalación
1. En el archivo "Incentivos pedidos FIAT": Extensiones → Apps Script, pegar `liquidacion.gs` y guardar.
2. Recargar la planilla y usar el menú **Liquidación → Armar / reparar estructura** (la primera vez pide permisos).
3. Después, para recalcular: **Liquidación → Actualizar liquidación**.

## Prueba con septiembre 26 (datos del 21/10)
| Línea | Calculado | Fiat | Diferencia |
|---|---:|---:|---:|
| Adelanto Margen | 188.478.807,97 | 188.478.807,94 | 0,03 |
| Recupero | −152.513.478 | −159.658.407 | 7.144.929 |
| Margen 60% | 162.704.465 | 163.410.372 | −705.908 |
| Margen 40% | 107.278.912 | 108.038.802 | −759.890 |
| Bonus pedidos aceptados | 28.078.984,30 | 28.078.984,30 | 0 |
| 2C | 68.860.122,68 | 68.860.122,49 | 0,19 |

Faltaban los precios de septiembre de TT1 y FA4, así que se usaron los de octubre. Fiat recupera más de lo que da
la regla: los 2 Titano PT1 con pedido y REU en septiembre (Portalupi y Portillo) no tienen recupero en el cálculo.
