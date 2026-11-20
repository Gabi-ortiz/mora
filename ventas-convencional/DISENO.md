# Ventas convencional — estructura común para las empresas del grupo

Estado: borrador v0 (20/11/2026). Base de diseño: modelo operativo de Chexa +
reportes de Turin. Pendiente de validar con Gabi antes de construir.

## Objetivo
FBS → Operativos PV (Stock, Preventas, Créditos, Gestoría, Entregas) → Reportes
(por empresa y de todo el grupo). Una plantilla única que se copia por empresa.

## Circuito de cada venta
| # | Etapa | Dónde vive | Quién escribe |
|---|-------|-----------|---------------|
| 1 | Carga de la preventa en FBS | FBS | Vendedor / admin de ventas |
| 2 | Actualización de la base | `BASE_FBS` | Importación (no se edita a mano) |
| 3 | Carga de datos de control (lo que FBS no tiene) | `PREVENTAS` | Control |
| 4 | Créditos | `CREDITOS` | Sector créditos |
| 5 | Gestoría | `GESTORIA` | Sector gestoría |
| 6 | Entrega y postventa | `ENTREGAS` | Sector entregas |

Regla: cada dato se carga en un solo lugar; el resto lo ve, no lo escribe.

## Libro por empresa (se copia de la plantilla)
- `CONFIG`: nombre de empresa, sucursales, vendedores, objetivos del mes,
  horarios del reporte, nombres/teléfonos de los responsables de cada sector.
- `BASE_FBS`: export de FBS tal cual (columnas de la BASE de Chexa: Clave,
  Fecha_Preventa, Preventa, OrigenOperacion, Tipo_Operacion, Modelo, Cliente,
  Telefono/Celular, Email, PrecioVenta, OP_Credito, OP_Usado, Sucursal,
  Vendedor, VIN, Color, FacturaFecha, FechaPatentamiento, FechaSalida, etc.).
- `STOCK`: (Gabi ya lo tiene avanzado).
- `PREVENTAS`: una fila por preventa con la clave normalizada, datos de FBS
  (por lookup) + columnas de control manuales (patenta cliente/empresa/reventa,
  tiene usado, doc usado OK, retira en el día, cancelado + fecha, tipo de venta).
- `CREDITOS`: banco, seguro elegido, consulta de inhibido, prenda entregada /
  fecha, firma de legajo / fecha, enviado a liquidar / fecha, liquidado / fecha.
- `GESTORIA`: recepción de carpeta, toma de usado, trámite armado, ingreso a
  registro, fecha de patente, patente, comentarios.
- `ENTREGAS`: fecha de entrega, primer contacto, comentarios de postventa.
- `WHATSAPP`: ver más abajo.
- `REPORTE`: se calcula solo (ver más abajo).
- `LOG`: una línea por corrida de sincronización / reporte.

### Clave de preventa (crítico)
Los archivos usan formatos distintos: `PV 00737/3`, `PV 29545/9`, `35330/0`.
Se define `CLAVE = número sin ceros a la izquierda + "/" + dígito` (ej. `737/3`),
calculada una sola vez en `BASE_FBS` y usada en todos los cruces. Sin esto no
se puede consolidar el grupo.

### Cómo llegan las filas a cada hoja de sector
Las hojas de sector tienen columnas que la gente tipea, así que no pueden ser
una fórmula tipo FILTER. Un script (Apps Script) agrega al final de cada hoja
las claves nuevas de `BASE_FBS` que todavía no están (append-only por clave),
copia los datos de FBS que corresponden (cliente, vehículo, teléfono) y deja
vacías las columnas de hitos. Una preventa cancelada en FBS se marca, no se
borra.

## Módulo WhatsApp (común a todas las empresas)
- `MENSAJES` (en `CONFIG` o hoja propia): plantillas por evento y por empresa,
  con variables `{nombre}`, `{vehiculo}`, `{empresa}`, `{responsable}`.
  Eventos de arranque (los de Chexa): bienvenida/entregas (con variante si
  tiene usado), créditos (envío de seguros), y los que cada terminal sume.
- `WHATSAPP`: por cada preventa arma el link `api.whatsapp.com/send?phone=…&text=…`
  (con prefijo país), muestra "Enviar", y registra enviado + fecha/hora.
- Envío manual con un clic (abre WhatsApp). Si más adelante se quiere envío
  automático real, habría que pasar a WhatsApp Business API; queda fuera de v1.
- Cada empresa edita solo sus plantillas; la lógica es la misma en todas.
- Observación del archivo actual de Chexa: en `Erika` el link de la fila 3
  quedó con `text=FALSE` (la columna de mensaje está desalineada con un
  checkbox). La plantilla debe separar bien columnas de mensaje y de estado.

## Reportes
- Por empresa (`REPORTE`): ventas del mes vs objetivo por sucursal/vendedor y
  por marca/tipo (0km, patentados, planes, usados), avance %, delta diario,
  pendientes por sector (créditos sin liquidar, gestoría sin patente,
  entregas pendientes), cancelaciones.
- Grupo: un libro consolidado que lee de cada empresa una tabla plana
  (`REPORTE_DATA`: empresa, sucursal, vendedor, marca, tipo, mes, cantidad,
  objetivo) y arma los totales y rankings del grupo.
- Actualización: Apps Script con triggers por horario (a definir, por ejemplo
  4 veces por día) + un botón (menú / dibujo con script asignado) "Actualizar
  ahora". Cada corrida deja registro en `LOG`.

## Precios y condiciones
Quedan por empresa (cada una en su `CONFIG` o hoja propia). Fuera de la
plantilla común por ahora; se evalúa automatizar más adelante.

## Plan de construcción propuesto
1. Plantilla vacía con las hojas y columnas anteriores.
2. Script de sincronización `BASE_FBS` → hojas de sector (por clave).
3. Módulo WhatsApp con plantillas configurables.
4. `REPORTE` de empresa con triggers y botón.
5. Migrar Chexa (ya la más cercana), luego Turin.
6. Libro de grupo.
7. Replicar al resto de las empresas.

## Preguntas abiertas
- ¿Cómo sale hoy el export de FBS (manual, IMPORTRANGE, otro)?
- Horarios exactos de actualización del reporte.
- Lista final de eventos de WhatsApp por sector.
- Qué otras empresas siguen después de Chexa y Turin.
- ¿Este diseño vive en este repo (`mora`) o en uno nuevo?
