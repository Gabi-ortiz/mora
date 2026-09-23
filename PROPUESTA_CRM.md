# CRM de gestión de mora — propuesta inicial

Estado: **v1 implementada** (`crm_mora.gs` + `crm_index.html`, instalación
en el README). Se apoya en el tablero de mora que ya existe en este repo
(`tablero_mora.gs`).

## Objetivo
Que cada responsable (hoy **Eze**, **Fabio** y **Santi**, columna A
"RESPONSABLE" de BASE) trabaje su cartera **al día, sobre el mes en curso**,
dejando registro estructurado de cada intento de contacto y del estado de
cada caso. Nosotros seguimos midiendo a mes vencido (como mide Fiat) con el
Tablero, y además podemos ver cuánto trabajo hizo cada uno y si se traduce
en baja de mora.

## Qué hay hoy en BASE (observado en el archivo)
- Col A **RESPONSABLE** (EZE / FABIO / SANTI), B Solicitud, C Grupo,
  H Cliente, I Teléfono, J Tel. alternativo, N Avance, O Estado,
  R..AD = C2..C14 con P / I / R.
- Al final hay columnas de texto libre por mes (notas tipo "6/3 No me
  responde llamada, le hablo por wsp…") + "ACT MOTIVO NO PAGO" /
  "Motivo completo". Es exactamente el registro que el CRM tiene que
  reemplazar por algo estructurado (fecha, canal, resultado, estado).
- BASE llega por IMPORTRANGE → **no se puede escribir ahí**. Las gestiones
  tienen que vivir en hojas propias, ligadas al plan por **Solicitud**.

## La lógica de prioridad (clave)
Fiat mide el mes siguiente la cuota N de los planes que estén en avance
N+1 (rango C2..CN). Entonces, **hoy**, los planes en avance
**3, 5, 7, 9 y 12** son los que el mes que viene se van a medir por las
cuotas 3, 5, 7, 9 y 12 — y lo que se haga este mes todavía cambia ese
número. Para esos planes el rango que va a contar es C2..C(avance actual),
o sea **incluye la cuota del mes en curso** (que hoy suele estar vacía hasta
que se paga).

Propuesta de cola de trabajo por responsable, en este orden:

| Prioridad | Qué | Por qué |
|---|---|---|
| **P1** | Avance ∈ {3,5,7,9,12} con alguna `I` en C2..C(avance-1) | Ya cuentan como mora el mes que viene si no se regularizan |
| **P2** | Avance ∈ {3,5,7,9,12} con la cuota del mes (C(avance)) sin pagar | Todavía no son mora; hay que asegurar que paguen antes del corte |
| **P3** | Resto de la cartera con `I` | Mora general de la empresa |
| — | Rescindidos (`R`) | Fuera de la cola (no se recuperan), salvo que definamos lo contrario |

Dentro de cada prioridad: primero los que hace más días que no se
contactan / nunca se contactaron este mes.

## Qué ve cada responsable (app)
Propuesta: **Web App de Google Apps Script** colgada del mismo archivo.
Ventajas: no hay que instalar nada, entran con su cuenta de Google, lee
BASE directo, guarda en hojas del mismo archivo, costo cero, y reusa el
código que ya tenemos (`clasificarFila`, config de columnas).

Pantallas:
1. **Mi cartera** — lista de sus casos ordenada por prioridad, con chips:
   avance, cuota que mide Fiat, estado del caso, último contacto, "sin
   gestión este mes". Filtros por prioridad / estado / avance.
2. **Ficha del plan** — datos del cliente (teléfonos con botón
   WhatsApp/llamar), grilla de cuotas C2..C14 coloreada P/I/R, historial
   completo de gestiones, y formulario rápido para registrar una gestión:
   - Canal: Llamada / WhatsApp / SMS / Mail
   - Resultado: Atendió / No atendió / Número erróneo / Mensaje enviado…
   - Estado del caso (a definir, ver abajo)
   - Fecha de promesa de pago (si aplica) + monto
   - Motivo de no pago (lista cerrada, reemplaza "ACT MOTIVO NO PAGO")
   - Nota libre
   - Próximo contacto (fecha) → alimenta la agenda
3. **Agenda de hoy** — promesas que vencen hoy y re-contactos agendados.
4. **Tablero de avance** (para nosotros) — por responsable: casos
   asignados, % contactados en el mes, intentos por caso, promesas
   cumplidas vs incumplidas, casos P1/P2 regularizados, y proyección del
   % de mora que va a medir Fiat el mes que viene (mismo cálculo del
   Tablero pero sobre avance actual y rango C2..C(avance)).

## Hojas nuevas (se crean solas)
- **Gestiones** (append-only, una fila por intento): FechaHora,
  Responsable (email), Solicitud, Canal, Resultado, EstadoCaso,
  MotivoNoPago, FechaPromesa, MontoPromesa, ProximoContacto, Nota,
  Avance y Clasificación al momento de la gestión (foto para medir
  después).
- **Casos** (una fila por Solicitud, estado vigente): EstadoCaso actual,
  última gestión, próximo contacto, promesa vigente. Se actualiza con cada
  gestión; evita recalcular todo el historial en cada carga.
- **Usuarios**: email ↔ nombre de responsable (EZE/FABIO/SANTI) y rol
  (responsable / supervisor). El supervisor ve todas las carteras.

## Estados del caso (a definir juntos — propuesta para arrancar)
Sin gestionar · Contactando (sin respuesta) · Contactado · Promesa de pago ·
Pagó (a verificar) · Regularizado (confirmado por P en BASE) ·
Negativa / no quiere pagar · Inubicable · Derivado (legal / supervisor).

"Regularizado" idealmente se marca **solo** cuando BASE muestra la `P`,
para que el estado no dependa solo de lo que carga el responsable.

## Plan de trabajo sugerido
1. **MVP (1ª iteración)**: hojas Usuarios/Gestiones/Casos + Web App con
   "Mi cartera" (priorizada) y "Ficha" con registro de gestión.
2. Agenda de hoy + promesas de pago con vencimiento.
3. Tablero de avance por responsable + proyección de mora del mes próximo.
4. Migrar (opcional) las notas mensuales de texto libre de BASE como
   gestiones históricas "importadas".

## Decisiones tomadas (23/09/2026)
- **Roles**: uno o más SUPERVISORES con control total (usuarios,
  asignación/reasignación, cualquier acción sobre cualquier plan, tablero
  "Supervisor"). RESPONSABLES (Godoy Santiago sgodoy@, Vaca Ezequiel evaca@,
  Aguero Fabio faguero@ — grupoantun.com.ar) ven solo su cartera.
- **Cuota del mes**: BASE la actualiza el supervisor a mano desde el
  sistema de Fiat; la cuota del mes queda `I` (o vacía) hasta que paga. El
  CRM trata vacío e `I` igual: "no paga todavía".
- **Asignación**: la hace el supervisor. Por defecto sale de la columna
  RESPONSABLE de BASE; las reasignaciones hechas desde el CRM tienen
  prioridad.
- **Rescindidos**: fuera de la cola de trabajo (se trabaja solo la cartera
  activa), pero visibles en la vista "Rescindidos / bajas".
- **Actualización de BASE**: manual por el supervisor; más adelante se
  evalúa automatizar. El CRM lee BASE en vivo en cada carga, así que ve la
  última actualización apenas se hace.

## Pendiente de definir
- Lista final de estados del caso, resultados de contacto y motivos de no
  pago (hoy son una propuesta, se cambian en `crm_mora.gs`).
- Qué más mostrar en el Tablero supervisor (hoy: proyección de lo que
  mide Fiat por cuota, avance por responsable, gestiones por día).
- Si Renunciado/Cancelado deben quedar fuera de la cartera activa como
  hoy (`CRM_ESTADOS_BAJA`).
- Metas por responsable (ej. % de P1+P2 contactados por semana).
- Migrar las notas mensuales de texto libre de BASE como gestiones
  (hoy se muestran como "Notas anteriores de la planilla" en la ficha).
