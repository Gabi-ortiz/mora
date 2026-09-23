# Tablero de control de mora — Fiat (planes de ahorro)

## Archivo de trabajo
https://docs.google.com/spreadsheets/d/12pLWcN7GLbmD7rBbeEdg_3dXHMWrQL88AtsGj4fH5wM

Archivo nuevo, armado desde cero por una cuestión de organización (reemplaza a
uno anterior llamado "Analisis mora GABI" que tenía todo mezclado: base +
cuadro de análisis en el mismo archivo, hoja "Hoja 2" para la base).

## Hojas
- **BASE**: la tabla cruda de planes, un plan por fila. La actualiza (los
  valores P/I/R) un compañero de trabajo. Llega por IMPORTRANGE desde otro
  archivo. **Confirmado contra el archivo actual** (21/09/2026): la hoja se
  llama exactamente "BASE", tiene 1858 planes cargados y las columnas caen
  en las mismas posiciones que en el archivo viejo (ver más abajo) — no hace
  falta tocar la configuración del script.
- **Historial_Mora**: log append-only, una fila por (fecha de snapshot,
  avance). Se genera con el script — no se edita a mano. Todavía no existe
  en el archivo nuevo, la crea el script solo la primera vez que corrés
  "Sacar foto ahora".
- **Tablero**: vista final con una fila por cada cuota de interés (3, 5, 7,
  9, 12), tomando siempre la última foto disponible del historial para el
  avance que corresponde a cada cuota. Se regenera con el script — no se
  edita a mano. Tampoco existe todavía, la crea "Actualizar tablero".
- **Log_Snapshot**: una línea por corrida del trigger diario, diga si hubo
  cambios en BASE o no. Tampoco existe todavía.
- **Detalle_Planes**: un plan por fila, leído en vivo de BASE (no del
  historial), con filtro nativo de Sheets ya armado para poder filtrar por
  Avance y ver qué planes puntuales caen en cada estado (Mora irregular,
  Rescindido, Renunciado, etc.). Se regenera con "Actualizar detalle de
  planes" — no se edita a mano. Tampoco existe todavía.

## Estructura de la hoja BASE (confirmada contra el archivo actual)
- Columna **N** (14): "Avance" — cuántas cuotas van de ese plan (entero,
  2 a 13 en los datos actuales).
- Columna **O** (15): "Estado" — texto: Ahorrista / Adjudicado / Rescindido /
  Renunciado / Cancelado.
- Columna **R** (18) en adelante: una columna por cuota, empezando en **C2**
  (columna R = C2, S = C3, T = C4, ... hasta AD = C14), con el valor cargado
  por el compañero:
  - `P` = cuota paga
  - `I` = cuota con pago irregular / en mora
  - `R` = plan rescindido (a partir de la cuota en que se rescindió, se
    completa con R en todas las columnas siguientes)

Estas posiciones son idénticas a las del archivo anterior, así que
`COL_AVANCE`, `COL_ESTADO` y `COL_C2` en `tablero_mora.gs` quedan tal cual
están (14, 15 y 18 respectivamente) — no requieren ajuste.

**Sobre "de dónde leer" por fila**: la columna "Avance" define cuántas
columnas de cuota tiene cargadas ese plan puntual (para avance = 9, el rango
cargado va de columna R/C2 a columna Y/C9). Esto es la descripción de la
fila individual, no confundir con la lógica de "mes atrasado" del punto
siguiente, que usa un rango distinto (con margen de un avance) y ya está
probada.

## Lógica de negocio (ya validada contra datos reales, no re-derivar de cero)

### 1) "Mes atrasado": qué cuota mide cada avance
El análisis de una cuota N se hace mirando los planes cuyo **Avance = N + 1**
(un avance de diferencia, para dar margen a correcciones tardías de pago), y
dentro de esas filas se mira el rango de columnas **C2 hasta C(N)** — es
decir, `ncols = avance - 2` columnas a partir de la columna C2 (columna R).

Ejemplo probado: avance = 13 → analiza cuota 12 → rango C2:C12 (hasta columna
AB), **nunca** incluye C13/AC. Esto es intencional y ya está confirmado
correcto — no es un bug a corregir.

### 2) Clasificación por plan (fila), dentro del rango de columnas del punto 1
- **Rescindido**: al menos una columna del rango = `R` (no hace falta que
  sean todas).
- **Mora irregular / Impagos**: al menos una columna del rango = `I`.
- **Pagado al día**: TODAS las columnas del rango = `P`. Se separa en:
  - **Pagos Adjudicados**: de esos, los que tienen Estado = "Adjudicado".
  - **Pagos Ahorristas**: de esos, los que tienen Estado = "Ahorrista".
- **Cartera total** = cantidad de filas con ese Avance.
- **Cartera activa** = Cartera total − Rescindidos.
- **TOTAL PAGOS** = Pagos Adjudicados + Pagos Ahorristas.
- **res+irre+imp** = Rescindidos + Impagos (mora irregular).
- **% de mora** = res+irre+imp / Cartera total.

Estas fórmulas fueron validadas fila por fila contra el cuadro manual
original ("Cuota 3/5/7/9/12 AGOSTO") y coinciden exacto en Rescindidos e
Impagos en los 5 casos, y en Pagos Adjudicados/Ahorristas en 3 de 5 (la
diferencia en los otros 2 es por drift normal de Estado entre el momento en
que se cargó el cuadro a mano y el momento en que se corrió el script — no
es un error de fórmula).

### 3) Por qué existe Historial_Mora (y no se calcula todo en vivo)
La columna "Avance" es LIVE: un plan que hoy está en avance 10 puede pasar a
avance 11 en cualquier momento (aprox. una vez por mes, fecha no siempre
predecible). Si el cuadro se recalculara en vivo contra BASE, un recálculo
tardío daría un número distinto al que fue el corte real. Por eso:
- Un snapshot diario (trigger) recorre TODOS los valores de Avance presentes
  en BASE y guarda una fila por avance en Historial_Mora, con la fecha.
- El snapshot solo escribe si detecta cambios reales en BASE desde la última
  vez (hash MD5 sobre Avance + Estado + C2..C14 de todas las filas — ignora
  a propósito columnas irrelevantes como teléfono/vendedor). Si no hay
  cambios, no duplica filas; solo deja constancia en Log_Snapshot.
- El Tablero nunca lee BASE directamente: siempre lee la última foto
  disponible en Historial_Mora para el avance que corresponde a cada cuota
  de interés (avance = cuota + 1).

### 4) Cuotas que se siguen en el Tablero
3, 5, 7, 9 y 12 (configurable en `CUOTAS_TABLERO`).

### 5) Detalle_Planes: de los números del Tablero a los planes concretos
El Tablero da agregados (cuántos rescindidos, cuántos en mora, etc.) para
las 5 cuotas de interés, pero para trabajar la cartera (llamar, reclamar,
etc.) hace falta saber QUIÉNES son. Para eso está Detalle_Planes:
- Lee BASE en vivo (no el historial) — acá interesa la situación actual de
  cada plan, no una foto vieja.
- Clasifica cada plan con la misma regla de "mes atrasado" que usa el
  Tablero (rango C2..C(avance-1)): Rescindido / Mora irregular / Pagado al
  día / Sin cuotas para analizar (avance=2, todavía no hay rango que mirar).
- Muestra también el Estado tal cual está en BASE (Ahorrista / Adjudicado /
  Rescindido / Renunciado / Cancelado) en una columna aparte — un plan puede
  ser "Renunciado" en Estado y "Pagado al día" en la clasificación de mora
  al mismo tiempo, son dos cosas distintas.
- Trae el filtro de Sheets ya activado (ícono de embudo en el encabezado):
  filtrás por columna "Avance" para pararte en la misma cuota que estás
  mirando en el Tablero (recordá: avance = cuota + 1), y después por
  "Clasificación mora" (Rescindido / Mora irregular) o por "Estado"
  (Renunciado) para ver la lista de planes de ese grupo puntual.

## Script actual
Ver `tablero_mora.gs` en este repo — es la versión funcionando, probada e
instalada como Apps Script bound al archivo anterior, con la configuración
ya confirmada contra el archivo nuevo. Para instalarlo:
1. Abrir el archivo → Extensiones > Apps Script.
2. Borrar el contenido de `Code.gs` (o el archivo que haya por defecto) y
   pegar el contenido de `tablero_mora.gs`.
3. Guardar y recargar la hoja de cálculo (para que aparezca el menú "Mora").
4. Menú "Mora" > "Instalar snapshot automático diario" (una sola vez).
5. Menú "Mora" > "Sacar foto ahora" para la primera foto.
6. Menú "Mora" > "Actualizar tablero" para generar el Tablero por primera vez.
7. Menú "Mora" > "Actualizar detalle de planes" para generar Detalle_Planes.

De ahí en más el trigger diario (~4am) se encarga de sacar la foto solo
cuando hay cambios reales en BASE; "Actualizar tablero" y "Actualizar
detalle de planes" se pueden correr cuando se quiera (no dependen del
trigger) para refrescar cada vista.

## Historial de decisiones de diseño (para no repreguntar)
- Se probó y descartó comparar "avance actual" contra "avance actual - 1"
  para chequear variación: NO son la misma cuota (miden C2:C(avance-1) cada
  uno), así que es normal que difieran — no indica error.
- Se sacó la columna "PeriodoReal" de Historial_Mora (generaba ruido, y no
  aplica limpio a un historial que loguea TODOS los avances, no solo los de
  interés).
- Se sacó la columna "Mes análisis" del Tablero (redundante con "Fecha
  foto").
- En el Tablero, "Período real" se muestra como etiqueta legible tipo
  "Cuota 3 AGOSTO" (derivada de la fecha de la foto), no como fecha cruda,
  para no confundir con la columna "Avance".

## CRM de gestión de mora (`crm_mora.gs` + `crm_index.html`)
App web para que cada responsable trabaje su cartera al día y deje registro
de cada gestión, y para que los supervisores tengan control total. Diseño y
decisiones en `PROPUESTA_CRM.md`.

### Acceso
- Se entra con **mail @grupoantun.com.ar + clave del CRM**. Cualquier otro
  dominio se rechaza (al ingresar y al dar de alta usuarios).
- Todo usuario nuevo, y todo blanqueo, queda con la clave inicial
  **Turin3800** y la app obliga a cambiarla en el primer ingreso (mínimo 8
  caracteres, distinta de la inicial).
- Las claves se guardan con hash + sal en `CRM_Usuarios` (nunca en texto
  plano, y nunca viajan al navegador).
- Sesión de 6 h. 5 intentos fallidos bloquean ese mail 15 minutos (un
  blanqueo lo desbloquea). Blanquear o desactivar a un usuario le corta
  la sesión en el momento.

### Roles
- **SUPERVISOR** (al inicio: Guillermo Rietschi grietschi@ y Gabriel Ortiz
  gortiz@): ve todas las carteras, reasigna planes (de a uno desde la
  ficha, o varios con los checkboxes), registra gestiones y cambia el
  estado de cualquier plan, administra usuarios (alta de supervisores o
  responsables, rol, alias, activar/desactivar, **blanquear clave**) y ve
  el "Tablero supervisor". Siempre tiene que quedar al menos un
  supervisor activo.
- **RESPONSABLE** (Godoy Santiago sgodoy@, Vaca Ezequiel evaca@, Aguero
  Fabio faguero@): ve y gestiona solo su cartera. El servidor lo controla
  en cada llamada, no solo la pantalla.

### Cómo se asigna la cartera
1. Si el supervisor reasignó el plan desde el CRM → ese responsable
   (queda en `CRM_Casos.ResponsableEmail`).
2. Si no → la columna A "RESPONSABLE" de BASE (EZE/FABIO/SANTI), cruzada
   con el "Alias BASE" de cada usuario en `CRM_Usuarios`.

### Prioridad (situación de cada plan, calculada en vivo desde BASE)
Un plan hoy en avance N tiene cargadas C2..C(N-1); la cuota del mes en
curso es C(N) (vacía o `I` hasta que paga). El mes que viene pasa a avance
N+1 y, si N ∈ {3,5,7,9,12}, Fiat mide C2..C(N).
- **P1**: avance ∈ {3,5,7,9,12} con alguna `I` en C2..C(N-1).
- **P2**: avance ∈ {3,5,7,9,12}, vencidas pagas, cuota del mes sin `P`.
- **P3 / P4**: lo mismo en el resto de los avances.
- **Al día**: todo `P`, incluida la cuota del mes.
- **Rescindido / Baja**: alguna `R` o Estado Rescindido/Renunciado/
  Cancelado → fuera de la cola, solo consulta en "Rescindidos / bajas".

Dentro de cada prioridad se ordena por última gestión (los nunca
gestionados primero).

### Hojas que crea (no editar a mano salvo CRM_Usuarios si hace falta)
- **CRM_Usuarios**: Email, Nombre, Rol, AliasBase, Activo.
- **CRM_Casos**: una fila por plan gestionado/reasignado con el estado
  vigente (estado del caso, último contacto, próximo contacto, promesa).
- **CRM_Gestiones**: log append-only, una fila por intento de contacto,
  con avance y situación del plan al momento de la gestión.
- **CRM_Auditoria**: acciones de supervisor (reasignaciones, usuarios).

### Instalación
El proyecto de Apps Script queda con **3 archivos separados** (no pegar
todo en uno):

| Archivo en Apps Script | Tipo | Contenido |
|---|---|---|
| `Code.gs` (el que ya existe) | Script | `tablero_mora.gs` (reemplazar todo: solo cambió el menú) |
| `crm_mora` | Script (＋ > Secuencia de comandos) | `crm_mora.gs` |
| `crm_index` | HTML (＋ > HTML) | `crm_index.html` — el nombre tiene que ser exactamente `crm_index` |

1. Crear/pegar los tres archivos como en la tabla y guardar.
2. Recargar la planilla → menú "Mora" > "CRM: inicializar hojas y
   usuarios" (pide autorización la primera vez). Crea las hojas CRM_* y
   los 5 usuarios iniciales con la clave Turin3800.
3. Implementar > Nueva implementación > engranaje > "Aplicación web":
   - Ejecutar como: **Yo**.
   - Quién tiene acceso: **Cualquier persona** si el dueño del script es
     una cuenta @gmail.com; si es una cuenta @grupoantun.com.ar se puede
     elegir "Cualquier usuario de grupoantun.com.ar" como capa extra.
     En los dos casos la app pide mail @grupoantun.com.ar + clave.
   - Compartir con los usuarios la URL que termina en `/exec`.
4. Cada cambio de código posterior: Implementar > Administrar
   implementaciones > lápiz > Versión: "Nueva versión" > Implementar (la
   URL no cambia).
