# Señales comerciales — lector de carpeta y mapa de la política comercial

## Carpeta
https://drive.google.com/drive/folders/103O9LnAq7tljmQpwdFl7dUlwkecehcd_

De ahora en más todas las señales nuevas se suben ahí. Al 01/10/2026 tiene
33 archivos (señales Fiat / Jeep / RAM Plan, tablas de cuotas y sellados,
opcionales, listas de cambio de modelo). Seguramente faltan señales.

## Lector / alerta (`senales.gs`)
Script que va en el archivo de objetivos e incentivos (ver OBJETIVOS.md),
junto con `objetivos.gs`, que es el que arma el menú:
- Hoja **SEÑALES**: un registro por archivo de la carpeta (y subcarpetas),
  con Nº, marca, tipo y mes detectados del título, link, quién lo subió y
  una columna **Estado** (Pendiente / Leída / Cargada) para saber qué señal
  ya se volcó al archivo de objetivos.
- Hoja **FALTANTES**: números de señal que no están en la carpeta, desde la
  1500 hasta la última. La numeración es **una sola secuencia para Fiat,
  Jeep y RAM** (1599 Fiat, 1600 Jeep, 1603 Fiat, 1604 Jeep), así que un
  faltante puede ser de otra marca o de otro tema. Sirve como lista para
  pedirle al zonal. Hoy, entre la 1527 y la 1604, hay 27 en la carpeta y
  faltan 51.
- **Mail** a gortiz@grupoantun.com.ar (`MAIL_AVISO`) con los archivos nuevos
  cada vez que aparece algo (la primera corrida manda un único mail con todo
  lo registrado).

### Lectura del contenido (versión 2)
Además del nombre, el script **lee el archivo** (PDF e imágenes con OCR de
Google Drive; PowerPoint convirtiéndolo a Slides). Las tablas de cuotas,
opcionales y listas de cambio de modelo no se leen.
- **Carta de objetivos** (se reconoce por el texto "CARTA DE OBJETIVOS",
  aunque el archivo se llame "TURIN S.A_3.pdf"): carga sola en OBJETIVOS mes,
  concesionario, marca, categoría, cada indicador con su número, Nº y fecha
  de carta, fecha y % de flujo (en la fila de SUSCRIPCIONES). Si el mes ya
  estaba cargado y un número cambió, lo actualiza y anota el valor anterior
  en Notas. Indicadores nuevos se agregan solos a LISTAS_OBJ. La carta queda
  en Estado = "Cargada" y el mail muestra lo cargado para controlarlo.
- **Señal comercial**: columna **Afecta** (Objetivos, Incentivos, Categoría,
  Flujo, Mora / permanencia, Pedidos) y columna **Extracto** con los temas
  del "REF:" en lista corta (• Liquidación de flujo septiembre 2026 • Bonus
  cumplimiento pedidos totales ...) + el primer dato con % o $. Se limpia
  el ruido del OCR (encabezados "T O D O S L O S ...", mails, texto de
  cortesía). `rehacerExtractos` (desde el editor; ya no está en el menú)
  vuelve a leer las señales (no las cartas) con este formato. Los % de incentivo se cargan a mano en INCENTIVOS
  (cada señal tiene otro formato y un número mal leído afecta la
  liquidación).
- **Tabla desarmada por el OCR**: en la carta 09/2026 real, Google devuelve
  los nombres de los indicadores al final y pega 43 y 60 en "4360". En ese
  caso el lector reconoce la carta (mes, categoría, flujo, Nº, indicadores)
  pero NO adivina números: deja el objetivo vacío con "⚠ Completar
  objetivo", la carta queda "Pendiente" y el mail lo avisa. Si el mes ya
  tenía los objetivos cargados, no los toca.
- **Indicadores pegados por el OCR** ("PATENTAMIENTOS PEDIDOS TOTALES" en un
  renglón): se separan usando los indicadores de LISTAS_OBJ. "Armar /
  reparar estructura" borra las filas sin número que haya dejado una lectura
  anterior con el indicador pegado (y lo saca de LISTAS_OBJ).
- **Asignación por orden** (nombres en un renglón y números en otro): solo si
  es una tabla limpia, con los nombres en renglones seguidos y los números
  justo debajo, también seguidos. Si el OCR los mezcló, no se asigna nada
  (con el texto real de la carta 09/2026, mezclar el orden daba
  suscripciones = 43).
- **Números pegados por el OCR** ("4360" = 43 y 60; "2727454" = 27, 27, 4 y
  54): `deducirNumeros` prueba todas las formas de repartirlos entre los
  indicadores y se queda con las que cumplen: cada objetivo entre 1 y 999,
  ninguno mayor que SUSCRIPCIONES y PEDIDOS TOTALES ≥ suma de los PEDIDOS
  por modelo (Cronos y Titano son parte del total). Si queda una sola, se
  usa; si quedan varias, se elige la más parecida a los objetivos del mes
  cargado más cercano, solo si se diferencia claramente. Lo deducido queda
  "⚠ Verificar" (Estado Pendiente). Probado: carta 09/2026 → 146/43/60;
  carta 07/2026 → 127/27/27/4/54.
- Lee hasta 12 archivos por corrida (límite de 6 minutos de Apps Script);
  los 33 ya registrados se van leyendo en las próximas corridas automáticas
  o con `leerSenalesRegistradas` desde el editor.
- Probado contra el texto real de la Carta 09/2026 y contra un formato
  "tipo OCR" (nombres y números en líneas separadas).

Instalación: ver OBJETIVOS.md (requiere activar el servicio avanzado
**Drive API**). Desde el menú **Objetivos**: "Revisar
carpeta de señales ahora". La revisión automática cada hora ya está
instalada (`instalarTriggerSenales`, desde el editor si hubiera que
reinstalarla).

## Lo que dicen las señales (resumen para el dashboard)
Fuente principal: Señal madre **1065/2021** (Nueva Política Comercial Fiat
Plan) + actualizaciones mensuales. Algunos PDF son imágenes y el texto no se
pudo extraer completo (p. ej. tablas de 1576); revisar contra el PDF.

### Categoría ABC (trimestral, por mora)
- Señal **1595/2026** (vigencia OCT-NOV-DIC 2026): **mora cuota 6**,
  A ≤ 43%, C ≥ 51%, B el resto. Cierre de medición 10/09/2026 sobre
  vencimientos cuota 6 de MAY-JUN-JUL. (La 1065 original sumaba NPS; la
  1595 ya no lo menciona.)
- ⚠ El tablero de mora hoy sigue las cuotas 3, 5, 7, 9 y 12 (incentivo de
  permanencia) pero **no la cuota 6**, que es la que define la categoría.

### Objetivos mensuales (carta de objetivos)
- Suscripciones, Patentamientos y Pedidos totales, según el peso del CC
  sobre la marca a nivel país (lo define Fiat).
- **Flujo**: un día del mes (en septiembre, 23/09) hay que tener ingresado
  el 65% o más del objetivo de **suscripciones**.

### Incentivos de suscripción (1065, por categoría)
Según % de cumplimiento del objetivo de suscripciones (OUT < 50%, Poco
satisfactorio 50–89,9%, Satisfactorio 90–109,9%, Extra bonus 110–129,9%,
Super objetivo ≥ 130%): incentivo base + adicionales por flujo, ratio de
débito automático y débito automático plus. B y C que quedan OUT no cobran
captura; C tiene volumen máximo (110%). Recupero de incentivo si la
solicitud se rescinde/renuncia sin 2 cuotas pagas (A 50%, B 70%, C 80%).

### Pedidos (septiembre 2026: señales 1599 y 1603)
- **2A Bonus cumplimiento pedidos totales** (≥ 100% del objetivo de
  pedidos): A 1,00%, B 0,50%, C 0,25%. Requiere 100% del objetivo de
  suscripciones.
- **2B Pedidos adicionales al objetivo**: A 0,50%, B 0,25%, C 0,10% por
  cada pedido excedente, liquidado sobre precio concesionario sin IVA del
  NC1 (Cronos Drive Plus).
- **2C Extra bonus pedidos**: % por modelo (Cronos, Fiorino, Fastback,
  Strada, Titano, Pulse, Toro...). La 1603 bajó la condición de 100% a
  **80% del objetivo de patentamientos**. Excluye Cronos 90/10 B90 desde
  grupo 18002.

### Otros
- **Incentivo de permanencia** (1576, JUL-SEPT 2026): por mora en cuotas 3,
  5, 7, 9 y 12 (es lo que mide el tablero de mora).
- **Bonus de entrega** por NPS mensual (1065).
- **Margen comisional**: 60% al aceptar el pedido, 40% al ingresar el REU
  (recupero si no hay REU a los 60 días).
