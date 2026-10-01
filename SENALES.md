# Señales comerciales — lector de carpeta y mapa de la política comercial

## Carpeta
https://drive.google.com/drive/folders/103O9LnAq7tljmQpwdFl7dUlwkecehcd_

De ahora en más todas las señales nuevas se suben ahí. Al 01/10/2026 tiene
33 archivos (señales Fiat / Jeep / RAM Plan, tablas de cuotas y sellados,
opcionales, listas de cambio de modelo). Seguramente faltan señales.

## Lector / alerta (`senales.gs`)
Script que va en un Google Sheet (el futuro archivo de objetivos e
incentivos, o uno propio):
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
- **Mail** con los archivos nuevos cada vez que aparece algo (la primera
  corrida manda un único mail con todo lo registrado).

Instalación:
1. En el Sheet: Extensiones > Apps Script, pegar `senales.gs`, guardar.
2. Recargar: aparece el menú **Señales**.
3. Señales > **Revisar carpeta ahora** (pide permisos de Drive y Gmail).
4. Señales > **Instalar revisión automática (cada hora)**.

Si el script va en el mismo proyecto de Apps Script que otro (por ejemplo
Pedidos PDA), hay que unir los `onOpen` en uno solo.

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
