# Mora Plan de Ahorro – Tableros FIAT

Script de Google Apps Script (`mora_fiat.gs`) que arma, sobre el Google Sheet **MORA - PDA**, los tableros de seguimiento de mora a partir de la hoja `BASE`.

## Instalación
1. En el Sheet: **Extensiones → Apps Script**.
2. Borrar el código que haya y pegar el contenido de `mora_fiat.gs`. Guardar.
3. Elegir la función `construirTableros` y **Ejecutar** (autorizar la primera vez).
4. Recargar el Sheet: aparece el menú **Mora FIAT** para reconstruir cuando haga falta.

Todas las hojas usan fórmulas: cuando se actualiza `BASE`, los tableros se recalculan solos. El script no modifica `BASE`.

## Hojas que se crean
| Hoja | Qué muestra |
|---|---|
| PARAMETROS | Tabla de umbrales y % de incentivo por cuota (editable) |
| CALC | Una fila por plan: estado actual, cuotas impagas, cuota que mide FIAT y estado FIAT |
| Tablero Medición FIAT | Cuotas 3/5/7/9/12 del mes vencido (avance hoy 4/6/8/10/13, evaluando C2..C(N-1)), % mora, tramo e incentivo |
| Tablero Estado Actual | Por avance: al día, mora, rescindidos, mora vencida vs. solo la cuota del mes, y proyección de la próxima medición |
| Detalle Mora FIAT | Planes que quedaron en mora en la medición del mes vencido |
| Gestión Mes | Planes en mora hoy en avance 3/5/7/9/12 (los que FIAT mide el mes próximo) |

## Reglas
- **Rescindido**: `Estado` = Rescindido o Renunciado.
- **Mora**: al menos una `I` en las cuotas evaluadas. **Al día**: ninguna `I`.
- **% mora** = (en mora + rescindidos) / cartera total (al día + mora + rescindidos).
- **Tramo A** si % mora < umbral A; **Tramo B** si está entre "desde" y "hasta" (inclusive); si no, sin cobro.
