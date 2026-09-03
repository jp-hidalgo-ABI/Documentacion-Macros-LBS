[Volver a la macro de salida](README.md) · [Índice general](../README.md)

# Runbook del operador — macro de salida

El orden de los botones, los diálogos de cada paso, qué revisar antes de avanzar y qué hacer
cuando algo falla.

**La secuencia es lineal y no se puede alterar.** Cada paso consume el resultado del
anterior. En particular, `SummaryOptimizar` corrido dos veces sobre la misma hoja produce
resultados incorrectos.

## Antes de empezar

1. Tener a la mano los cuatro archivos que devolvió LBS y el libro del Plan.
2. Cerrar cualquier otra instancia de Excel con archivos grandes: esta macro consume mucha
   memoria.
3. Revisar `Plan!V1`. En una corrida productiva debe estar vacía o en `NO`. Ver
   [10-diagnostico-y-errores.md](10-diagnostico-y-errores.md#modo-prueba-planv1).

## Regla general: no volver a hacer clic

Mientras una macro corre, la bandera `SK_MacroBusy` bloquea la re-entrada
(`tms_fg14/modulo2.vba:6`). Si se hace clic en un botón mientras otro está corriendo,
aparece:

```
Ya hay una macro LBS en ejecucion. Espera a que termine (mira Plan!X / barra de estado).
```

(`tms_fg14/modulo2.vba:602`, `2515`; `tms_fg14/modulo3.vba:648`)

Esa guarda existe porque sin ella Excel se cerraba sin aviso. El progreso se ve en `Plan!X1`
y en la barra de estado; si un botón parece no responder, ahí está la respuesta.

## Paso 1 — `ProcesoMacro`

`tms_fg14/modulo1.vba:121`

Importa los cinco archivos de entrada. Antes de pedir nada, limpia las cinco hojas destino
(`tms_fg14/modulo1.vba:134-157`).

Después presenta **cinco diálogos en secuencia**, cada uno precedido por un aviso con el
nombre del archivo que espera:

| Orden | Aviso | Archivo que hay que seleccionar | Obligatorio |
|---|---|---|---|
| 1 | `Seleccionar Shipments` | `Shipments (*).xlsx` | Sí |
| 2 | `Seleccionar STP Equipment Associtation` | `STP Equipment Association (*).xlsx` | Sí |
| 3 | `Seleccionar Order Failures` | `Order Failures (*).xlsx` | **No** |
| 4 | `Seleccionar Pallet Container Association` | `Pallet Container Association (*).xlsx` | Sí |
| 5 | `Seleccionar Plan` | `Planificacion_Plantas_Restos_v19.xlsm` | Sí |

El aviso del paso 2 tiene una errata en el código (`Associtation`,
`tms_fg14/modulo1.vba:176`); es cosmética.

**El orden importa.** No hay validación de que el archivo seleccionado sea el correcto: la
macro copia la primera hoja del libro que se elija a la hoja destino que corresponda a ese
paso. Seleccionar el archivo equivocado no produce un error inmediato, solo datos mal
colocados que se manifiestan mucho después.

### Order Failures es opcional

Si se cancela ese diálogo, aparece
`No se seleccionó archivo Order Failures. Se continuará con el siguiente paso.`
(`tms_fg14/modulo1.vba:196`) y el proceso sigue. Cancelar cualquiera de los otros cuatro
**aborta** la importación completa.

Es razonable: `STP Failures` alimenta el remonte de `SummaryFallo`, que es una fase de
recuperación. Sin ese archivo la corrida es válida, solo se pierden oportunidades de rescate.

### El Plan tiene que ser el `.xlsm`, no el export

El diálogo del Plan acepta `.xls`, `.xlsx`, `.xlsm` y `.xlsb`
(`tms_fg14/modulo1.vba:239`), y busca una hoja llamada `Plan` dentro. El archivo MEX KA no
tiene esa hoja, así que falla con:

```
Error al importar Plan (<etapa>):
<descripción> (<número>)
Filas detectadas: N
Sugerencia: use Planificacion_Plantas_Restos.xlsm (no el export MEX KA).
```

(`tms_fg14/modulo1.vba:254-257`). La sugerencia final está ahí porque es la confusión más
frecuente.

Las etapas posibles son `Abrir Plan`, `Detectar filas Plan` y
`Copiar Plan A:AF (N filas)`.

### Qué revisa al terminar

El mensaje de éxito es `Proceso Finalizado` (`tms_fg14/modulo1.vba:268`), y al final activa
la hoja `User Guide`.

Antes de avanzar conviene verificar que las cinco hojas tengan datos y que el conteo de filas
sea razonable. Ver los rangos y ordenamientos que aplica cada importación en
[02-entradas-lbs.md](02-entradas-lbs.md).

### Errores posibles

**`La hoja "<nombre>" reporta N filas (max 50000). Revise celdas sueltas al final de las columnas clave antes de importar.`**
(`tms_fg14/modulo1.vba:33-35`) — Casi siempre es una celda con un espacio muy abajo en el
archivo fuente, que hace que Excel calcule mal la última fila. Abrir el archivo, ir al final
con Ctrl+Fin y borrar las filas sobrantes.

## Paso 2 — `SummaryOK`

`tms_fg14/modulo2.vba:514`

Construye `Pedidos Surtidos` desde `Shipments`, cruzando con `STP-Equipment Assoc`,
`Pallet Container Association` y el `Plan`. Es el paso más largo de todos.

El progreso se ve en `Plan!X1`, con el nombre de la fase en curso.

### Qué revisar al terminar

**La columna `AV` (`REVISION MANUAL`)** es lo primero. Filtrarla y revisar todo lo que no
esté vacío.

**La columna `AG` (motivo)** para entender por qué hay filas en No Planeado.

**La hoja `STP Failures`**, donde la macro deja el detalle de las advertencias.

### Advertencias posibles

**`Proceso completado con N linea(s) sin Armado en Plan. Ver STP Failures.`**
(`tms_fg14/modulo2.vba:1699`) — Hay renglones del Plan sin armado, así que no se pudo
convertir cartonaje a tarimas. Se resuelve del lado de la macro de entrada: agregar la fila en
`ArmadoChep` o correr `SincronizarCatalogoHomologado`, y después rehacer el pipeline
completo.

**`Proceso completado con N camion(es) sobre tope de peso. Ver columna AV y STP Failures.`**
(`tms_fg14/modulo2.vba:1719`) — Uno o más camiones exceden el techo de peso. El tope general
es 29 000 kg (`SK_MAX_PESO_KG`, `tms_fg14/modulo2.vba:10`) y para los carriles Full de
Alsuper y Go Mart es 52 500 kg (`LBS_FULL_MAX_PESO_KG`, `tms_fg14/modulo2.vba:12`). Requiere
decisión manual: bajar carga o dividir el camión.

Ninguna de las dos detiene el proceso. Son avisos.

### Error

**`SummaryOK failed at phase [<fase>]`** (`tms_fg14/modulo2.vba:1741`) — Error inesperado. El
mensaje incluye la fase, el contenido de `Plan!X1` y el número y descripción del error de
VBA. La fase indica qué parte del motor falló.

## Paso 3 — `SummaryOptimizar`

`tms_fg14/modulo2.vba:2502`

Consolida las filas en camiones y aplica los filtros de eficiencia y de cupo.

### La advertencia más importante del runbook

**No correr `SummaryOptimizar` dos veces sobre la misma hoja.** La macro consolida filas y
reasigna folios; una segunda pasada consolida sobre lo ya consolidado y produce camiones
sobrecargados o vacíos.

La misma advertencia aparece en el mensaje de cierre de `SincronizarCatalogoHomologado` de la
macro de entrada (`merged/modulo7.vba:2071`): *"No re-ejecutar Optimizar sobre hoja ya
optimizada"*.

Si se corrió por error, hay que volver al paso 2 y rehacer `SummaryOK`.

### Qué revisar al terminar

El mensaje de éxito es (`tms_fg14/modulo2.vba:2557-2559`):

```
SummaryOptimizar completado.
Filtro por cadena, consolidacion y cupo aplicados.
```

Después de eso:

1. **La columna `AR`** (fill rate) contra los umbrales de la hoja
   `EFICIENCIA POR CADENA`.
2. **La columna `AG`** para ver qué se descartó y por qué. El motivo más común es
   `Descartado por baja eficiencia`.
3. **La columna `AD`** (folio) para revisar cuántos camiones se formaron.
4. **La columna `AV`** de nuevo.

Detalle de los gates y los pisos de llenado en
[05-eficiencia-y-descartes.md](05-eficiencia-y-descartes.md).

### Error

**`SummaryOptimizar failed at phase [<fase>]`** (`tms_fg14/modulo2.vba:2575-2577`) — Mismo
formato que el de `SummaryOK`: fase, `Plan!X1`, número y descripción del error.

### Fases que se pueden correr por separado

| Macro | Ubicación | Cuándo usarla |
|---|---|---|
| `FiltrarPorEficiencia` | `tms_fg14/modulo2.vba:4828` | Para reaplicar solo el filtro de eficiencia tras ajustar los umbrales |
| `ConsolidarNoPlaneados` | `tms_fg14/modulo2.vba:5075` | Para reintentar subir filas No Planeado a camiones con espacio |

Son fases que `SummaryOptimizar` ya ejecuta. Correrlas sueltas es una herramienta de ajuste
fino, no parte del flujo normal.

## Paso 4 — `PartirTarimasFULL`

`tms_fg14/modulo5.vba:1414`

Divide cada camión Full en sus dos cajas, agregando los sufijos `a` y `b` al folio de la
columna `AD`.

### Qué revisar al terminar

**Que cada folio `a` tenga su `b`.** Un folio `a` sin `b` significa que el reparto entre
cajas quedó desbalanceado. El comentario del código señala el caso
(`tms_fg14/modulo2.vba:4767`): *"tratar cada caja como camion de 40 duplica cupo y rompe el
balance a/b (p.ej. a=31 sin b)"*.

**La columna `AV`** una vez más.

### Advertencia

**`Partir Fulles: <mensaje> Ver columna AV y STP Failures.`**
(`tms_fg14/modulo5.vba:1905`) — Hubo desbordamiento al partir: algo no cupo en las dos cajas.
El detalle está en `AV` y en `STP Failures`.

Detalle en [06-partir-tarimas-full.md](06-partir-tarimas-full.md).

## Paso 5 — `SummaryFallo`

`tms_fg14/modulo3.vba:597`

Intenta remontar en camiones con espacio las filas que quedaron fuera: los restos, los
descartes por eficiencia y los pedidos que LBS reportó como fallidos.

### Qué revisa al terminar

El mensaje varía según el resultado (`tms_fg14/modulo3.vba:847-861`):

```
Proceso completado correctamente.

Fallos: 2.4 min
```

o, si hubo problemas de armado:

```
Fallos completado con N linea(s) sin Armado en Plan. Ver STP Failures.

Fallos: 2.4 min
```

El tiempo transcurrido se muestra en minutos si pasa de 60 segundos, y en segundos si no
(`tms_fg14/modulo3.vba:841-845`).

Si el modo prueba está activo, agrega una tercera línea
(`tms_fg14/modulo3.vba:853-856`):

```
TEST MODE (Plan!V1): se ignoro 'Shippable qty is 0' al remountar.
```

**Si aparece esa línea en una corrida productiva, la corrida no es válida.** Significa que
se remontaron filas que LBS había marcado con cantidad embarcable cero.

Detalle en [07-fallos-y-remonte.md](07-fallos-y-remonte.md).

## Paso 6 — `CompararCartonajes`

`tms_fg14/modulo3.vba:886`

Compara el cartonaje de `Pedidos Surtidos` contra el del `Plan` y marca en la columna `AG`
las líneas del Plan que LBS nunca reportó, con el motivo `No reportado por LBS`.

Es un paso de conciliación: identifica lo que se perdió en el camino entre lo que se pidió y
lo que se va a embarcar.

## Paso 7 — `AdicionalesSAP`

`tms_fg14/modulo6.vba:1`

Prepara los adicionales para SAP en la hoja `AdicionalesSAP`, los agrupa con
`ConsolidarItems` (`tms_fg14/modulo6.vba:108`) y los exporta con `ExportarACSV`
(`tms_fg14/modulo6.vba:152`).

### Mensajes

| Mensaje | Significado |
|---|---|
| `El archivo CSV se ha generado en la ruta: <ruta>` (`tms_fg14/modulo6.vba:196`) | Éxito |
| `Error: No se pudo obtener la ruta del archivo.` (`tms_fg14/modulo6.vba:169`) | No se pudo resolver dónde escribir. Revisar que el libro esté en disco local |
| `Error <número>: <descripción>` (`tms_fg14/modulo6.vba:200`) | Error inesperado |

Detalle del formato en
[08-salidas-tms-y-sap.md](08-salidas-tms-y-sap.md#adicionales-sap).

## Paso 8 — `TMS`

`tms_fg14/modulo4.vba:1`

El último paso. Construye la hoja `Data` con el mapeo de campos que espera TMS y llama a
`CreaTMStemplate` (`tms_fg14/modulo4.vba:57`) para escribir el archivo
`ProcessShipmentOrderCreate_DS_*.xls`.

### Antes de correrlo

**Revisar la columna `AV` por última vez.** Todo lo que quede marcado en `REVISION MANUAL` se
va a exportar tal cual a TMS. Es el último punto donde se puede corregir.

Detalle en [08-salidas-tms-y-sap.md](08-salidas-tms-y-sap.md).

## Resumen de la secuencia

| Paso | Macro | Aborta si falla | Qué revisar después |
|---|---|---|---|
| 1 | `ProcesoMacro` | Sí | Que las 5 hojas tengan datos |
| 2 | `SummaryOK` | No (avisa) | `AV`, `AG`, `STP Failures` |
| 3 | `SummaryOptimizar` | No (avisa) | `AR`, `AG`, `AD`, `AV` |
| 4 | `PartirTarimasFULL` | No (avisa) | Que cada `a` tenga su `b`, `AV` |
| 5 | `SummaryFallo` | No (avisa) | El mensaje de tiempo y la línea de modo prueba |
| 6 | `CompararCartonajes` | No | Motivos `No reportado por LBS` en `AG` |
| 7 | `AdicionalesSAP` | Sí | La ruta del CSV |
| 8 | `TMS` | Sí | `AV` antes de correrlo |

## Las tres columnas que hay que mirar siempre

| Columna | Nombre | Qué buscar |
|---|---|---|
| `AG` | Motivo | Por qué una fila quedó fuera o qué marca especial lleva |
| `AR` | Fill rate | El porcentaje que reportó LBS, contra el umbral de la cadena |
| `AV` | `REVISION MANUAL` | Todo lo que el motor no pudo resolver solo |

Mapa completo en [03-pedidos-surtidos.md](03-pedidos-surtidos.md).

## Herramientas cuando algo no cuadra

| Situación | Herramienta |
|---|---|
| Dos filas que deberían ir en el mismo camión quedaron separadas | `LBS_DiagnosticarConsolidacion` (`tms_fg14/modulo2.vba:22403`) |
| La hoja `Consolida` está vacía o desactualizada | `LBS_SeedConsolidaSheet` (`tms_fg14/modulo2.vba:5407`) |
| Las alturas de `AO`/`AP` salen mal | `LBS_SyncTihiSheet` (`tms_fg14/modulo2.vba:8239`) |
| Se editó `Consolida` a media corrida | `LBS_ResetConsMap` (`tms_fg14/modulo2.vba:5256`) |

Detalle en [10-diagnostico-y-errores.md](10-diagnostico-y-errores.md).
