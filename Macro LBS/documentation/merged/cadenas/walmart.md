[Volver a cadenas](README.md) · [Macro de entrada](../README.md) · [Índice general](../../README.md)

# Walmart — macro de entrada

Cubre `WAL MART SC` (SuperCenter) y `WAL MART BA` (Bodega Aurrera).

## Identificación

| Regla | Cómo se detecta | Cita |
|---|---|---|
| Cadena Walmart | `LEFT(LCase(Plan!G), 3) = "wal"` | `merged/modulo1.vba:460` |
| Sufijo del item | Coincidencia exacta con `WAL MART SC` o `WAL MART BA` | `merged/modulo1.vba:967-968` |
| SKU de excepción | Lista de 14 SKU base en `LBS_IsWalmartExceptionSkuBase` | `merged/modulo1.vba:1016-1033` |
| Item de split de restos | Sufijo `_WAL` o `_WAL_*` sobre un SKU de excepción | `merged/modulo1.vba:1038-1053` |

Nótese la asimetría: la detección de cadena para el `Group Id` usa un prefijo laxo
(`wal`), mientras que el sufijo del item exige el nombre exacto. Un `Plan!G` con
`"WALMART SC"` (sin espacio) dispararía la lógica de EXC28 pero recibiría el sufijo genérico
`LEFT(cadena,3)` = `WAL`, que casualmente coincide con el de SuperCenter. Con
`"WAL MART BA"` mal escrito el resultado sería incorrecto.

## Sufijos

| Cadena | Sufijo | `Item Id` |
|---|---|---|
| `WAL MART SC` | `WAL` | `<SKU>_WAL` |
| `WAL MART BA` | `WAL_BA` | `<SKU>_WAL_BA` |

## La excepción EXC28

Es la regla más importante de Walmart en esta macro.

### Los 14 SKU

`merged/modulo1.vba:1027-1031`

```
3010443  3000461  3005293  3017917  3017868  3017867  3002998
3006108  3017869  3017871  3009696  3008239  3008268  3008947
```

`LBS_IsWalmartExceptionSkuBase` acepta tanto el SKU desnudo como el `Item Id` completo:
si el texto tiene un guion bajo, toma la parte anterior como base
(`merged/modulo1.vba:1021-1026`).

### La condición

`merged/modulo1.vba:458-476`. Se activa cuando se cumplen las tres condiciones a la vez:

1. La cadena empieza con `wal`.
2. El SKU está en la lista de 14.
3. `Plan!N = 28` **y** `Plan!O = 0`, es decir, el renglón es exactamente 28 tarimas completas
   sin restos.

Cuando se cumple, el `Group Id` del STR (columna `J`) se reescribe como:

```
EXC28_<Stock Transport Request ID>
```

El comentario del código explica el objetivo (`merged/modulo1.vba:458-459`):

```
' Override: Walmart exception SKU fully built (N=28, O=0) gets unique EXC28_<STR id>
' so LBS cannot mix/split that 28-pallet truck with other STP lines. Do NOT set Required Equipment.
```

En español: un camión de 28 tarimas de un solo SKU es un camión perfecto y no hay nada que
optimizar. El `Group Id` único impide que LBS lo mezcle con otras líneas o lo parta.

### El refuerzo con `Consolidation Class`

`merged/modulo1.vba:530-535`

El `Group Id` por sí solo no bastaba. El comentario lo dice:
*"Walmart EXC28: same keep-together as Super Ofertas 35 (Group Id alone still allowed LBS
splits)"*. Así que además se pone la `Consolidation Class` (columna `AB`) igual al ID del
STR, con la misma prioridad que tiene el caso de Super Ofertas de 35 tarimas.

Es decir, un camión EXC28 lleva tres marcas simultáneas:

| Campo | Valor |
|---|---|
| `Group Id` (col `J`) | `EXC28_<STR id>` |
| `Consolidation Class` (col `AB`) | El `STR id` |
| `Required Equipment` | **Vacío** |

### Por qué `Required Equipment` va en blanco

Este es el detalle contraintuitivo. Un camión de 28 tarimas necesita el equipo `Z4290_WAL`,
que es precisamente el que se define para eso. Pero **no** se declara en el STR.

El comentario de la constante lo documenta (`merged/modulo1.vba:16-18`):

```
' LBS Walmart 28-pallet FULL_PALLET truck (EBL only). STR Required Equipment left blank
' like working 1v2 exports — forcing Z4290_WAL on STR collapsed Walmart fill to ~11%.
```

Forzar el equipo en el STR hizo que el llenado de Walmart cayera a alrededor del 11%. El
equipo se declara solo en `equipmentByLaneByDay`, donde funciona como disponibilidad, y se
deja que LBS lo elija.

`ValidarExportMEXKA` verifica las tres marcas por separado y reporta cada desvío
(`merged/modulo6.vba:3005`, `3012`, `3020`). Ver
[../06-validaciones-mexka.md](../06-validaciones-mexka.md#grupo-15--walmart-exc28).

## Split de restos en `itemConnections`

`LBS_IsWalmartSplitRestosItem` (`merged/modulo1.vba:1038-1053`)

Los mismos 14 SKU, cuando aparecen con sufijo `_WAL` o `_WAL_*`, reciben atributos distintos
en `itemConnections`. El comentario lo especifica (`merged/modulo1.vba:1035-1037`):

```
' WAL MART (*_WAL*): SKUs que deben partir tarimas completas en restos.
' itemConnections: Level/Layer vacios, PalletLevel=2 (como mix NO). Resto de SKUs Walmart
' siguen Condiciones Cadenas (mix SI = 4/3/2). Aplica a *_WAL y *_WAL_* (ej. _WAL_BA).
```

| Tipo de item | `Level` | `Layer` | `PalletLevel` |
|---|---|---|---|
| SKU de excepción con sufijo `_WAL*` | Vacío | Vacío | `2` |
| Resto de los SKU de Walmart | `4` | `3` | `2` (desde `Condiciones Cadenas`) |

Dejar `Level` y `Layer` vacíos le dice a LBS que ese SKU no admite mezcla en capas: la tarima
se parte pero no se comparte.

La función reconoce las dos formas del sufijo: `_WAL` al final del texto
(`merged/modulo1.vba:1045-1046`) o `_WAL_` en medio, como en `_WAL_BA`
(`merged/modulo1.vba:1048-1051`).

## Equipo

Walmart es una de las cadenas que **no emite Full** en `equipmentByLaneByDay`. El comentario
de `LBS_CadenaEnforceCatalogBodyType` lo confirma (`merged/modulo1.vba:1146-1147`):
*"Walmart / solo-sencillo still skip Full emit"*.

El equipo esperado es sencillo: `Z4260` o `Z4290_WAL`. `ValidarExportMEXKA` los cuenta como
mensaje informativo (`merged/modulo6.vba:1852-1854`):

```
WM: revisar N lane(s) con Full en destinos WM
WM: N registros Z4260/Z4290_WAL (sencillo esperado)
```

El primero es una invitación a revisar: si aparecen carriles Full en destinos Walmart, algo
está mal en el catálogo o en `Handling`.

## Cómo validarlo

| Qué revisar | Dónde |
|---|---|
| Los tres campos de EXC28 en el STR | `ValidarExportMEXKA`, grupo `WM EXC28:` del reporte |
| Que no haya `Group Id` compartido entre camiones EXC28 | Hallazgo `WM EXC28: <hoja> Group Id compartido` (`merged/modulo6.vba:3128`) |
| Que los items no-excepción no entren en un grupo EXC28 | Hallazgo `WM EXC28: export STR fila N Group Id=... con Item Id no-exception` (`merged/modulo6.vba:3048`) |
| El conteo de equipo sencillo | Mensajes OK `WM:` del reporte |

Fixtures y scripts disponibles: los archivos `tms_fg14/walmart*/*.tsv` y los scripts
`scripts/validate_walmart*.py` corresponden al lado de **salida**. Del lado de entrada, la
verificación práctica es el reporte de `ValidarExportMEXKA`.

## Problemas conocidos

**El renglón de 28 tarimas no genera `EXC28_`.** Verificar que `Plan!N` sea exactamente 28 y
`Plan!O` exactamente 0. Un armado desactualizado puede dar 27.9 o 28.0001, y la comparación
es de igualdad exacta (`merged/modulo1.vba:470`). Correr `CambioOrigen`, que recalcula `N` y
`O` desde el armado (`merged/modulo3.vba:811-817`).

**El llenado de Walmart en LBS es muy bajo.** Revisar que `Required Equipment` esté en blanco
en el STR. Este es el síntoma documentado del ~11%.

**Items `_WAL` y `_WAL_BA` del mismo SKU se tratan como distintos.** Es lo correcto:
SuperCenter y Bodega Aurrera son cadenas distintas con armados potencialmente distintos.

## En la macro de salida

Walmart es también la cadena con más lógica en la macro de salida: cupo de 28 tarimas, piso
de llenado de 40%, altura máxima de 1.6 m, TI HI, sándwiches y reconstrucción desde
descartes. Ver [../../tms_fg14/cadenas/walmart.md](../../tms_fg14/cadenas/walmart.md).
