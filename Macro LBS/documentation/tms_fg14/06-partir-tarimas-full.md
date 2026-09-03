[Volver a la macro de salida](README.md) · [Índice general](../README.md)

# `PartirTarimasFULL` — dividir los Fulls en cajas

Código: [tms_fg14/modulo5.vba](../../tms_fg14/modulo5.vba), 1 928 líneas. La macro pública es
`PartirTarimasFULL` (`tms_fg14/modulo5.vba:1414`).

Un camión Full es en realidad **dos cajas** enganchadas. LBS lo entrega como un solo embarque
de hasta 40 tarimas; TMS necesita dos unidades de 20. Este módulo hace esa división y le
    10|pone los sufijos `a` y `b` a los folios.

Todo el módulo usa el prefijo `PF_` y **mantiene su propia copia de las constantes y de los
diccionarios de catálogo**, independiente de `modulo2`. Es duplicación deliberada: el módulo
se puede correr por separado sin arrastrar el estado de `SummaryOptimizar`.

## Qué se parte y qué no

`PF_IsFullGroupRow` (`tms_fg14/modulo5.vba:365-377`) decide qué filas entran:

```vba
    20|yVal = CStr(ws.Cells(r, "Y").Value)
If InStr(1, yVal, "Full", vbTextCompare) = 0 Then Exit Function
bVal = UCase$(Trim$(CStr(ws.Cells(r, "B").Value)))
mVal = UCase$(Trim$(CStr(ws.Cells(r, "M").Value)))
' Catalog Mode Mix is SOT: if Y is Full, split (Soriana/City Club included when catalog F).
PF_IsFullGroupRow = (bVal = "AUTOSERVICIO" Or mVal = "OXXO")
```

Dos condiciones, ambas necesarias:

    30|1. La columna `Y` (`Tipo de Unidad`) contiene `Full`.
2. La columna `B` (`Clasificación Cadena`) es `AUTOSERVICIO`, **o** la cadena es exactamente
   `OXXO`.

La segunda condición explica una omisión que sorprende: **los mayoristas no se parten.** Su
clasificación en `B` es `MAYORISTA`, así que nunca entran. Es coherente con el diseño de
cupos, donde `LBS_IsMayoristaSingleFullTruck` usa el cupo de embarque completo y nunca el de
caja (ver [04-motor-armado-cargas.md](04-motor-armado-cargas.md)).

El comentario declara al catálogo como fuente de verdad: si `Y` dice `Full`, se parte, y eso
incluye a Soriana y City Club cuando el catálogo los declara como `F`.
    40|
## Los cupos: `maxShipment` y `maxTarimas`

Dos números por grupo (`tms_fg14/modulo5.vba:667-675`):

```vba
Private Function PF_MaxTarimasShipment(ByVal ws As Worksheet, ByVal r As Long) As Long
    PF_MaxTarimasShipment = PF_CatalogFullShipmentCap(ws, r)
    If PF_MaxTarimasShipment < 2 Then PF_MaxTarimasShipment = 40
End Function

    50|Private Function PF_MaxTarimasForRow(ByVal ws As Worksheet, ByVal r As Long) As Long
    PF_MaxTarimasForRow = PF_MaxTarimasShipment(ws, r) \ 2
    If PF_MaxTarimasForRow < 1 Then PF_MaxTarimasForRow = 20
End Function
```

| Nombre | Qué es | Respaldo |
|---|---|---|
| `maxShipment` | Tarimas del embarque completo, del `Catalogo Mode Mix` | `40` |
| `maxTarimas` | Tarimas por caja, la mitad exacta con división entera | `20` |

    60|La división es entera (`\ 2`), así que un embarque de 35 da cajas de 17. La tarima impar
sobrante se acomoda en la última caja abierta.

Las constantes propias del módulo (`tms_fg14/modulo5.vba:1-16`):

| Constante | Valor | Comentario original |
|---|---|---|
| `PF_MAX_PESO_KG` | `29000` | (sin comentario) |
| `PF_MAX_GROUP_EXTRA_ROWS` | `50` | Presupuesto de filas extra por grupo |
| `PF_MAX_OVERFLOW_SEQ` | `20` | Tope de embarques `R2`, `R3`, … por grupo |
| `PF_CAP35_SHIPMENT` | `35` | `' SKU-level Full shipment cap (sheet "Cadenas 35 Tarimas": allowlisted cadena + SKU).` |
    70|| `PF_CAP35_SHEET` | `"Cadenas 35 Tarimas"` | |
| `PF_MAYORISTA_FULL_SHIPMENT` | `36` | |

## El bloque: `PF_BlockKey`

Antes de partir hay que saber qué filas forman un mismo embarque. `PF_BlockKey`
(`tms_fg14/modulo5.vba:18-39`) normaliza el folio quitándole todo lo que le agregó una corrida
anterior:

1. **Quita los sufijos de letra minúscula** del final, en ciclo. `02092026P-1020ab` queda en
    80|   `02092026P-1020`.
2. **Quita el sufijo de desbordamiento** `R<n>`: busca la última `R` y, si lo que sigue es
   numérico, corta ahí.

Eso hace la macro **idempotente**: se puede volver a ejecutar sobre una hoja ya partida y
reagrupa los folios correctamente antes de repartir.

`PF_BlockKeyForRow` (`tms_fg14/modulo5.vba:56-65`) agrega el único caso especial por cadena:

```
' LA COMER: R y C no comparten bloque PartirFulles (col AE = Plan!W).
    90|```

Para La Comer se concatena la marca `R` o `C` de la columna `AE`, de modo que los pedidos de
reposición y los de confirmación nunca comparten bloque.

## El conteo de tarimas físicas

`PF_PhysicalTarimas` (`tms_fg14/modulo5.vba:101-109`) toma `T` si es positivo, y si no delega
en `PF_TarimasTW`. Este último tiene la regla más sutil del módulo
(`tms_fg14/modulo5.vba:75-77`):

   100|```
' T = tarimas completas; W = indicador numero de tarima (0/1).
' T=0 and W=1 => 1 tarima (fila restos OXXO).
' T>0, W=1, U>0 y T>=4 => T + 1 (charola sobre ancla multi-tarima; anclas chicas ya cuentan en T).
```

El umbral `T >= 4` es la parte no obvia. Una fila con tarimas completas y restos ocupa una
tarima adicional solo si tiene cuatro o más tarimas completas. Con menos, la charola ya cabe
en el espacio contado por `T`.

   110|## El flujo por grupo

Para cada grupo de folio (`tms_fg14/modulo5.vba:1507-1586`):

```mermaid
flowchart TB
  A["Reordenar filas del folio<br/>LBS_ReordenarYCompletar"] --> B["Limpiar AV de las filas Programado"]
  B --> C{"Es fila Full?<br/>PF_IsFullGroupRow"}
  C -->|no| N["Siguiente fila"]
  C -->|si| D["Sumar tarimas del bloque"]
  D --> E{"total > maxShipment?"}
   120|  E -->|si| F["Recortar el exceso a No planeado<br/>PF_TrimTwoTruckGroupToShipmentCap"]
  E -->|no| G
  F --> G{"Ya esta partido?<br/>PF_BlockFullySplit"}
  G -->|si y dentro de cupo| N
  G -->|no| H["Resetear AD al folio base"]
  H --> I["Repartir en cajas a, b, c...<br/>sufijo por caja"]
  I --> J{"Se agoto el embarque?"}
  J -->|si| K["Abrir shipment R2, R3...<br/>PF_SpinNewShipment"]
  J -->|no| L["Recalcular cartonaje y totales"]
  K --> L
   130|```

### Paso 0: reordenar

Lo primero que hace la macro es reordenar la hoja
(`tms_fg14/modulo5.vba:1476-1481`), con esta justificación:

```
' Filas de un mismo folio contiguas: los merges de SummaryOptimizar (metro/gate) reasignan
' AD sin mover filas, y el escaneo de grupos Full de abajo corta en la primera fila ajena.
```

   140|El escaneo de grupos avanza secuencialmente y **se detiene en la primera fila que no
pertenece al bloque** (`tms_fg14/modulo5.vba:1524-1533`). Si las filas de un folio quedaron
dispersas tras los merges de `SummaryOptimizar`, el grupo se cortaría a la mitad y se
repartiría mal.

### Paso 1: limpiar `AV`

```vba
' Limpiar flags de revision previos en filas Programado (evita acumular en AV al re-ejecutar).
```

   150|(`tms_fg14/modulo5.vba:1483-1488`.) Es la única macro que limpia `AV` sistemáticamente. Se
limpia solo en filas `Programado`: las banderas de las filas `No planeado` se conservan.

### Paso 2: recortar el exceso

Si el bloque trae más tarimas que el cupo de embarque, el exceso se manda a `No planeado`.
El comentario es tajante (`tms_fg14/modulo5.vba:1538`):

```
' Full: trim excess above catalog shipment cap to No planeado — no R2.
```

   160|Las cadenas Full **nunca** desbordan a un embarque `R2`. El exceso se descarta. El mecanismo
`R2` existe pero está reservado para otros casos.

### Paso 3: detectar si ya está partido

Dos comprobaciones distintas:

- `PF_TryNormalizePreSplitTwoTruckGroup` (`tms_fg14/modulo5.vba:1546`) — intenta normalizar un
  grupo que ya venía partido.
- `PF_BlockFullySplit` (`tms_fg14/modulo5.vba:152`) — verifica que **todas** las filas
   170|  `Programado` tengan sufijo de letra y que ningún sufijo exceda `maxTarimas`.

El primero tiene un rechazo explícito (`tms_fg14/modulo5.vba:1547-1548`):

```
' Reject "already split" when still over truck/shipment cap (e.g. stale 20+18).
```

Un grupo con sufijos `a` de 20 y `b` de 18 parece partido, pero suma 38 y puede seguir sobre
cupo. En ese caso se rehace desde cero.

   180|### Paso 4: repartir

Se resetean todos los `AD` del grupo al folio base
(`tms_fg14/modulo5.vba:1564-1567`) y se reparte tarima por tarima. El sufijo arranca en `"a"`
y avanza con `PF_AdvanceTruckSuffix` (`tms_fg14/modulo5.vba:681-684`), que simplemente
incrementa el código ASCII:

```vba
suffix = Chr$(Asc(suffix) + 1)
tarimasRemaining = maxTarimas
```

   190|Así que un grupo muy grande puede llegar a `c`, `d`, `e`… No hay tope de letra explícito,
aunque en la práctica el recorte al cupo de embarque limita a dos cajas.

`PF_PlaceAmount` (`tms_fg14/modulo5.vba:1168-1189`) decide cuántas tarimas van en la caja
actual: el mínimo entre lo que falta colocar, lo que queda en la caja y lo que queda en el
embarque.

`PF_UseTwoTruckSplit` (`tms_fg14/modulo5.vba:1164-1166`) determina el modo:

```vba
PF_UseTwoTruckSplit = (totalTarimas > maxTarimas And totalTarimas <= maxShipment)
   200|```

Más de una caja pero no más de un embarque. Y después del recorte se refuerza
(`tms_fg14/modulo5.vba:1580-1583`):

```
' After trim, Full with boxCap < total <= shipCap must use two-truck path (never R2).
```

### El presupuesto de filas

   210|Antes de repartir se calcula un presupuesto (`tms_fg14/modulo5.vba:1570`):

```vba
groupRowBudget = (groupEnd - groupStartRow + 1) + totalTarimas + PF_MAX_GROUP_EXTRA_ROWS
```

Filas actuales del grupo, más una por tarima, más 50 de margen. Es un tope de seguridad: el
reparto inserta filas cuando tiene que dividir una fila en dos, y sin límite un ciclo mal
condicionado podría insertar sin fin.

## El caso especial de OXXO: separar los restos
   220|
`PF_SplitOxxoRestosRow` (`tms_fg14/modulo5.vba:1325-1351`) se ejecuta en cada iteración del
reparto. Solo aplica cuando la cadena es exactamente `OXXO` y la fila trae `T > 0` **y**
`U > 0` al mismo tiempo.

Lo que hace es partir la fila en dos:

```vba
' Parent keeps full tarimas (restos peeled off); recalc S=T*V+U.
ws.Cells(r, "S").Value = origS - uV
ws.Cells(r, "U").Value = 0
   230|' Restos row: T=0 so recalc sets S=U=restos. Keep restos in U, not just S.
ws.Cells(r + 1, "S").Value = uV
ws.Cells(r + 1, "T").Value = 0
ws.Cells(r + 1, "U").Value = uV
ws.Cells(r + 1, "W").Value = 1
```

- La fila original conserva las tarimas completas, con `U = 0`.
- La fila nueva lleva solo los restos, con `T = 0` y `W = 1`.

La razón: OXXO no permite partir una tarima entre dos cajas. Separar los restos deja que las
   240|tarimas completas y el resto se asignen a cajas distintas sin partir nada. Ver
[cadenas/oxxo-neto.md](cadenas/oxxo-neto.md).

Es la única inserción de filas del módulo que ocurre dentro del ciclo, y ajusta tanto
`lastRow` como `groupEnd` para que el recorrido siga siendo válido.

## El recálculo del cartonaje

Al partir un grupo, el cartonaje de cada fila tiene que cuadrar con las tarimas que le
quedaron. La fórmula es la inversa de la que usa `SummaryOK`:

   250|```
S = T * V + U
```

donde `V` es el armado. El módulo guarda `groupOrigS` (la suma original del grupo,
`tms_fg14/modulo5.vba:1562`) y lo compara al final, para que la suma de las partes iguale al
total original.

## El desbordamiento: los embarques `R2`

Cuando un grupo no cabe ni en el embarque, `PF_SpinNewShipment`
   260|(`tms_fg14/modulo5.vba:686`) abre uno nuevo agregando `R2`, `R3`, … al folio base, hasta
`PF_MAX_OVERFLOW_SEQ = 20`.

`PF_IsOverflowShipment` (`tms_fg14/modulo5.vba:677-679`) detecta esos folios buscando una `R`,
y `PF_BlockKey` los normaliza al reagrupar.

Como se dijo arriba, las cadenas Full ya no usan este camino: el exceso se recorta antes. El
mecanismo queda para casos que el recorte no cubre.

## El registro en `STP Failures`

   270|`PF_WriteOverflowLog` (`tms_fg14/modulo5.vba:1353-1403`) escribe **tres bloques** al final de
la hoja `STP Failures`, cada uno con su propio encabezado. La hoja se usa como bitácora, no
solo como entrada de LBS.

### Bloque 1: grupos sobre el límite

Encabezado: `Shipment origen` | `Shipment nuevo` | `Tarimas` | `Mensaje`

```
REVISION MANUAL: <N> tarimas superan limite <cupo>
```
   280|
### Bloque 2: desbordamientos

Mismo encabezado. Mensaje fijo:

```
Overflow >36: tarimas movidas a nuevo shipment
```

El `36` está escrito literalmente en el texto (`tms_fg14/modulo5.vba:1384`), aunque el cupo
real dependa del catálogo. Es un mensaje heredado.
   290|
### Bloque 3: camiones sobre peso

Encabezado propio: `Camion AD` | `Peso kg` | `Peso ton` | `Mensaje`

```
REVISION MANUAL: peso supera 29 ton
```

Con el peso en kilos y en toneladas redondeado a un decimal.

   300|**Estos tres bloques se agregan al final de la hoja sin borrar lo anterior.** Correr
`PartirTarimasFULL` dos veces deja dos juegos de bitácora. Solo `ProcesoMacro` limpia
`STP Failures`.

## El indicador de progreso

`PF_SetProgress` (`tms_fg14/modulo5.vba:1405-1412`) escribe en `Plan!X1` con la hora, y llama
a `DoEvents` una de cada cuatro veces:

```vba
If (PF_ProgressTick And 3) = 0 Then DoEvents
   310|```

Los mensajes tienen la forma:

```
PartirFulles: inicio 0/1234 | 09:15:22
PartirFulles: fulls 500/1234 | 09:15:40
PartirFulles: grupo 12 fila 340 | 09:15:41
```

Si la macro se cae, el último mensaje dice en qué grupo y en qué fila fue.
   320|
## Qué revisar después

| Dónde | Qué buscar |
|---|---|
| `AD` | Que los folios Full traigan sufijo `a` y `b`. Un Full sin sufijo no se partió |
| `AV` | Banderas de tarimas y de peso en las filas `Programado` |
| `STP Failures` | Los tres bloques de bitácora al final de la hoja |
| `AG` | Filas nuevas en `No planeado` por recorte al cupo de embarque |
| `S` | Que la suma del cartonaje por bloque siga igual a la de antes de partir |

   330|El fixture [tms_fg14/split bug.tsv](../../tms_fg14/split%20bug.tsv) y
[tms_fg14/surtido before split.tsv](../../tms_fg14/surtido%20before%20split.tsv) son casos
reales de esta fase, útiles para comparar el antes y el después.

## El manejo de error

La macro tiene una etiqueta `PartirTarimasFULL_Fail` que restaura `ScreenUpdating`,
`Calculation` e `EnableEvents` antes de propagar el error. A diferencia de `SummaryOK`, no usa
la bandera `SK_MacroBusy`, así que **sí es posible re-entrarla con un doble clic.** Conviene
esperar a que el progreso de `Plan!X1` deje de moverse antes de tocar otro botón.
