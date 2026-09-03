[Volver a la macro de salida](README.md) · [Índice general](../README.md)

# La hoja `Pedidos Surtidos`

Es la hoja central de la macro de salida. Todo lo que hacen `SummaryOK`,
`SummaryOptimizar`, `PartirTarimasFULL`, `SummaryFallo` y `CompararCartonajes` es leerla y
reescribirla. Cuando el resultado sale mal, aquí está la evidencia.

**Una fila representa una combinación de camión + pedido + SKU + tarima.** No es una fila por
    10|camión ni una fila por pedido. Un camión (identificado por el folio de la columna `AD`)
suele ocupar entre 15 y 40 filas.

## `SK_PS_LAST_COL` = 46, y por qué importa

El comentario del código es explícito (`tms_fg14/modulo2.vba:111-112`):

```
' Pedidos Surtidos A:AT = 46 cols; keep clear/write under ~60k cells per COM op.
Private Const SK_PS_LAST_COL As Long = 46
```
    20|
La columna 46 es `AT`. Eso significa que el bloque que se escribe y se limpia en masa es
**`A:AT`**, no `A:AV`:

- El volcado de la PCA arma un arreglo de `46` columnas y lo descarga de un golpe
  (`tms_fg14/modulo2.vba:713`, `777`).
- La limpieza inicial y el borrado de la cola posterior a la consolidación también van sobre
  `A:AT`, en trozos de 1 000 filas (`SK_ClearPSRangeChunkedFrom`,
  `tms_fg14/modulo2.vba:157-177`).

    30|Las dos últimas columnas quedan **fuera** de ese bloque y se escriben celda por celda,
en momentos distintos del pipeline:

| Col | Contenido | Quién la escribe |
|---|---|---|
| `AU` | Peso total del camión, incluyendo la tara de las tarimas | `tms_fg14/modulo2.vba:1651`, `1657` |
| `AV` | Las banderas de `REVISION MANUAL` | `SK_AppendReviewFlag`, `tms_fg14/modulo2.vba:433-444` |

La consecuencia práctica: **`AU` y `AV` no se limpian con el resto de la hoja.** Si una
corrida deja banderas en `AV` y la siguiente no vuelve a tocar esas filas, las banderas
quedan. Hay lógica específica para quitarlas selectivamente
    40|(`SK_StripReviewFlagMatching`, `tms_fg14/modulo2.vba:446-467`), pero conviene revisar `AV`
con la expectativa de que puede traer residuo.

## Mapa completo de columnas

Los encabezados vienen del fixture [tms_fg14/soriana/sample.tsv](../../tms_fg14/soriana/sample.tsv).
La columna "Se llena en" indica en qué fase aparece el valor.

### El volcado inicial de la PCA — 8 columnas

Estas son las únicas columnas que existen justo después de leer la
    50|`Pallet Container Association` (`tms_fg14/modulo2.vba:766-774`):

| Col | Encabezado | Valor | Origen en PCA |
|---|---|---|---|
| `D` | (sin encabezado) | `Unit Id` | PCA `G` |
| `E` | (sin encabezado) | `Unit Type`: `SINGLE`, `PALLET` o `SANDWICH` | PCA `H` |
| `P` | `Pedido` | El pedido, cortado en el primer `_` | PCA `Z` |
| `S` | `Cartonaje Total` | Cajas | PCA `S` (`Order Raw Quantity`) |
| `W` | `Numero de Tarima` | `1` o `0`, según la regla de conteo de unidad | Calculado |
| `X` | `SKU` | El `ItemId` con sufijo de cadena | PCA `Q` |
| `AD` | `Confirmcion` | El folio: `DDMMAAAA` + `P-` + `Shipment Id` | PCA `C` + fecha de hoy |
    60|| `AT` | `Peso tarima` | Kilogramos | PCA `T` (`Weight`) |

#### La regla de `W` en el volcado

`W` no es un conteo, es un **marcador de una sola tarima**
(`tms_fg14/modulo2.vba:757-764`):

```vba
If unitID = lastUnitID Then
    wAssignBuf = 0
ElseIf UCase$(pallet) = "SINGLE" Then
    70|    If unitCount > 1 Then wAssignBuf = 1 Else wAssignBuf = 0
Else
    wAssignBuf = 1
End If
```

En palabras: la primera fila de cada unidad física se marca con `1` y las demás con `0`, para
que la misma tarima no se cuente dos veces. Las unidades `SINGLE` con un solo renglón son la
excepción y reciben `0`, porque su tarima se cuenta desde `T` (tarimas completas) y no desde
`W`.

    80|Esto depende del ordenamiento por `Shipment Id` + `Unit Id` que aplicó `ProcesoMacro`
(ver [02-entradas-lbs.md](02-entradas-lbs.md)). La comparación `unitID = lastUnitID` es
contra la fila **inmediatamente anterior**, no contra un diccionario, así que una hoja
desordenada rompe el conteo.

Para eso existe el primer recorrido de la PCA, que solo lee la columna `G` y cuenta cuántas
filas tiene cada unidad (`tms_fg14/modulo2.vba:676-709`).

### Lo que copia el Plan — la identidad comercial

Se llenan durante la consolidación de `SummaryOK`, buscando la fila del Plan por
    90|`pedido|material` (`tms_fg14/modulo2.vba:1243-1259`):

| Col | Encabezado | Origen | Cita |
|---|---|---|---|
| `A` | `Fecha Fill Rate` | Plan | `tms_fg14/modulo2.vba:1243` |
| `B` | `Clasificación Cadena` | Plan (`Autoservicio` / `Mayorista`) | `tms_fg14/modulo2.vba:1245` |
| `G` | `Fecha asignado pedido` | Plan | `tms_fg14/modulo2.vba:1246` |
| `K` | `Fecha OC` | Plan | `tms_fg14/modulo2.vba:1248` |
| `L` | **`Origen`** | Plan, o el `Lane Id` de `Shipments` como respaldo | `tms_fg14/modulo2.vba:1250` |
| `M` | **`Cadena`** | Plan | `tms_fg14/modulo2.vba:1251` |
| `N` | `CEDIS` | Plan | `tms_fg14/modulo2.vba:1252` |
   100|| `O` | `Destinatario` | Plan | `tms_fg14/modulo2.vba:1253` |
| `Q` | `Cotización SAP` | Plan | `tms_fg14/modulo2.vba:1254` |
| `R` | **`Vigencia`** | Plan; si falta, `STP-Equipment Assoc!L` | `tms_fg14/modulo2.vba:1240`, `1232` |
| `AB` | `UPC` | Plan | `tms_fg14/modulo2.vba:1255` |
| `AC` | `Presentacion` | Plan | `tms_fg14/modulo2.vba:1256` |
| `AE` | `Comentario Comercial` | Plan | `tms_fg14/modulo2.vba:1257` |
| `AF` | `Comentario Planeación` | Plan | `tms_fg14/modulo2.vba:1258` |
| `AH` | `Tipo de tarima` | Plan (`Tarima Chep` / `Plastica` / `Madera`) | `tms_fg14/modulo2.vba:1259` |

`M` es la columna más consultada de toda la macro: es la que decide qué reglas de cadena se
   110|aplican. Prácticamente cada función del motor empieza con una consulta del tipo
`LBS_IsWalmartChain(CStr(ws.Cells(r, "M").Value))`.

`L` merece atención aparte. Cuando el Plan no la trae, o trae un marcador de inventario
(`NoInventario`, `FaltaInventario`, `FABRICA`), se deriva del `Lane Id` de `Shipments`
(`LBS_IsInventoryPlaceholderOrig`, `tms_fg14/modulo2.vba:4702-4707`). Un origen mixto dentro
de un mismo folio es motivo de split en varias cadenas.

### Lo que calcula el motor — el armado

| Col | Encabezado | Significado | Cita |
   120||---|---|---|---|
| `H` | **`Estatus`** | `Programado` o `No planeado`. El filtro maestro | `tms_fg14/modulo2.vba:1138` |
| `I` | `Contador` | Posición de la fila dentro del camión, empezando en 1 | `tms_fg14/modulo2.vba:22138` |
| `J` | `Semana` | `Format$(Now, "ww")` — la semana ISO de la corrida | `tms_fg14/modulo2.vba:1136` |
| `T` | **`Tarimas completas`** | Tarimas llenas de este SKU | `tms_fg14/modulo2.vba:1298`, `1337` |
| `U` | **`Restos de charolas`** | Cajas que no llenan tarima | `tms_fg14/modulo2.vba:1301-1307` |
| `V` | `Armado` | Cajas por tarima, del catálogo | `tms_fg14/modulo2.vba:1265` |
| `W` | **`Numero de Tarima`** | Tarima parcial: cuenta como espacio ocupado en el camión | `tms_fg14/modulo2.vba:1313-1330` |
| `X` | `SKU` | **Se sobrescribe dos veces.** Ver abajo | `tms_fg14/modulo2.vba:1401`, `22121` |
| `Y` | `Tipo de Unidad` | Nombre legible del equipo | `tms_fg14/modulo2.vba:1387` |
   130|| `Z` | `Etiquetas y total de tarimas por unidad` | `SUM(T) + SUM(W)` del folio, en su última fila `Programado` | `tms_fg14/modulo2.vba:2145-2147` |
| `AA` | **`Material`** | El SKU **sin** el sufijo de cadena | `tms_fg14/modulo2.vba:1158` |
| `AD` | **`Confirmcion`** | El folio del camión | `tms_fg14/modulo2.vba:1160` |
| `AG` | **`Motivo de incumplimiento`** | El motivo, o un marcador interno | `tms_fg14/modulo2.vba:1289-1291` |
| `AI` | `Craft` | Los `Unit Id` de la unidad sándwich, separados por coma | `tms_fg14/modulo2.vba:1147` |
| `AK` | `Fecha programa` | `STP-Equipment Assoc!J` | `tms_fg14/modulo2.vba:1395` |
| `AL` | `**` | Peso de los restos de la unidad (solo Walmart) | `tms_fg14/modulo2.vba:6484`, `6511` |
| `AN` | `canary` | Suma de restos de la unidad; también marca inicio de sándwich | `tms_fg14/modulo2.vba:6470`, `6497` |
| `AO` | **`Altura SKU`** | Altura en metros de este SKU en la tarima | `tms_fg14/modulo2.vba:8828`, `9509` |
| `AP` | **`Altura Total`** | Suma de `AO` por unidad `AD` + `AI` | `tms_fg14/modulo2.vba:8974` |
   140|| `AR` | **`Cumplimiento`** | El fill rate del embarque, de `Shipments!L` | `tms_fg14/modulo2.vba:4923` |
| `AT` | `Peso tarima` | Peso de esta fila en kg | `tms_fg14/modulo2.vba:1162` |
| `AU` | `Peso total SKU` | **Peso total del camión**, tarimas incluidas | `tms_fg14/modulo2.vba:1651` |
| `AV` | `Ruta` | **Las banderas de `REVISION MANUAL`** | `tms_fg14/modulo2.vba:438`, `440` |

### Columnas que la macro no escribe

`C` (`ORIGEN 1`), `F` (`INICIATIVAS VA DAY`), `AJ` (`Hectolitros`), `AM` (`ENTREGA`),
`AQ` (`Dias planeacion/ Transporte`) y `AS` (`Peso Material`) aparecen en el encabezado del
fixture pero ninguna macro les escribe. Son columnas de la plantilla que se llenan con
fórmulas de la hoja o a mano.

   150|## Tres advertencias sobre los encabezados

Hay tres puntos donde el encabezado de la plantilla y lo que el código realmente guarda no
coinciden. Vale la pena tenerlos presentes porque desorientan al leer la hoja.

### 1. `AT` / `AU` no son los pesos que dicen ser

El encabezado dice `Peso tarima` en `AT` y `Peso total SKU` en `AU`. El código guarda:

- `AT` = el peso de **esa fila** (ese SKU, en ese camión, en esa tarima), tomado del `Weight`
    160|  de la PCA (`tms_fg14/modulo2.vba:1162`).
- `AU` = el peso **de todo el camión**, calculado como la suma de los `AT` del folio más la
  tara de las tarimas (`tms_fg14/modulo2.vba:1648-1651`):

```vba
pesoCamion = pesoCamion + SK_SafeDbl(wsPS.Cells(r, "AT").Value)
...
wsPS.Cells(i, "AU").Value = pesoCamion + (LBS_PesoTarimaKg() * CDbl(zTotal))
```

`AU` solo tiene valor en una fila por folio, y es contra ese valor que se compara el techo de
peso. Cuando lo excede se agrega la bandera correspondiente en `AV`
   170|(`tms_fg14/modulo2.vba:1661-1663`).

### 2. `AV` dice `Ruta` pero contiene las alertas

Todas las banderas de `REVISION MANUAL` van a `AV`. `SK_AppendReviewFlag`
(`tms_fg14/modulo2.vba:433-444`) las **concatena** con `; ` en lugar de sobrescribir, así que
una fila puede acumular varias:

```
REVISION MANUAL: tarimas <26 (24 tarimas); REVISION MANUAL: peso >29 ton (30,150 kg)
```

   180|El catálogo completo de banderas está en [10-diagnostico-y-errores.md](10-diagnostico-y-errores.md).

### 3. `X` se reescribe como marca de camión

`X` empieza siendo el `ItemId` de la PCA, con el sufijo de cadena incluido. Durante la
consolidación se sobrescribe con `AA`, el material sin sufijo (`tms_fg14/modulo2.vba:1401`):

```vba
' Reemplazo campo SKU por Material
wsPS.Cells(filaDestino, "X").Value = wsPS.Cells(filaDestino, "AA").Value
```

   190|Y al final del reordenamiento se sobrescribe **otra vez**, ahora como marca de inicio de
camión (`tms_fg14/modulo2.vba:22120-22137`):

```vba
' Marca de camion (col X = SKU): solo la primera fila del folio = 1.
ws.Cells(i, "X").Value = 1
...
' El resto de filas del folio en blanco, para que el camion no "reinicie" en 1
' al unir dos cadenas del mismo grupo (Walmart SC + BA, City Club + Soriana).
ws.Cells(i, "X").Value = ""
```

   200|Después de una corrida completa, `X` contiene un `1` en la primera fila de cada camión y
nada en el resto. **Para leer el SKU hay que usar `AA`, no `X`.** El segundo comentario
explica el motivo del cambio: cuando dos cadenas del mismo grupo de consolidación comparten
camión, dejar la marca en cada cadena hacía que el conteo de camiones se reiniciara a mitad
del folio.

## Los cuatro campos que hay que entender

### `H` — `Programado` contra `No planeado`

`H` es el filtro maestro. Casi todas las funciones del motor arrancan con:
   210|
```vba
If Trim$(CStr(ws.Cells(i, "H").Value)) <> "Programado" Then GoTo NextRow
```

- `Programado` (`tms_fg14/modulo2.vba:1138`) — la fila va en un camión que se va a despachar.
- `No planeado` — la fila quedó fuera. El motivo está en `AG`.

Las filas `No planeado` no son basura: son la reserva desde la que las cadenas rellenan
camiones que no llegan a su piso de llenado, y es lo que hace `ConsolidarNoPlaneados`
(`tms_fg14/modulo2.vba:5075`).

   220|Una condición que el motor cuida explícitamente: las filas vacías no deben quedar entre las
`Programado` y las `No planeado`. `LBS_CompactBlankPSRows` las elimina, con este comentario
(`tms_fg14/modulo2.vba:2142`):

```
' Drop cleared ghost rows so they cannot sit between Programado and No planeado.
```

### `AD` — el folio del camión

`AD` es lo que identifica un camión. Su forma base es
   230|
```
DDMMAAAA + "P-" + <Shipment Id>
```

donde `DDMMAAAA` es la fecha del día en que se corrió la macro
(`confirmDate = Format(Date, "DDMMYYYY")`, `tms_fg14/modulo2.vba:671`) y el `Shipment Id`
viene de la PCA. Ejemplo: `02092026P-1020`.

`PartirTarimasFULL` le agrega sufijos `a` y `b` al partir un Full en dos cajas. Dos funciones
manejan eso:

   240|- `SK_ShipmentIdFromAD` (`tms_fg14/modulo2.vba:406-421`) extrae el `Shipment Id` limpio.
- `SK_HasTruckSuffix` (`tms_fg14/modulo2.vba:423-431`) detecta si el folio ya trae sufijo.
- `LBS_BaseFolioAD` (`tms_fg14/modulo2.vba:4752`) devuelve el folio sin sufijo.

El detalle está en [04-motor-armado-cargas.md](04-motor-armado-cargas.md) y
[06-partir-tarimas-full.md](06-partir-tarimas-full.md).

### `T`, `U` y `W` — la aritmética de las tarimas

Las tres describen cuánto espacio ocupa la fila en el piso del camión, y hay que leerlas
juntas:
   250|
| Columna | Qué cuenta |
|---|---|
| `T` | Tarimas **completas** de este SKU |
| `U` | Cajas de **resto**, las que no llenan una tarima |
| `W` | `1` si esta fila ocupa además una tarima **parcial** compartida |

La relación con el cartonaje es (`tms_fg14/modulo2.vba:1294-1307`):

```
T = parte entera de S / V          (S = cartonaje, V = armado)
   260|U = S - (T * V)
```

Y el espacio total que ocupa un camión es `SUM(T) + SUM(W)`, que es exactamente lo que se
escribe en `Z` (`tms_fg14/modulo2.vba:2145-2147`). Esa suma es la que se compara contra el
cupo de la cadena.

La distinción entre `U` y `W` es la que hace posible el apilado: varias filas con `U > 0`
pueden compartir una sola tarima física, y en ese caso solo una de ellas lleva `W = 1`. La
lógica que lo decide vive en `SF_ApplySandwichTAdjust` y `SF_RecalcWFromTarimas`
(`tms_fg14/modulo3.vba:331`, `383`).

   270|Una regla concreta que conviene conocer: una fila califica como **ancla de sándwich** solo si
tiene restos. El comentario del código lo dice así
(`tms_fg14/modulo2.vba:6422-6423`):

```
' Ancla sandwich solo si la fila trae restos (U>0). T>0 con U>0 es ancla mixta valida.
' T>0 con U=0 es tarima completa sin restos: no es ancla.
```

### `AG` — el motivo, y los marcadores internos

   280|`AG` mezcla dos cosas. Cuando la fila es `No planeado`, contiene el motivo en texto legible.
Cuando la fila es `Programado`, puede contener un marcador interno del motor:

| Valor | Significado |
|---|---|
| `LBS_SANDWICH_ANCHOR` | Esta fila es el ancla de una unidad sándwich |
| `LBS_SANDWICH` | Esta fila es miembro de una unidad sándwich, pero no el ancla |

Se escriben en `tms_fg14/modulo2.vba:1289-1291` y se consultan en
`LBS_RowIsSandwichStackMember` (`tms_fg14/modulo2.vba:6412-6420`).

   290|Hay lógica dedicada a limpiar valores obsoletos: `LBS_ClearStaleDescartadoAGOnProgramado`
(`tms_fg14/modulo2.vba:4684`) borra motivos de descarte que quedaron en filas que después se
volvieron `Programado`, y `LBS_ClearMountAGUnlessNoReportado`
(`tms_fg14/modulo2.vba:4675`) hace lo mismo preservando la marca de no reportado por LBS.

El catálogo de motivos está en [05-eficiencia-y-descartes.md](05-eficiencia-y-descartes.md).

## El orden final de las filas

Después de todo el procesamiento, la hoja se reordena con una llave sintética
(`tms_fg14/modulo2.vba:22079-22081`):
   300|
```
"K" | folio (5 dígitos) | prioridad de confirmación | AI (6 dígitos)
    | prioridad ANCHOR/SANDWICH/otro | 9999 - T | número de fila
```

Las filas `No planeado` reciben una llave que empieza con `"Z"`, lo que las manda al final.

El prefijo `"K"` y el `NumberFormat = "@"` no son adorno. El comentario explica el problema
que resolvieron (`tms_fg14/modulo2.vba:22076-22078`):

   310|```
' Prefijo "K" + NumberFormat @: si la clave es solo digitos Excel la castea a
' Double, pierde AI en los digitos bajos y deja Contador partido (P-1215).
```

Sin la letra, Excel convertía la llave a `Double`, perdía los dígitos menos significativos
—precisamente los del `Unit Id`— y el contador de un folio quedaba partido en dos bloques.

El resultado es que la hoja queda agrupada por camión, y dentro de cada camión las unidades
sándwich quedan contiguas con su ancla primero.

   320|## Los límites de tamaño

| Constante | Valor | Comentario original | Cita |
|---|---|---|---|
| `SK_PS_LAST_COL` | `46` | `' Pedidos Surtidos A:AT = 46 cols; keep clear/write under ~60k cells per COM op.` | `tms_fg14/modulo2.vba:111-112` |
| `SK_PS_CLEAR_CHUNK` | `1000` | Filas por bloque de limpieza | `tms_fg14/modulo2.vba:113` |
| `LBS_WALMART_PCA_CHUNK` | `1000` | Filas por bloque de lectura de la PCA | `tms_fg14/modulo2.vba:109` |
| `LBS_WALMART_PCA_MAX_ROWS` | `250000` | Tope de filas de la PCA | `tms_fg14/modulo2.vba:110` |

Los comentarios de la zona de limpieza repiten la misma advertencia tres veces, lo que da
   330|una idea de cuántas veces se tropezó con ella (`tms_fg14/modulo2.vba:157`, `161-162`):

```
' Clear A2:AT in row chunks (full-range Clear can exceed Excel COM cell limit).
' Clear A:AT from startRow..lastRow in chunks. Used after consolidar to wipe the
' leftover PCA dump (often 50k+ rows) — one-shot Rows().Clear hard-closes Excel.
```

El último renglón se detecta por la columna `AD`, con `A` como respaldo
(`SK_PS_ClearLastRow`, `tms_fg14/modulo2.vba:148-155`). Se usa `AD` porque el folio siempre
está presente en una fila válida, mientras que `A` puede venir vacía si el Plan no trajo
   340|fecha de fill rate.
