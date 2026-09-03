[Volver a la macro de salida](README.md) · [Índice general](../README.md)

# Las entradas de LBS

`ProcesoMacro` ([tms_fg14/modulo1.vba:121](../../tms_fg14/modulo1.vba)) importa cinco
archivos a cinco hojas del libro. No transforma nada: copia valores, ordena y en un caso
filtra. Todo el trabajo de interpretación lo hace `SummaryOK` después.

Este documento describe cada hoja: qué rango se importa, cómo queda ordenada y qué columnas
    10|lee realmente el motor. La última parte importa más de lo que parece: de las 31 columnas de
`Shipments`, la macro solo usa tres.

## El orden de importación

La secuencia de diálogos es fija y está escrita en el código en este orden
(`tms_fg14/modulo1.vba:159-249`):

| # | Diálogo | Hoja destino | Rango que se copia | Obligatorio |
|---|---|---|---|---|
| 1 | `Seleccionar Shipments` | `Shipments` | `A:AE` | Sí |
    20|| 2 | `Seleccionar STP Equipment Associtation` | `STP-Equipment Assoc` | `A:W` | Sí |
| 3 | `Seleccionar Order Failures` | `STP Failures` | `A:AE` | **No** |
| 4 | `Seleccionar Pallet Container Association` | `Pallet Container Association` | `A:AD` | Sí |
| 5 | `Seleccionar Plan` | `Plan` | `A:AF` | Sí |

Cancelar cualquier diálogo salvo el tercero salta a `Salida` y deja el libro con las hojas
anteriores ya cargadas y las posteriores vacías (`tms_fg14/modulo1.vba:161`, `178`, `214`,
`240`). No hay mensaje de aviso: la macro simplemente termina y activa `User Guide`. Es un
estado inconsistente y hay que volver a correr `ProcesoMacro` desde el principio.

    30|El paso 3 sí avisa y continúa:

```
No se seleccionó archivo Order Failures. Se continuará con el siguiente paso.
```

Antes de pedir el primer archivo, la macro limpia las cinco hojas y apaga cualquier
autofiltro activo (`tms_fg14/modulo1.vba:134-157`). Los anchos de limpieza son 31, 23, 31, 30
y 32 columnas respectivamente.

## Todos los archivos se leen de `Sheets(1)`
    40|
Los cuatro archivos de LBS se leen de la **primera hoja del libro**, sin importar cómo se
llame (`tms_fg14/modulo1.vba:164`, `181`, `199`, `217`). Si el archivo descargado de LBS
tiene una portada, una hoja de parámetros o cualquier cosa antes de los datos, se importa esa
en su lugar y `SummaryOK` produce basura sin quejarse.

El Plan es la excepción: se busca la hoja **por nombre** (`Sheets("Plan")`,
`tms_fg14/modulo1.vba:247-249`) y se abre en solo lectura.

## `Shipments`
    50|
Un renglón por embarque que LBS logró armar. Encabezados tomados de
[tms_fg14/walmart/shipmets.tsv](../../tms_fg14/walmart/shipmets.tsv).

| Col | Encabezado | Lo usa el motor |
|---|---|---|
| `A` | `Shipment Id` | **Sí.** Es la llave. `tms_fg14/modulo2.vba:1546` |
| `B` | `Status` | No |
| `C` | `Equipment Id` | No — el equipo se toma de `STP-Equipment Assoc` |
| `D` | `Origin Location Id` | No |
    60|| `E` | `Destination Location Id` | No |
| `F` | `Shipment Date` | No |
| `G` | `Delivery Date` | No |
| `H` | `Lane Lock Date` | No |
| `I` | `Cube` | No |
| `J` | `Weight` | **Sí.** Peso del embarque según LBS. `tms_fg14/modulo2.vba:1549` |
| `K` | `Load Target` | No |
| `L` | `Load Efficiency` | **Sí.** El fill rate que reporta LBS. `tms_fg14/modulo2.vba:4822` |
| `M` | `Creation Date` | No |
| `N` | `Urgent` | No |
    70|| `O` | `Rank` | No |
| `P` `Q` | `Delivery Window Begin` / `End` | No |
| `R` `S` | `Ship Window Begin` / `End` | No |
| `T` `U` | `Shift Number` / `Is Shift End Of Day` | No |
| `V` | `Laden Length` | No |
| `W` | `Reference Number` | No |
| `X` | `Imperial Units` | No |
| `Y` | `User` | No |
| `Z` | `Lane Id` | **Sí.** De aquí sale la planta origen. Ver abajo |
| `AA` | `Pieces` | No |
    80|| `AB` | `Pallets` | No |
| `AC` `AD` | `Order Value` / `Declared Value` | No |
| `AE` | `Operation Type` | No |

**Ordenamiento:** ascendente por `A`, con encabezado, sobre el rango `A1:AE`
(`tms_fg14/modulo1.vba:166-172`).

### La planta origen sale del `Lane Id`

`SF_BuildShipOrigenLookup` (`tms_fg14/modulo3.vba:137-174`) construye un diccionario de
    90|`Shipment Id` a planta. No lee la columna `D` (`Origin Location Id`): parsea el `Lane Id` de
la columna `Z`.

Y no la busca por letra fija. Usa `SF_HeaderCol(wsShip, "Lane Id", 31)`, es decir **localiza
la columna por el texto de su encabezado** entre las primeras 31 columnas
(`tms_fg14/modulo3.vba:152-153`). Lo mismo para `Shipment Id`, con `A` como respaldo si no lo
encuentra. Esto hace la lectura tolerante a que LBS reordene columnas, pero **no** a que
traduzca los encabezados.

Un `Lane Id` típico es `BOD_DS_PC05_400060001`: la planta es el `PC05` del medio. La misma
planta se registra bajo varias llaves —el id crudo, el sufijo tras el último guion y ese
   100|sufijo con el prefijo `P-`— para que se pueda encontrar tanto desde el `Shipment Id` como
desde el folio `AD` de `Pedidos Surtidos`, que tiene la forma `DDMMAAAAP-<shipment>`
(`tms_fg14/modulo3.vba:126-134`, y la búsqueda inversa en `SF_DeriveOrigenHint`,
`tms_fg14/modulo3.vba:176-207`).

## `STP-Equipment Assoc`

Un renglón por pedido colocado, con el equipo que le tocó. Es la hoja más angosta que se
importa (`A:W`, 23 columnas) y de la que menos se usa: el motor la lee como bloque `A1:Z`
(`tms_fg14/modulo2.vba:1087`) y solo consulta cuatro columnas.
   110|
| Col | Rol | Cita |
|---|---|---|
| `A` | `Shipment Id`. Llave de `dictSTPShip` | `tms_fg14/modulo2.vba:1099` |
| `B` | `Equipment Id`. Se traduce a nombre legible contra la hoja `Equipments` | `tms_fg14/modulo2.vba:1379-1380` |
| `C` | `Order Id`. Llave de `dictSTP`, canonizada y cortada en el primer `_` | `tms_fg14/modulo2.vba:1098` |
| `J` | Fecha programa. Va a `Pedidos Surtidos!AK` | `tms_fg14/modulo2.vba:1393` |
| `L` | Vigencia. Respaldo cuando el Plan no la trae | `tms_fg14/modulo2.vba:1232` |

Las demás columnas (cantidades embarcables, item, porcentajes de tarima) las lee LBS pero no
la macro.
   120|
### El `Order Id` se canoniza

`dictSTP` no guarda el `Order Id` tal cual. Aplica dos transformaciones
(`tms_fg14/modulo2.vba:1098`):

1. `Split(..., "_")(0)` corta en el primer guion bajo. De `1127992973_400084097_2913` queda
   `1127992973`.
2. `SF_CanonPedido` normaliza el resultado.

   130|El paso 2 existe por un problema concreto de precisión: los pedidos de algunas cadenas son
números que exceden el rango de un `Long` de VBA. `validate_summaryok_preflight.py:63-68`
tiene una comprobación específica para eso y emite `INFO: STP col C pedidos > Long: N
(canonical keys required)`. Si el pedido se convirtiera a número se perdería resolución y dos
pedidos distintos colapsarían en la misma llave.

El mismo script advierte cuando la columna `C` viene vacía (`WARN: STP col C empty on N
row(s)`), que es la causa habitual de que un embarque quede sin equipo asignado.

**Ordenamiento:** ascendente por `A`. Nótese que el `SetRange` dice `A1:AE` aunque la hoja
   140|solo tiene datos hasta `W` (`tms_fg14/modulo1.vba:186`); es inofensivo porque las columnas
extra están vacías.

## `STP Failures`

Los pedidos que LBS **no** pudo colocar, con el motivo. Es la única entrada opcional.
Encabezados tomados de [tms_fg14/rollback/stp faillures.tsv](../../tms_fg14/rollback/stp%20faillures.tsv).

| Col | Encabezado | Para qué sirve al operador |
|---|---|---|
| `A` | `Order Id` | El pedido que falló |
   150|| `B` | `Lane Id` | El carril, y con él la planta origen |
| `C` `D` | `From Location Id` / `To Location Id` | Origen y destino |
| `E` | `Item Id` | El SKU, con el sufijo de cadena incluido |
| `F` | **`Reason`** | El motivo. La columna que hay que leer |
| `G` | `Total Shippable Quantity` | Cajas totales |
| `H` | `Scheduled Shippable Quantity` | Cuántas sí se programaron |
| `I` | `Failed Shippable Quantity` | Cuántas quedaron fuera |
| `J` | `Raw Per Shippable Quantity` | Factor de conversión |
| `K` `L` `M` | `Total` / `Scheduled` / `Failed Raw Quantity` | Lo mismo en unidades crudas |
| `N` | `Approved` | Bandera de LBS |
   160|| `O`…`S` | `Start Date`, `End Date`, `Arrival Date`, `Demand Due Date`, `Lock Date` | Las fechas de la ventana |
| `T`…`W` | `Initially Peggable`, `Initial Pegging Violation`, `Initial Lock Violation`, `Initially Unlocked` | Diagnóstico interno de LBS |
| `X` `Y` | `Criticality` / `Inv Adj Criticality` | Prioridad |
| `Z` `AA` | `Equipment Availability` / `Equipment Target Type` | Diagnóstico de equipo |
| `AB` | `Doesnt Fit Equipment` | Marca de que no cabe |
| `AC` | `Inv Adj Shippable Times` | Diagnóstico de inventario |
| `AD` | `User` | `System` cuando lo generó LBS |

### Los motivos que más se repiten

   170|| Texto en `F` | Qué pasó | Dónde se corrige |
|---|---|---|
| `Missing required To facility, Invalid Item` | El destinatario o el item no existen en los maestros del export | Macro de entrada: `SiteMaster` / `ItemMaster_ABPP`. Ver [06-validaciones-mexka.md](../merged/06-validaciones-mexka.md) |
| `likely failed due to stp group <grupo>` | El pedido cayó en un grupo STP que se saturó | Suele resolverse solo cuando `SummaryFallo` remonta el resto |
| Referencias a `Equipment` | No había equipo disponible en la ventana | Macro de entrada: revisar `Plan!W1` y la ventana de `equipmentByLaneByDay` |

Un patrón concreto de la primera fila del ejemplo vale la pena señalar: el `Item Id`
`3018106_WAL_BA` con motivo `likely failed due to stp group 33_PC13_400084097`. Los grupos
STP `33_*` y `20_*` son los de la excepción EXC28 de Walmart, y su fallo es esperable cuando
el pedido no alcanza a llenar el camión de 28 tarimas.

   180|**Ordenamiento:** ascendente por `A`, rango `A1:AE` (`tms_fg14/modulo1.vba:201-207`).

## `Pallet Container Association`

**La entrada más importante.** Un renglón por cada unidad física que LBS armó: qué tarima,
con qué SKU, cuántas cajas y cuánto pesa. Es de aquí de donde nace cada fila de
`Pedidos Surtidos`. Encabezados tomados de [tms_fg14/oxxo/pca.tsv](../../tms_fg14/oxxo/pca.tsv).

| Col | Encabezado | Lo usa el motor |
|---|---|---|
| `A` | `Loading Seq` | No (solo como respaldo para contar filas) |
   190|| `B` | `Load Space Index` | No |
| `C` | `Shipment Id` | **Sí.** Forma el folio `AD` |
| `D` | `Equipment Id` | No |
| `E` `F` | `Lane Id` / `Actual Lane Id` | No |
| `G` | **`Unit Id`** | **Sí.** El identificador de la unidad física. Va a `PS!D` |
| `H` | **`Unit Type`** | **Sí.** `SINGLE`, `PALLET` o `SANDWICH`. Va a `PS!E` |
| `I`…`L` | `X1` `Y1` `X2` `Y2` | No — la geometría en el piso del camión |
| `M` | `Deck Level` | No |
| `N` `O` `P` | `Sub LayerId` / `Sub UnitId` / `Sub UnitType` | No |
| `Q` | **`ItemId`** | **Sí.** El SKU con sufijo de cadena. Va a `PS!X` |
   200|| `R` | `Raw Quantity` | No |
| `S` | **`Order Raw Quantity`** | **Sí.** El cartonaje. Va a `PS!S` |
| `T` | **`Weight`** | **Sí.** El peso en kg. Va a `PS!AT` |
| `U` | `Cum Weight` | No |
| `V` | `Height` | No — la altura se recalcula desde `TI HI` |
| `W` | `Cum Height` | No |
| `X` `Y` | `Volume` / `Cum Volume` | No |
| `Z` | **`OrderId`** | **Sí.** El pedido. Va a `PS!P` tras cortar en el primer `_` |
| `AA` | `Stack Group` | No |
| `AB` | `Rank` | No |
   210|| `AC` | `Is Urgent` | No |
| `AD` | `Is ProductionPlan` | No |

### El ordenamiento importa

Es la única hoja que **no** se ordena por `A`. Se ordena por dos llaves
(`tms_fg14/modulo1.vba:220-223`):

1. `C` (`Shipment Id`) ascendente
2. `G` (`Unit Id`) ascendente

   220|Ese orden no es cosmético: `SummaryOK` recorre la hoja secuencialmente comparando cada
`Unit Id` con el anterior para decidir a qué renglón le asigna el conteo de tarima en `W`
(`tms_fg14/modulo2.vba:757-764`). Si la hoja llega desordenada, los renglones de una misma
unidad quedan separados y el conteo de tarimas sale mal.

### El filtrado de la columna `Z`

Después de ordenar, la macro borra hacia atrás **toda fila cuya columna 26 (`Z`, `OrderId`)
esté vacía** (`tms_fg14/modulo1.vba:226-234`):

```vba
   230|For i = ultFila To 2 Step -1
    If Trim(.Cells(i, 26).Value) = "" Then
        .Rows(i).Delete
    End If
Next i
```

Son los renglones de relleno que LBS emite para describir espacio vacío del camión. Sin
pedido no hay nada que despachar. El recorrido va de abajo hacia arriba precisamente para que
el borrado no desplace las filas pendientes de revisar.

   240|### El tope de filas

`SummaryOK` valida el tamaño de esta hoja aparte, con su propio límite
(`tms_fg14/modulo2.vba:644-649`):

```
PCA tiene N filas (max <LBS_WALMART_PCA_MAX_ROWS>).
Revise celdas sueltas al final de Pallet Container Association.
```

   250|El comentario del código explica por qué se lee por bloques y no de un tirón
(`tms_fg14/modulo2.vba:642`):

```
' Filas en PCA (G preferred, A fallback). Never load A1:Z in one shot — hard-closes Excel.
```

Cargar `A1:Z` completo de una PCA grande no da un error: **cierra Excel sin aviso**. De ahí
que todo el volcado se haga en trozos de `LBS_WALMART_PCA_CHUNK` filas.

`validate_summaryok_preflight.py:46-51` avisa antes de correr:
   260|`WARN: PCA rows=N (approaching old 10000 cap; patched macro required)`.

## `Plan`

El mismo Plan que produjo la macro de entrada, con el mismo mapa de columnas descrito en
[merged/02-plan-y-parametros.md](../merged/02-plan-y-parametros.md). Los datos empiezan en la
**fila 3**.

Diferencias en cómo lo lee esta macro:

   270|- **Se importa solo `A:AF`** (32 columnas). El comentario de `TMS_PLAN_LAST_COL` explica que
  leer hasta `CC` (81 columnas) inflaba el rango por encima de las 65 536 celdas y provocaba
  `Overflow (6)` (`tms_fg14/modulo1.vba:1`).
- **El último renglón se detecta con cuatro columnas clave**, no una:
  `TMS_PlanDataLastRow` usa `I` (destino), `J` (pedido), `L` (vigencia) y `AA` (cartonaje
  exportable) y toma el máximo (`tms_fg14/modulo1.vba:84-87`).
- **El motor lo carga como bloque `A1:AD`** con `Value2` (`tms_fg14/modulo3.vba:584`), y
  construye el índice desde la fila 3.

### La llave de cruce con `Pedidos Surtidos`

   280|`SF_BuildPlanLookup` (`tms_fg14/modulo3.vba:573-595`) indexa el Plan por:

```
SF_CanonPedido(Plan!J) & "|" & Plan!T
```

Es decir **pedido + material**. Una particularidad: el pedido se lee con `.Text`, no con
`.Value` (`tms_fg14/modulo3.vba:588`), para tomar exactamente lo que muestra la celda y no la
representación numérica interna. El material se lee del arreglo, columna 20 (`T`).

Cuando una fila de `Pedidos Surtidos` no encuentra su pareja en el Plan, todos los campos que
   290|se copian desde ahí —cadena, CEDIS, destinatario, vigencia, armado, tipo de tarima— quedan
vacíos, y la fila queda esencialmente inservible. Es el síntoma de haber importado un Plan
distinto al que se usó para generar el export de LBS.

### `SF_BuildPlanOrigenLookup` y los pedidos de doble planta

Hay un segundo índice, más fino, que incluye la planta:
`pedido|material|ORIGEN -> fila` (`tms_fg14/modulo3.vba:228-230`).

Junto con él se construye opcionalmente `dictPlanDual`, que marca las llaves
`pedido|material` que aparecen bajo **dos o más plantas distintas** en el Plan
   300|(`tms_fg14/modulo3.vba:218-224`). El comentario lo describe como *riesgo de stamp cruzado*:
si el mismo pedido y material se surten desde dos plantas, copiar ciegamente el origen de la
primera coincidencia le pondría la planta equivocada a la mitad de las filas.

### El error de importación del Plan

Es el único paso con manejo de error propio. `planStep` va guardando en qué punto está
(`Abrir Plan`, `Detectar filas Plan`, `Copiar Plan A:AF (N filas)`) y el mensaje de error lo
incluye junto con la sugerencia más útil (`tms_fg14/modulo1.vba:253-257`):

```
Error al importar Plan (<paso>):
   310|<descripción> (<número>)
Filas detectadas: N
Sugerencia: use Planificacion_Plantas_Restos.xlsm (no el export MEX KA).
```

Esa sugerencia apunta al error más frecuente: seleccionar el archivo `MEX KA PLANTS_Restos_v*.xlsx`
en lugar del libro de planeación. El export MEX KA no tiene hoja `Plan`, así que
`Sheets("Plan")` falla.

## Resumen de verificación antes de `SummaryOK`
   320|
| Hoja | Qué revisar |
|---|---|
| `Shipments` | Que la fila 1 tenga los encabezados en inglés, en particular `Lane Id` y `Shipment Id` |
| `STP-Equipment Assoc` | Que la columna `C` no tenga celdas vacías |
| `STP Failures` | Leer la columna `F` antes de continuar; ahí está lo que LBS no pudo colocar |
| `Pallet Container Association` | Que el conteo de filas sea razonable y que ya no queden filas con `Z` vacía |
| `Plan` | Que sea el mismo Plan de la corrida y que los datos arranquen en la fila 3 |

El script [scripts/validate_summaryok_preflight.py](../../scripts/validate_summaryok_preflight.py)
automatiza buena parte de esto sobre el `.xlsm` ya cargado.
