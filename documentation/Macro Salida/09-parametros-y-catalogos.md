[Volver al índice de la macro de salida](README.md)

# Parámetros y catálogos

Todos los números que gobiernan el armado están en uno de dos lugares: **constantes en el
código** (hay que editar el VBA para cambiarlas) u **hojas de parámetros del libro** (las
edita el planeador sin tocar código). Esta página es el inventario completo de ambos.

La regla general de diseño es: si el número lo negocia el cliente y cambia por carril o por
SKU, vive en una hoja. Si es un límite estructural o un piso acordado a nivel cadena, vive
como constante.

---

## Parte 1 — Constantes

Cada tabla reproduce el comentario original del VBA, porque ahí está la justificación de
negocio: fechas de acuerdo con el cliente, referencias a `Max Pallet Count` del maestro de
equipos, notas de bugs de Excel.

### Guardas de tamaño e importación

Archivo [tms_fg14/modulo1.vba](../../tms_fg14/modulo1.vba).

| Constante | Valor | Línea | Comentario original |
|---|---|---|---|
| `TMS_PLAN_LAST_COL` | `32` (= `AF`) | `modulo1.vba:2` | `' Plan TMS: columnas utiles A:AF (32). CC (81) inflaba el rango >65536 celdas -> Overflow (6).` |
| `TMS_COPY_CHUNK_ROWS` | `200` | `modulo1.vba:3` | (sin comentario) |
| `TMS_MAX_IMPORT_ROWS` | `50000` | `modulo1.vba:4` | (sin comentario) |
| `TMS_MAX_CELLS_PER_OP` | `60000` | `modulo1.vba:6` | `' Excel: evitar pasar > ~65000 celdas en un solo .Value / .ClearContents` |

Detalle en la sección "Guardas de tamaño" del [README](README.md#guardas-de-tamaño).

### Estructura de `Pedidos Surtidos`

| Constante | Valor | Línea | Comentario original |
|---|---|---|---|
| `SK_PS_LAST_COL` | `46` (= `AT`) | `modulo2.vba:112` | `' Pedidos Surtidos A:AT = 46 cols; keep clear/write under ~60k cells per COM op.` |
| `SK_PS_CLEAR_CHUNK` | `1000` | `modulo2.vba:113` | (sin comentario) |
| `LBS_WALMART_PCA_CHUNK` | `1000` | `modulo2.vba:109` | (parte del bloque de caché plano de PCA) |
| `LBS_WALMART_PCA_MAX_ROWS` | `250000` | `modulo2.vba:110` | `' Flat PCA height cache: unit|mat|qty -> Double (sum of unique O layers). Avoids nested Dictionary-per-row + full A1:Z load that hard-closes Excel on large PCA sheets.` |

`SK_PS_LAST_COL = 46` llega hasta `AT`, no hasta `AV`. Las dos columnas que quedan fuera
(`AU` peso del camión y `AV` `REVISION MANUAL`) se escriben celda por celda precisamente
porque no forman parte del bloque masivo. Ver
[03-pedidos-surtidos.md](03-pedidos-surtidos.md).

### Cupos de camión (tarimas)

| Constante | Valor | Línea | Comentario original |
|---|---|---|---|
| `LBS_METRO_TRUCK_CAP` | `26` | `modulo2.vba:19` | `' LBS - capacidad de camion (tarimas) para montar restos por metro (Soriana/City Club).` |
| `LBS_WALMART_TRUCK_CAP` | `28` | `modulo2.vba:21` | `' LBS - capacidad de camion (tarimas) para los grupos WALMART (Bodega Aurrera + SuperCenter).` |
| `LBS_FULL_TRUCK_CAP` | `40` | `modulo2.vba:25` | `' LBS - capacidad maxima por confirmacion COMEXTRA Full caja seca (catalogo mode mix = 40).` |
| `LBS_FULL_BOX_CAP` | `20` | `modulo2.vba:29` | `' LBS - caja (a o b) de un Full: ~20 tarimas por sufijo de camion (PartirFulles).` |
| `LBS_COMEXTRA_SPLIT_CAP` | `20` | `modulo2.vba:27` | `' LBS - aviso blando: LBS suele armar splits a/b de ~20 tarimas por confirmacion.` |

El `26` es el cupo por omisión de un sencillo y aparece en el nombre como "metro" porque el
caso original fue el armado por metro de Soriana / City Club, pero se usa como fallback
general.

### Cupos de OXXO

| Constante | Valor | Línea | Comentario original |
|---|---|---|---|
| `LBS_OXXO_BOX_CAP` | `18` | `modulo2.vba:31` | `' LBS - fallback OXXO Full shipment / box when Catalogo Mode Mix miss.` |
| `LBS_OXXO_SHIPMENT_CAP` | `36` | `modulo2.vba:32` | (mismo bloque) |
| `LBS_OXXO_SENCILLO_CAP` | `24` | `modulo2.vba:35` | `' OXXO Sencillo from equipment Max Pallet Count (client 29/07/2026): Z5290_OXX=24 (not catalog S/26 / generic Z5290), Z4290_OXX=22, Z5290_OXX_MTY=28.` |
| `LBS_OXXO_Z4290_SENCILLO_CAP` | `22` | `modulo2.vba:36` | (mismo bloque) |
| `LBS_OXXO_MTY_SENCILLO_CAP` | `28` | `modulo2.vba:37` | (mismo bloque) |
| `LBS_OXXO_MTY_DEST` | `"400101621"` | `modulo2.vba:38` | (mismo bloque) |

El comentario fecha el acuerdo con el cliente (29/07/2026) y dice explícitamente que los
cupos de sencillo de OXXO vienen del `Max Pallet Count` del equipo y **no** del catálogo
Mode Mix. Es la excepción más documentada del código.

### Cupos de mayoristas

| Constante | Valor | Línea | Comentario original |
|---|---|---|---|
| `LBS_CAP35_SHIPMENT` | `35` | `modulo2.vba:40` | `' SKU-level Full shipment cap (sheet "Cadenas 35 Tarimas": allowlisted cadena + SKU, Viajes a 35=YES).` |
| `LBS_CAP35_SHEET` | `"Cadenas 35 Tarimas"` | `modulo2.vba:41` | (mismo bloque) |
| `LBS_MAYORISTA_FULL_SHIPMENT` | `36` | `modulo2.vba:43` | `' Allowlisted mayorista Full (non-listed SKU): single-truck ceiling 36 (no a/b).` |
| `PF_CAP35_SHIPMENT` | `35` | `modulo5.vba:6` | `' SKU-level Full shipment cap (sheet "Cadenas 35 Tarimas": allowlisted cadena + SKU).` |
| `PF_CAP35_SHEET` | `"Cadenas 35 Tarimas"` | `modulo5.vba:7` | (mismo bloque) |
| `PF_MAYORISTA_FULL_SHIPMENT` | `36` | `modulo5.vba:8` | (mismo bloque) |

`modulo5` duplica las tres constantes de mayoristas porque `PartirTarimasFULL` tiene su
propia copia de la lógica de cupos y su propia caché. **Si se cambia una, hay que cambiar la
otra.** Ver [mayoristas-cap35.md](cadenas/mayoristas-cap35.md).

### Pisos de llenado

Son fracciones del cupo. El piso efectivo en tarimas es `piso x cap`, redondeado según el
sitio de uso.

| Constante | Valor | Piso efectivo | Línea | Comentario original |
|---|---|---|---|---|
| `LBS_WALMART_MIN_FILL` | `0.4` | 12 de 28 | `modulo2.vba:65` | `' LBS - WALMART: piso de llenado de camion (tarimas). El % de "EFICIENCIA POR CADENA" es fill rate de PEDIDO (col AR, gate de eficiencia) para Walmart/Alsuper — truck floor is this constant (cap 28 -> piso 12). ClubCity (Soriana/City Club) truck floor is LBS_CLUBCITY_MIN_FILL (70% of metro cap 26 -> piso 19).` |
| `LBS_CLUBCITY_MIN_FILL` | `0.7` | 19 de 26 | `modulo2.vba:67` | `' LBS - CLUBCITY (Soriana + City Club): piso de llenado post-consolidacion (70% del cap).` |
| `LBS_OXXO_MIN_FILL` | `0.9` | 33 de 36 | `modulo2.vba:69` | `' LBS - OXXO Full: piso de llenado post-consolidacion (90% del shipCap 36 -> piso 33).` |
| `LBS_COMEXTRA_MIN_FILL` | `0.9` | 36 / 18 / 24 | `modulo2.vba:73` | `' LBS - COMEXTRA: piso de llenado post-consolidacion (90% del cap). Full shipment / unsuffixed leftover: shipCap 40 -> piso 36. Full a/b caja: 20 -> piso 18. Sencillo: 26 -> piso 24. Fill toward 40 when possible; under piso -> baja eficiencia.` |
| `LBS_LACOMER_MIN_FILL` | `0.8` | 21 de 26 | `modulo2.vba:78` | `' LBS - LA COMER: piso de llenado post-consolidacion (cap 26 -> piso 21 = 80%).` |
| `LBS_CHEDRAUI_MIN_FILL` | `0.8` | 21 de 26 | `modulo2.vba:81` | `' LBS - CHEDRAUI: piso de llenado post-consolidacion (cap 26 -> piso 21 = 80%). Independent of EFICIENCIA POR CADENA (AR gate); truck floor is always 80%.` |

El comentario de `LBS_WALMART_MIN_FILL` es el más importante de todo el archivo, porque
resuelve la confusión que más se repite en operación: **el porcentaje de la hoja `EFICIENCIA
POR CADENA` y el piso de llenado del camión son dos cosas distintas.** El primero se compara
contra `AR` (fill rate del pedido) y el segundo contra las tarimas físicas del camión. Un
camión de Walmart con 15 tarimas pasa el piso (12) aunque el pedido tenga un fill rate bajo,
y viceversa. Detalle en [05-eficiencia-y-descartes.md](05-eficiencia-y-descartes.md).

### Rescate de pedidos y grupos

| Constante | Valor | Línea | Comentario original |
|---|---|---|---|
| `LBS_CHEDRAUI_RESCUE_MIN_FILL` | `0.8` | `modulo2.vba:56` | `' LBS - CHEDRAUI: rescate de pedido todo-fallido si la carga combinada >= este % del cap.` |
| `LBS_WALMART_RESCUE_MIN_FILL` | `0.5` | `modulo2.vba:60` | `' LBS - WALMART: rescate de grupo todo-fallido si la carga combinada >= este % del cap (28 tarimas). Solo para el gate de grupo (LBS_WalmartGroupEfficiencyGate); el piso post-consolidacion es LBS_WALMART_MIN_FILL.` |

El rescate es lo contrario del descarte: cuando **todas** las líneas de un pedido (Chedraui)
o de un grupo (Walmart) reprobaron el gate de eficiencia, pero juntas llenarían suficiente
camión, se rescatan en lugar de tirarse. Walmart rescata desde el 50 % del cap, Chedraui
desde el 80 %.

### Peso

| Constante | Valor | Línea | Comentario original |
|---|---|---|---|
| `SK_MAX_PESO_KG` | `29000` | `modulo2.vba:10` | (sin comentario propio) |
| `LBS_FULL_MAX_PESO_KG` | `52500` | `modulo2.vba:12` | `' Catalogo Mode Mix Full weight ceiling (52.5 t) for Alsuper/Go Mart Full lanes.` |
| `PF_MAX_PESO_KG` | `29000` | `modulo5.vba:2` | (sin comentario) |
| `SK_PESO_TARIMA_DEFAULT` | `30` | `modulo2.vba:14` | `' Tarima tare (kg). All 30 until tech specs; LBS_PesoTarimaKg picks by type.` |
| `SK_PESO_TARIMA_CHEP` | `30` | `modulo2.vba:15` | `' TODO tech specs` |
| `SK_PESO_TARIMA_PLASTICA` | `30` | `modulo2.vba:16` | `' TODO tech specs` |
| `SK_PESO_TARIMA_MADERA` | `30` | `modulo2.vba:17` | `' TODO tech specs` |

Las cuatro taras valen 30 kg. La estructura está lista para diferenciarlas por tipo de
tarima (`LBS_PesoTarimaKg`, `modulo2.vba:11490`) en cuanto el cliente entregue las fichas
técnicas; los `TODO tech specs` son de los autores del código, no de esta documentación.

El techo normal es 29 t. Los carriles Full de Alsuper y Go Mart usan 52.5 t porque el
catálogo Mode Mix lo declara así. `LBS_MaxPesoKgForRow` (`modulo2.vba:11506`) es quien
decide cuál aplica por fila.

### Altura

| Constante | Valor | Línea | Comentario original |
|---|---|---|---|
| `LBS_WALMART_MAX_HEIGHT_M` | `1.6` | `modulo2.vba:23` | `' LBS - WALMART: altura maxima de tarima/unidad (suma AO -> AP) en metros.` |
| `LBS_LACOMER_MAX_HEIGHT_M` | `1.6` | `modulo2.vba:83` | `' LBS - LA COMER: altura maxima de tarima/unidad (suma AO -> AP), trial = Walmart 1.6 m.` |
| `LBS_WALMART_LAYER_ALTO_TOL_CM` | `0.2` | `modulo2.vba:99` | `' Shared-cama match: Abs(AltoA-AltoB) <= 0.2 cm. Mixed cluster uses TI-1 cases/cama. Applies to Armado-en-cama TI HI chains (LBS_ChainAllowsSharedCamas).` |
| `LBS_WALMART_TIHI_BL_TOL_M` | `0.005` | `modulo2.vba:96` | (tolerancia de comparación contra la línea base) |

La palabra `trial` en el comentario de La Comer es literal: el límite de 1.60 m para La
Comer se puso a prueba copiando el de Walmart y no viene de una especificación propia del
cliente.

La tolerancia de 0.2 cm es la que decide si dos SKU distintos pueden compartir una cama: si
la altura de caja difiere en más de 2 mm, no comparten. En un clúster mezclado se usa `TI-1`
cajas por cama en lugar de `TI`, para dejar holgura.

### Catálogo TI HI

| Constante | Valor | Línea |
|---|---|---|
| `LBS_WALMART_TIHI_SHEET` | `"TI HI"` | `modulo2.vba:94` |
| `LBS_WALMART_TIHI_MIN_KEYS` | `50` | `modulo2.vba:95` |
| `LBS_WALMART_TIHI_URL` | URL de SharePoint a `TI HI VALIDADO ACOMODOS.xlsx` | `modulo2.vba:92-93` |

El comentario del bloque (`modulo2.vba:89-91`) fija la regla operativa más importante del
catálogo:

```
' LBS - WALMART TI HI cache: key cadena|material -> Array(TI, HI, AltoCm, AltArmCm, PesoCase)
' Lookups read only from in-workbook sheet "TI HI" (never open files during SummaryOK).
' Populate that sheet with Public LBS_SyncTihiSheet (User Guide button).
```

Es decir: durante `SummaryOK` la macro **nunca** abre el archivo de SharePoint. Solo lee la
hoja `TI HI` del propio libro. Actualizar esa hoja es un paso manual y previo.
`LBS_WALMART_TIHI_MIN_KEYS = 50` es el umbral por debajo del cual se considera que el
catálogo no cargó bien.

### Banderas y límites de proceso

| Constante | Valor | Línea | Comentario original |
|---|---|---|---|
| `SK_PROGRESS_MIN_SEC` | `0.5` | `modulo2.vba:9` | `' Throttle Plan!X1 + DoEvents (~2/sec). Phase-name changes always flush.` |
| `SF_PROGRESS_MIN_SEC` | `0.5` | `modulo3.vba:5` | `' Throttle Plan!X1 + DoEvents (~2/sec). Phase-name changes always flush.` |
| `LBS_IGNORE_SHIPPABLE_QTY_ZERO_DEFAULT` | `False` | `modulo2.vba:85` | `' TEST MODE Plan!V1: allow remount of "Shippable qty is 0" AG rows (default off).` |
| `LBS_WALMART_SALVAGE_TRACE` | `False` | `modulo2.vba:76` | `' Temporary: Debug.Print salvage decisions (ckey / spare / AO / skip) in Fill+Rebuild. Keep False for normal runs — True floods Immediate window and can freeze Excel.` |
| `PF_MAX_GROUP_EXTRA_ROWS` | `50` | `modulo5.vba:3` | (tope de filas insertadas por grupo al partir) |
| `PF_MAX_OVERFLOW_SEQ` | `20` | `modulo5.vba:4` | (tope de camiones de overflow `R2`, `R3`, …) |

`LBS_WALMART_SALVAGE_TRACE` merece una advertencia: el comentario dice que ponerla en `True`
inunda la ventana Inmediato y **puede congelar Excel**. Solo se activa para depurar un caso
concreto, con pocos datos.

---

## Parte 2 — Hojas de parámetros

Estas sí las edita el planeador. Todas se leen con caché por sesión, así que un cambio en
media corrida no se refleja hasta la siguiente lectura; `LBS_ResetConsMap`
(`modulo2.vba:5256`) limpia las cachés de `Consolida`, `Cadenas 35 Tarimas` y `EFICIENCIA
POR CADENA` de golpe.

### `Catalogo Mode Mix`

Es el mismo catálogo de carriles que usa la macro de entrada. Aquí se lee solo para obtener
el cupo de tarimas por carril.

| Columna | Encabezado | Uso en esta macro |
|---|---|---|
| `A` | `Key` | — |
| `B` | `ID Origen Moderno` | Parte de la llave de carril |
| `C` | `Origen Moderno` | — |
| `D` | `Destinatario` | Parte de la llave de carril; define el último renglón útil |
| `E` | `Cadena` | — |
| `F` | `Destino` | — |
| `G` | `KM Moderno` | — |
| `H` | `Mode Mix` | `F` (Full) o `S` (Sencillo). Cualquier otro valor descarta el renglón |
| `I` | `Tipo Equipo` | — |
| `J` | `Especializado` | — |
| `K` | `Pallets Max` | **El cupo.** Debe ser numérico y mayor que cero |
| `L` | `Peso Max` | — |

`LBS_EnsureCatalogCaps` (`modulo2.vba:5606`) construye cuatro diccionarios: cupo Full y cupo
Sencillo, cada uno por carril (`ID Origen Moderno|Destinatario`) y por destinatario a secas.

El diccionario por destinatario tiene una salvaguarda importante
(`modulo2.vba:5668-5679`): si dos renglones del mismo destinatario declaran cupos distintos
para el mismo Mode Mix, el valor se marca como `-1` y ese destinatario queda **inutilizable**
para la búsqueda por destino. Solo el carril completo sirve. Es deliberado: adivinar entre
dos cupos contradictorios produce camiones mal armados que nadie detecta.

La hoja se resuelve por nombre exacto `Catalogo Mode Mix`, y si no existe se busca cualquier
hoja cuyo nombre contenga `Catalogo` y `Mode` (`LBS_GetCatalogoModeMixSheet`,
`modulo2.vba:5555`).

### `Cadenas 35 Tarimas`

Lista blanca de combinaciones cadena + SKU que pueden llevar 35 tarimas por camión.

| Columna | Encabezado | Uso |
|---|---|---|
| `A` | `ID Cliente` | **Se ignora** (`' dest col A ignored`, `modulo2.vba:5299`) |
| `B` | `Cadena` | Se normaliza a mayúsculas sin espacios |
| `C` | `SKU` | Se canoniza con `SF_CanonMaterial` |
| `D` | `Viajes a 35` | Debe ser `YES`, `SI` o `SÍ`. Cualquier otra cosa descarta el renglón |

Datos desde la fila 2. `LBS_EnsureCap35` (`modulo2.vba:5300`) aplica **dos** filtros: la
bandera de la columna `D` y la lista blanca de cadenas del código
(`LBS_IsCap35AllowChain`, `modulo2.vba:5276`), que solo admite `TOBEDISTRIBUTIONS`,
`CONASUPER`, `NETO`, `CABRITOABARROTERO` y `SUPEROFERTAS`.

Consecuencia práctica: **agregar una cadena nueva a la hoja no hace nada.** Si el cliente
pide extender el cupo de 35 a otra cadena, hay que editar la lista blanca del VBA en los dos
módulos (`modulo2.vba:5276` y su equivalente en `modulo5`). Ver
[mayoristas-cap35.md](cadenas/mayoristas-cap35.md).

### `TI HI`

Copia sin filtrar del catálogo homologado `TI HI VALIDADO ACOMODOS.xlsx`. Alimenta el
cálculo de altura (`AO`, `AP`) y el peso por caja.

| Columna | Encabezado | Uso |
|---|---|---|
| `A` | `KEY` | — |
| `B` | `Kam Comercial` | — |
| `C` | `Responsable` | — |
| `D` | `CPFR` | — |
| `E` | `Clasificación` | — |
| `F` | `Cadena` | **Llave.** Define también dónde está el renglón de encabezados |
| `G` | `Material` | **Llave.** Define el último renglón con datos |
| `H` | `Descripción` | — |
| `I` | `Armado` | — |
| `J` | `Cajas x camas` | `TI` |
| `K` | `Camas` | `HI` |
| `L` | `Peso Bruto de material` | Peso por caja, en kg |
| `M` | `Largo (cm) (profundidad)` | — |
| `N` | `Ancho (cm)` | — |
| `O` | `Alto (CM)` | Altura de una caja, en cm |
| `P` | `Tarima` | — |
| `Q` | `Altura Armado` | Altura de la tarima armada, en cm |
| `R` | `Peso por tarima` | — |
| `S`, `T` | Vueltas de playo | — |

`LBS_WalmartTihiLoadSheetIntoDict` (`modulo2.vba:8059`) guarda por cada llave
`cadena|material` el arreglo `Array(TI, HI, AltoCm, AltArmCm, PesoCase)`. Detalles que
importan:

- El encabezado **no está necesariamente en la fila 1**. Se localiza buscando `Cadena` en la
  columna `F`, y normalmente es la fila 2 porque la fila 1 del archivo original está oculta
  (`modulo2.vba:8080-8085`). La sincronización copia también esa fila oculta para que la
  distribución coincida con el archivo fuente.
- Un renglón se descarta si `TI <= 0` o si `Alto (CM) <= 0` (`modulo2.vba:8102`). Sin esos
  dos datos no se puede calcular altura.
- Ante llaves duplicadas gana la que tenga **mayor** `Altura Armado`
  (`modulo2.vba:8107-8111`), que es el criterio conservador: si hay duda, se asume el armado
  más alto.
- Se cargan **todas** las cadenas de la hoja, no solo Walmart
  (`LBS_WalmartTihiIsSupportedCadena`, `modulo2.vba:8055`, devuelve verdadero para cualquier
  cadena no vacía). El prefijo `Walmart` en los nombres de función es histórico: Walmart fue
  el primer cliente que necesitó TI HI.

Cómo se actualiza: [10-diagnostico-y-errores.md](10-diagnostico-y-errores.md#sincronizar-la-hoja-ti-hi).

### `Consolida`

Mapa de destinatario a grupo de consolidación. Es lo que decide qué destinatarios pueden
compartir camión aunque sean CEDIS o cadenas distintas (multi-stop).

| Columna | Encabezado | Uso |
|---|---|---|
| `A` | `Destinatario` | Se compara contra `Pedidos Surtidos!O`. Formato de texto |
| `B` | `Grupo` | Nombre libre del grupo |

Datos desde la fila 2. El comentario del cargador (`modulo2.vba:5217-5224`) describe la
jerarquía completa:

1. Si la hoja `Consolida` existe y tiene al menos un renglón válido, **es la fuente de
   verdad**.
2. Si no existe o está vacía, se usa la tabla embebida `LBS_LoadConsDefaults`
   (`modulo2.vba:5354`), con unos 45 destinatarios agrupados en `CLUBCITY*` y `WALMART*`.
3. Un destinatario que no aparezca en el mapa **es su propio grupo**: viaja solo, no se une a
   nadie.

El punto 3 es una corrección deliberada de un comportamiento anterior. El comentario
(`modulo2.vba:5437-5442`) lo explica:

```
' verdad del multistop. Si el destinatario esta en el mapa, usa ese grupo; si no,
' el destinatario es su propio grupo (no se une a un grupo Consolida por metro del
' CEDIS — eso mezclaba 7482/7494/7464 MX no listados en WALMARTMX, p.ej. P-1095).
```

Antes, un destinatario sin mapa se unía al grupo de su CEDIS por cercanía, y eso metía
destinatarios de Walmart MX no listados dentro de `WALMARTMX`. El folio `P-1095` es el caso
que lo destapó.

La tabla embebida incluye dos entradas que son lo contrario de una consolidación
(`modulo2.vba:5400-5402`):

```
' Destinatarios que NO se consolidan (grupo propio, sin compañeros de camion).
LBS_AddCons "400001590", "6388 SMO"
LBS_AddCons "400055707", "7505 Chalco Secos"
```

Al darles un grupo con nombre único se garantiza que nunca compartan camión. Es el mismo
efecto que no listarlos, pero explícito.

Para crear la hoja desde cero: `LBS_SeedConsolidaSheet`, documentado en
[10-diagnostico-y-errores.md](10-diagnostico-y-errores.md#sembrar-la-hoja-consolida).

### `EFICIENCIA POR CADENA`

Umbral de fill rate por cadena, comparado contra `Pedidos Surtidos!AR`. Se cachea en
`mLBS_EficRef` vía `LBS_LoadEficRefCache` (`modulo2.vba:2581`). Detalle completo en
[05-eficiencia-y-descartes.md](05-eficiencia-y-descartes.md).

Recordatorio: este porcentaje **no** es el piso de llenado del camión. Son controles
distintos que se aplican en momentos distintos.

### `Equipments`

Maestro de equipos. En esta macro se usa para una sola cosa: traducir el código de equipo a
su descripción.

| Columna | Uso |
|---|---|
| `A` | Código de equipo (llave) |
| `B` | Valor que se copia |

`SummaryOK` la carga en `dictEquip` (`modulo2.vba:1109-1117`) recorriendo desde la fila 2. Si
la hoja no existe, el diccionario queda vacío y la macro continúa sin fallar.

### `IDPlantas`

Traduce el código de planta (`PC01`, `PC05`, …) al `ID Origen Moderno` numérico que usa el
catálogo Mode Mix.

| Columna | Uso |
|---|---|
| `A` | `ID Origen Moderno` (numérico) |
| `B` | Código de planta (llave, en mayúsculas) |

Ojo con el orden: la llave está en `B` y el valor en `A`, al revés de lo habitual
(`modulo2.vba:5581-5590`).

`LBS_LoadPcToOrigenMap` (`modulo2.vba:5571`) construye el mapa en dos pasadas: primero
`IDPlantas` y después `OrigenModernoMap` (columna `A` código, columna `C` id). La primera en
registrar cada código gana, así que **`IDPlantas` tiene prioridad sobre `OrigenModernoMap`**
cuando ambas definen el mismo código.

`modulo5` mantiene su propia copia de este cargador (`PF_LoadPcToOrigenMap`,
`modulo5.vba:418`), con la misma duplicación que las constantes de Cap35.

### Celdas de configuración en `Plan`

| Celda | Contenido | Dirección |
|---|---|---|
| `Plan!V1` | Modo prueba: `SI` / `1` / `TRUE` permite remontar filas con `Shippable qty is 0` | Lectura |
| `Plan!X1` | Indicador de progreso: `<fase> \| hh:mm:ss` | Escritura |

`LBS_IgnoreShippableQtyZeroEnabled` (`modulo2.vba:6049`) acepta `SI`, `S`, `1`, `TRUE`, `ON`,
`YES`, `Y` como verdadero y `NO`, `N`, `0`, `FALSE`, `OFF` como falso. Cualquier otro valor,
incluida la celda vacía, cae en el valor por omisión de la constante
`LBS_IGNORE_SHIPPABLE_QTY_ZERO_DEFAULT`, que es `False`.

`Plan!X1` la escriben los tres indicadores de progreso (`SK_SetProgress` en `modulo2.vba:505`,
`SF_SetProgress` en `modulo3.vba:7`, `PF_SetProgress` en `modulo5.vba:1405`), con un límite
de dos actualizaciones por segundo salvo cambio de fase.

---

## Qué se cambia dónde

Resumen para cuando el cliente pide un ajuste:

| Petición del cliente | Dónde se cambia |
|---|---|
| "Este carril ahora lleva 30 tarimas" | Hoja `Catalogo Mode Mix`, columna `K` |
| "Este SKU de Neto va a 35" | Hoja `Cadenas 35 Tarimas`, columna `D` = `YES` |
| "Extender el cupo de 35 a una cadena nueva" | Código: `LBS_IsCap35AllowChain` en `modulo2.vba:5276` **y** su gemela en `modulo5` |
| "Estos dos destinatarios viajan juntos" | Hoja `Consolida`: mismo `Grupo` en la columna `B` |
| "Este destinatario ya no se consolida" | Hoja `Consolida`: quitarlo, o darle un grupo propio con nombre único |
| "Subir el fill rate mínimo de una cadena" | Hoja `EFICIENCIA POR CADENA` |
| "Bajar el piso de llenado de camión de Walmart" | Código: `LBS_WALMART_MIN_FILL` en `modulo2.vba:65` |
| "Cambió el TI HI de un SKU" | Actualizar el archivo homologado y correr `LBS_SyncTihiSheet` |
| "Subir el cupo de sencillo de OXXO" | Código: `LBS_OXXO_SENCILLO_CAP` en `modulo2.vba:35` (no el catálogo) |
| "Cambió el peso máximo de un camión" | Código: `SK_MAX_PESO_KG` / `LBS_FULL_MAX_PESO_KG`; `PF_MAX_PESO_KG` en `modulo5` |
