[Volver a la macro de salida](README.md) · [Índice general](../README.md)

# El motor de armado de cargas

Este documento explica los conceptos que atraviesan toda la macro: el folio, la llave de
consolidación, el cupo, el sándwich, la altura y el peso. Las reglas específicas de cada
cliente están en [cadenas/](cadenas/README.md); aquí está la maquinaria que todas comparten.

## El folio `AD` y sus sufijos

    10|El folio identifica un camión. `SummaryOK` lo construye al volcar la PCA
(`tms_fg14/modulo2.vba:671`, `773`):

```
Format(Date, "DDMMYYYY") & "P-" & <Shipment Id de la PCA>
```

Por ejemplo `02092026P-1020`. La fecha es la del día en que se corre la macro, no la de la
carga: **volver a correr la macro al día siguiente cambia todos los folios.**

Tres funciones manejan el folio, y conviene distinguirlas:
    20|
| Función | Cita | Qué hace |
|---|---|---|
| `SK_ShipmentIdFromAD` | `tms_fg14/modulo2.vba:406-421` | Devuelve el `Shipment Id` limpio, sin fecha ni sufijo |
| `SK_HasTruckSuffix` | `tms_fg14/modulo2.vba:423-431` | Verdadero si el folio ya trae sufijo `a` o `b` |
| `LBS_BaseFolioAD` | `tms_fg14/modulo2.vba:4752-4768` | Devuelve el folio sin sufijo |

`SK_HasTruckSuffix` es la más importante de las tres, porque **el cupo depende de ella.** Un
Full sin sufijo se mide contra el cupo de embarque completo (40 tarimas); el mismo Full ya
partido en `...a` y `...b` se mide contra el cupo de caja (20). Ver la sección de cupo.

    30|El sufijo lo pone `PartirTarimasFULL` ([06-partir-tarimas-full.md](06-partir-tarimas-full.md)).
Cuando se abren camiones nuevos durante el remonte, el folio se acuña con
`LBS_ReassignRowsToFolioAD` (`tms_fg14/modulo2.vba:4784-4799`).

## La llave de consolidación

Dos filas pueden compartir camión solo si comparten la llave de consolidación. La construye
`LBS_ConsolidaKey` (`tms_fg14/modulo2.vba:1460-1474`), y su comentario la define así:

```
' LBS - Clave de camion (consolidacion) = Origen/plant (col L) + grupo [+ La Comer R/C]
    40|' + vigencia (col R). Hard rule: never mix distinct Plan vigencias on one truck.
' "" cuando la fila no participa en consolidacion.
```

En concreto:

```
UCase(Plan!L) | grupo de consolidación [ | R/C de La Comer ] | vigencia YYYYMMDD
```

Los cuatro componentes:
    50|
### 1. La planta (`L`)

Nunca se mezclan plantas en un camión. Cuando `L` trae un marcador de inventario
(`NoInventario`, `FaltaInventario`, `FABRICA`) se sustituye por el origen derivado del
`Lane Id` de `Shipments` (`LBS_IsInventoryPlaceholderOrig`, `tms_fg14/modulo2.vba:4702-4707`;
`LBS_RowEffectiveOrigen`, `tms_fg14/modulo2.vba:4743-4750`).

### 2. El grupo de consolidación

`LBS_ConsolidaGroup` (`tms_fg14/modulo2.vba:5444-5458`) traduce el destinatario de la columna
    60|`O` a un grupo, usando la hoja `Consolida`. Si el destinatario no está listado,
**el grupo es el destinatario mismo**:

```vba
' Destinatario no listado en Consolida: grupo propio (sin multistop accidental).
LBS_ConsolidaGroup = dest
```

Ese comportamiento por omisión es deliberado y es la decisión de diseño más importante de la
hoja `Consolida`: un destinatario nuevo nunca se une accidentalmente a un multi-stop
existente. Solo consolida con otros si alguien lo agrega a la hoja explícitamente.
    70|
Si el destinatario viene vacío, la función devuelve cadena vacía y la fila **no consolida con
nada**:

```vba
' Sin Destinatario: no inventar grupo por metro (evitar colision con nombres Consolida).
```

Los comentarios que encabezan la función registran un error concreto que se corrigió
(`tms_fg14/modulo2.vba:5439-5443`): agrupar por CEDIS mezclaba los destinatarios 7482, 7494 y
7464 —que no están en el grupo `WALMARTMX`— en un mismo camión, y el síntoma fue el folio
    80|`P-1095`.

### 3. La marca R/C de La Comer

Solo para La Comer, se agrega la marca de confirmación `R` o `C` leída de la columna `AE`
(`LBS_LaComerConfRC`, invocada en `tms_fg14/modulo2.vba:5470`). Es lo que separa los folios de
reposición de los de confirmación. Detalle en [cadenas/la-comer.md](cadenas/la-comer.md).

### 4. La vigencia (`R`)

La regla es dura: **nunca se mezclan vigencias distintas en un camión.**
    90|
`LBS_VigenciaKey` (`tms_fg14/modulo2.vba:2688-2703`) normaliza la columna `R` a `YYYYMMDD`.
Cuando no logra interpretar la celda devuelve `"NOVIG"`, y el comentario aclara la semántica:

```
' Sin fecha valida -> "NOVIG" (agrupan entre si, nunca con filas fechadas).
```

Es decir, las filas sin vigencia legible forman su propio grupo. No se cuelan en un camión
fechado.

   100|El parseo tiene tres caminos porque la columna `R` llega en formatos distintos según de dónde
salió (`LBS_TryParseVigenciaDate`, `tms_fg14/modulo2.vba:2640-2683`):

1. Fecha real de Excel, con `IsDate`.
2. Serial numérico de Excel, validando que esté entre 1 y 2 958 466 para descartar basura.
3. Texto `DD/MM/YYYY` o `D/M/YYYY`, parseado a mano.

El tercer camino existe por el locale. El comentario lo dice directo:
`' Mexico-safe d/m/yyyy or dd/mm/yyyy when IsDate fails (locale).` Con configuración regional
en inglés, `IsDate("31/08/2026")` devuelve falso y sin este respaldo la vigencia se perdía.
   110|
Y todo el bloque está envuelto en manejo de error con esta justificación
(`tms_fg14/modulo2.vba:2687`):

```
' On Error: col R a veces trae Error/Null/texto; nunca tumbar Optimizar por una celda.
```

### La llave alterna de Soriana y City Club

Hay una segunda llave, `LBS_ClubCityLaneKey` (`tms_fg14/modulo2.vba:5478-5487`), que es la
   120|misma **sin la vigencia**:

```
UCase(L) | grupo   (solo si el grupo empieza con "CLUBCITY")
```

Se usa en una segunda pasada de consolidación, tal como explica el comentario:

```
' ClubCity lane without vigencia: plant|CLUBCITY* (MTY/GDL/QRO/TUL).
' Pass-1 trucks still use LBS_ConsolidaKey (includes vigencia); pass-2 merges by lane.
```

   130|La primera pasada respeta la vigencia; la segunda une camiones del mismo carril aunque las
vigencias difieran. Es la única excepción autorizada a la regla dura de vigencia, y aplica
solo a Soriana y City Club. Detalle en [cadenas/soriana-city-club.md](cadenas/soriana-city-club.md).

## El cupo

El cupo es cuántas tarimas caben en el camión. Se compara contra `SUM(T) + SUM(W)` del folio,
que es lo que se escribe en la columna `Z`.

`LBS_TruckCapForRow` (`tms_fg14/modulo2.vba:5847-5899`) lo resuelve fila por fila, y su
estructura es una cascada de casos especiales. En orden de evaluación:

   140|### 1. Sencillo de OXXO, Neto o Mode Mix autoservicio

Si `Y` contiene `Sencillo` y la cadena es OXXO, Neto, Alsuper, Go Mart o Europea, el cupo se
resuelve contra el `Catalogo Mode Mix` (`LBS_SencilloCapForRow`, `tms_fg14/modulo2.vba:5841`).

Para OXXO hay además un ajuste por equipo (`LBS_ApplyOxxoSencilloEquipCap`,
`tms_fg14/modulo2.vba:5770-5780`), con este comentario:

```
' Map OXXO Sencillo catalog/lane pallets to equipment Max:
' MTY (400101621) -> Z5290_OXX_MTY 28; catalog <=22 (often 20) -> Z4290_OXX 22;
   150|' else Z5290_OXX 24 (never catalog S/26).
```

El paréntesis final es la clave: el catálogo dice 26 pero el equipo real de OXXO no da para
26, así que se fuerza a 24.

### 2. Full de mayorista de camión único

Si `Y` contiene `Full` y la cadena está en la lista blanca de mayoristas de camión único, el
cupo es siempre el de embarque completo, nunca el de caja
(`tms_fg14/modulo2.vba:5865-5869`):

   160|```
' Single truck: always shipment cap (35 Cap35 SKU / 36 other), never caja half.
```

Detalle en [cadenas/mayoristas-cap35.md](cadenas/mayoristas-cap35.md).

### 3. Full de OXXO, Neto, Soriana, City Club, Go Mart, Alsuper, Cap35, o folio con sufijo

Aquí está la bifurcación crítica (`tms_fg14/modulo2.vba:5879-5883`):

- Folio **con** sufijo `a`/`b`: cupo de caja, `LBS_FullBoxCapForRow` = catálogo / 2.
   170|- Folio **sin** sufijo: cupo de embarque completo, `LBS_CatalogFullShipmentCap`.

El comentario documenta un fallo que costó caro (`tms_fg14/modulo2.vba:5875-5878`):

```
' a/b caja: shipCap/2 (OXXO 18 / Soriana/GoMart 20). Pre-split: full shipment.
' ClubCity leftover box cupo is FolioAD/OpenTrucks only — NOT here.
' SummaryOK TrimSandwichW uses this; box on unsuffixed Full peels 40->20
' and hard-closes Excel on large Pedidos Surtidos.
```

   180|Aplicar el cupo de caja (20) a un Full sin partir (40) hacía que el motor intentara
descargar la mitad del camión, y en una hoja grande eso cerraba Excel.

### 4. Walmart

Walmart nunca usa el cupo métrico de 26, ni cuando su grupo de consolidación es un
destinatario suelto (`tms_fg14/modulo2.vba:5887-5897`):

```
' Walmart: always 28 (Sencillo) / 40 (Full) by chain — never metro 26 when
' Consolida group is a bare destinatario id (not "WALMART*").
   190|' Alsuper/GoMart/Europea use Mode Mix above (26 S / 40 F) — not Walmart 28.
```

### 5. Por omisión: el cupo del grupo

`LBS_TruckCapForGroup` (`tms_fg14/modulo2.vba:5497-5509`):

| Condición | Cupo |
|---|---|
| `Y` contiene `Full` | `40` (`LBS_FULL_TRUCK_CAP`) |
| El grupo empieza con `WALMART` | `28` (`LBS_WALMART_TRUCK_CAP`) |
   200|| Todo lo demás | `26` (`LBS_METRO_TRUCK_CAP`) |

### El cupo por folio ya existente

Cuando el motor quiere agregar carga a un camión que ya existe, usa una función distinta:
`LBS_TruckCapForFolioAD` (`tms_fg14/modulo2.vba:5905-5956`). Tiene tres excepciones propias:

- **COMEXTRA sin sufijo:** el cupo es el de embarque (40), para llenar hacia arriba.
  `' COMEXTRA unsuffixed Full leftover: FolioAD spare = shipCap 40 (fill toward 40 when possible)'`
- **Soriana y City Club sin sufijo:** si el camión ya trae `shipCap - 2` o más tarimas se le
  permite llegar al cupo de embarque; si no, se limita al de caja
   210|  (`' Near-complete shipment: allow spare up to shipCap before a/b split.'`).
- **Cualquier folio con sufijo:** cupo de caja, sin excepción.

### La tabla de constantes de cupo

| Constante | Valor | Comentario original |
|---|---|---|
| `LBS_METRO_TRUCK_CAP` | `26` | `' LBS - capacidad de camion (tarimas) para montar restos por metro (Soriana/City Club).` |
| `LBS_WALMART_TRUCK_CAP` | `28` | `' LBS - capacidad de camion (tarimas) para los grupos WALMART (Bodega Aurrera + SuperCenter).` |
| `LBS_FULL_TRUCK_CAP` | `40` | `' LBS - capacidad maxima por confirmacion COMEXTRA Full caja seca (catalogo mode mix = 40).` |
| `LBS_FULL_BOX_CAP` | `20` | `' LBS - caja (a o b) de un Full: ~20 tarimas por sufijo de camion (PartirFulles).` |
   220|| `LBS_COMEXTRA_SPLIT_CAP` | `20` | `' LBS - aviso blando: LBS suele armar splits a/b de ~20 tarimas por confirmacion.` |
| `LBS_OXXO_BOX_CAP` | `18` | `' LBS - fallback OXXO Full shipment / box when Catalogo Mode Mix miss.` |
| `LBS_OXXO_SHIPMENT_CAP` | `36` | (el mismo respaldo) |
| `LBS_OXXO_SENCILLO_CAP` | `24` | `' Z5290_OXX=24 (not catalog S/26 / generic Z5290), Z4290_OXX=22, Z5290_OXX_MTY=28.` |
| `LBS_OXXO_Z4290_SENCILLO_CAP` | `22` | (idem) |
| `LBS_OXXO_MTY_SENCILLO_CAP` | `28` | (idem) |
| `LBS_OXXO_MTY_DEST` | `"400101621"` | El destinatario de Monterrey |
| `LBS_CAP35_SHIPMENT` | `35` | `' SKU-level Full shipment cap (sheet "Cadenas 35 Tarimas": allowlisted cadena + SKU, Viajes a 35=YES).` |
| `LBS_MAYORISTA_FULL_SHIPMENT` | `36` | `' Allowlisted mayorista Full (non-listed SKU): single-truck ceiling 36 (no a/b).` |

   230|Todas en `tms_fg14/modulo2.vba:19-43`. Nótese la nota de fecha en el bloque de OXXO:
`' OXXO Sencillo from equipment Max Pallet Count (client 29/07/2026)`. Los cupos de OXXO no
salen del catálogo sino del `Max Pallet Count` del equipo, acordado con el cliente en esa
fecha.

### De dónde sale el cupo del catálogo

`LBS_EnsureCatalogCaps` (`tms_fg14/modulo2.vba:5606-5698`) lee la hoja `Catalogo Mode Mix` una
sola vez por sesión y arma cuatro diccionarios:

| Diccionario | Llave | Contenido |
   240||---|---|---|
| `mLBS_CatLaneCap` | `origenId\|destino` | Tarimas máximas Full por carril |
| `mLBS_CatDestCap` | `destino` | Tarimas máximas Full por destino |
| `mLBS_CatLaneCapS` | `origenId\|destino` | Lo mismo para Sencillo |
| `mLBS_CatDestCapS` | `destino` | Lo mismo para Sencillo |

Solo entran los renglones cuyo `mode mix` es `F` o `S` y cuyo máximo de tarimas es positivo
(`tms_fg14/modulo2.vba:5649-5656`).

Los diccionarios por destino tienen un mecanismo de desambiguación: si dos carriles del mismo
destino declaran cupos distintos, el valor se marca con `-1` y **el destino se descarta**
   250|(`tms_fg14/modulo2.vba:5669-5679`, `5684-5697`). El respaldo por destino solo se usa cuando
es inequívoco.

La prioridad de resolución para Sencillo está en el comentario
(`tms_fg14/modulo2.vba:5782-5783`):

```
' Catalog Mode Mix S Pallets Max for OXXO Sencillo, then clamp to *_OXX equipment.
' Cap priority: lane S / F-only / dest S -> OXXO equipment map (24 / 22 / MTY 28).
```

   260|El caso `F-only` es interesante: un Sencillo que viaja por un carril declarado únicamente
como Full. El código lo resuelve como `Z5290_OXX` y cita el carril donde apareció
(`' Sencillo on a Full-only catalog lane (e.g. LEON P-1254): treat as Z5290_OXX.'`,
`tms_fg14/modulo2.vba:5815`).

## Sándwich, camas y charolas

Estos tres términos describen cómo se apila la mercancía. Están en el
[glosario](../glosario.md), pero aquí está la mecánica.

### La unidad sándwich
   270|
Una **unidad sándwich** es una tarima física que carga producto de más de un SKU o de más de
un pedido. En la hoja se identifica por la pareja `AD` + `AI`:

- `AD` es el folio del camión.
- `AI` (`Craft`) contiene los `Unit Id` de la PCA que forman la unidad, separados por coma.

Cada unidad tiene exactamente un **ancla** y cero o más hermanas:

| `AG` | Rol |
|---|---|
   280|| `LBS_SANDWICH_ANCHOR` | El ancla. Lleva `W = 1`: paga el espacio de la tarima |
| `LBS_SANDWICH` | Hermana. Lleva `W = 0`: viaja en la tarima del ancla |

La regla de calificación como ancla (`LBS_RowQualifiesAsSandwichAnchor`,
`tms_fg14/modulo2.vba:6424-6428`) tiene el comentario más útil de la sección:

```
' Ancla sandwich solo si la fila trae restos (U>0). T>0 con U>0 es ancla mixta valida.
' T>0 con U=0 es tarima completa sin restos: no es ancla.
```

Una tarima llena no puede ser ancla porque no le queda espacio arriba.
   290|
Las cadenas que usan este conteo por unidad están en `LBS_IsSandwichWChain`
(`tms_fg14/modulo2.vba:6353-6358`): Walmart, Alsuper, Go Mart, Europea, Soriana, City Club y
COMEXTRA. La Comer queda **fuera** a propósito:

```
' LA COMER stays outside — FixLaComerSandwichAnchors owns its ANCHOR/W recovery.
```

Después de los merges, los renglones de una unidad pueden quedar dispersos. El agrupamiento
se hace sobre toda la hoja, no por filas contiguas, y el comentario explica por qué
   300|(`tms_fg14/modulo2.vba:6516-6519`):

```
' Agrupa folio+AI sobre TODA la hoja (no por filas contiguas): tras los merges una unidad
' puede quedar en fragmentos no adyacentes y cada fragmento recibia su propio ANCHOR.
```

El ancla se promueve a la primera posición del bloque con `LBS_PromoteSandwichAnchorRow`
(`tms_fg14/modulo2.vba:6376-6397`), que copia celda por celda hasta la columna 48 (`AV`).
El comentario del encabezado —`' celda a celda; sin arrays ni Cut/Insert'`— es otra defensa
contra el cierre de Excel.
   310|
La prioridad para elegir ancla cuando hay varios candidatos favorece la tarima base
(`LBS_SandwichAnchorTypePriority`, `tms_fg14/modulo2.vba:6366-6374`):

| Tipo de contenedor PCA | Prioridad |
|---|---|
| `PALLET` | 1 (mejor) |
| `SINGLE` | 2 |
| `SANDWICH` | 3 |
| `LAYER` | 4 |
| cualquier otro | 5 |
   320|
### Las camas compartidas

Una **cama** es un nivel horizontal de cajas dentro de la tarima. Dos SKU distintos pueden
compartir cama si sus cajas tienen prácticamente la misma altura.

El criterio es la tolerancia `LBS_WALMART_LAYER_ALTO_TOL_CM = 0.2`
(`tms_fg14/modulo2.vba:98-99`):

```
' Shared-cama match: Abs(AltoA-AltoB) <= 0.2 cm. Mixed cluster uses TI-1 cases/cama.
   330|' Applies to Armado-en-cama TI HI chains (LBS_ChainAllowsSharedCamas).
```

Dos milímetros de diferencia en la altura de la caja es el máximo. La segunda frase es una
penalización: cuando la cama es mixta, la capacidad baja de `TI` a `TI - 1` cajas, porque el
acomodo pierde una posición.

Las cadenas que permiten camas compartidas son catorce
(`LBS_ChainAllowsSharedCamas`, `tms_fg14/modulo2.vba:6342-6348`): Walmart, Soriana, City Club,
La Comer, HEB, Alsuper, Go Mart, Europea, Chedraui, Smart, Merco, Merza, Casa Ley y COMEXTRA.

   340|La compatibilidad de dos SKU se resuelve en `LBS_SkuLayerAltoCompatible`, que primero
verifica si son el mismo material (compatible por definición) y luego compara sus alturas
de TI HI.

### Las charolas y los restos

Un **resto** es el sobrante que no llena una tarima: la columna `U`. Cuando el producto viene
en charolas, ese resto puede apilarse sobre otra tarima, y de ahí el nombre de la columna
`U` en la plantilla: `Restos de charolas`.

El motor tiene funciones específicas para acomodarlos, todas con el patrón `Layer` en el
   350|nombre (`LBS_LayerRestosMetro` y familia). El principio es siempre el mismo: buscar una
tarima ancla con espacio y altura disponible, agregar el resto ahí, y dejar `W = 0` en la
fila del resto para que no pague tarima aparte.

## La altura: `AO` y `AP`

| Columna | Significado |
|---|---|
| `AO` | Altura en metros de **esta fila** dentro de la tarima |
| `AP` | Altura **total** de la unidad `AD` + `AI`, escrita en la última fila del grupo |

   360|`AP` es literalmente la suma de los `AO` de la unidad. `LBS_WalmartRebuildAPTouched`
(`tms_fg14/modulo2.vba:8942-8976`) recorre las unidades que cambiaron, acumula `AO` por
`AD & Chr(1) & AI` y escribe el total en la última fila de cada una.

### Cómo se calcula `AO`

Hay dos fuentes, y TI HI tiene prioridad. `LBS_WalmartFillAOFromPCA`
(`tms_fg14/modulo2.vba:9485-9517`) intenta en este orden:

1. **TI HI** (`LBS_WalmartTryApplyTihiAO`, `tms_fg14/modulo2.vba:8813-8837`).
2. **La altura del ancla en la PCA**, con las capas ya tomadas descontadas.
   370|3. **La altura por material en la PCA.**

El comentario justifica la prioridad (`tms_fg14/modulo2.vba:9500`):

```
' TI HI first for Walmart/La Comer sandwich/restos (PCA leftovers inflate partial layers).
```

LBS reporta la altura de la capa completa, así que un resto de media capa aparecía con la
altura de la capa entera. TI HI permite calcularla bien.

Solo se toca `AO` si la fila cumple cuatro condiciones (`tms_fg14/modulo2.vba:9486-9492`):
   380|`H = "Programado"`, la cadena está en el catálogo TI HI, `U > 0`, y `AO` está en cero.
La última condición hace la operación idempotente: no se recalcula lo ya calculado.

### La fórmula de TI HI

`LBS_WalmartTihiFormulaHeightM` es directa:

```
capas = techo(U / TI)
AO = capas * (Alto / 100)
```

   390|El techo se calcula sin `WorksheetFunction` (`LBS_WalmartTihiCeilLayers`):
`Int((uQty + ti - 0.0000001) / ti)`, con el épsilon para evitar que un cociente exacto suba
una capa por error de punto flotante.

Hay una segunda fórmula, la *source of truth*, que escala contra el armado completo
(`LBS_WalmartTihiSotHeightM`, `tms_fg14/modulo2.vba:8800-8810`):

```
AO_sot = (capas / HI) * (AlturaArmado / 100)
```

   400|Cuando las dos difieren por más de `LBS_WALMART_TIHI_BL_TOL_M = 0.005` m (5 mm), la
diferencia se anota en la columna **`BL`** como marca de diagnóstico
(`tms_fg14/modulo2.vba:8829-8833`). `BL` no está en el rango `A:AV` documentado: es una
columna de trabajo que solo aparece cuando hay discrepancia entre las dos fórmulas.

### El tope de altura

| Constante | Valor | Comentario original |
|---|---|---|
| `LBS_WALMART_MAX_HEIGHT_M` | `1.6` | `' LBS - WALMART: altura maxima de tarima/unidad (suma AO -> AP) en metros.` |
| `LBS_LACOMER_MAX_HEIGHT_M` | `1.6` | `' LBS - LA COMER: altura maxima de tarima/unidad (suma AO -> AP), trial = Walmart 1.6 m.` |
   410|
Diez cadenas aplican el tope de 1.6 m (`LBS_ChainEnforcesUnitHeight`,
`tms_fg14/modulo2.vba:11895-11900`): Walmart, La Comer, Soriana, City Club, HEB, Alsuper,
Chedraui, Go Mart, Europea y COMEXTRA.

Hay que distinguir dos cosas que el comentario separa explícitamente
(`tms_fg14/modulo2.vba:11893-11894`):

```
' True si la cadena aplica tope 1.6 m (discard/repack/Fill densify). AO populate is separate
' (LBS_ChainComputesTihiAO = all TI HI cadenas). Cap family = mix autoservicio only.
   420|```

Calcular `AO` (todas las cadenas del catálogo TI HI) y **hacer valer** el tope (solo esas
diez) son operaciones distintas. Una cadena puede tener alturas calculadas sin que el tope se
le aplique.

### La excepción de armado alto

Hay una excepción del cliente para SKU cuyo armado de fábrica ya rebasa 1.6 m
(`LBS_UnitHeightCapM`, `tms_fg14/modulo2.vba:11939-11943`):

   430|```
' Unit height cap (m). Default = chain max (1.6). Homogeneous SKU (Soriana/City Club/
' HEB/Chedraui) with TI HI Altura Armado > chain max uses Max(Altura Armado, Alto*HI)
' (client exception for armado already over 1.6 m). Alto*HI covers the ~2mm
' TIHI formula vs Altura Armado gap (e.g. Coronita 1.672 vs 1.67). Mixed units
' stay at chain max.
```

Tres condiciones para que aplique:

1. La cadena es Soriana, City Club, HEB o Chedraui.
   440|2. La unidad es **homogénea**: un solo material.
3. La `Altura Armado` de TI HI para ese material supera 1.6 m.

En ese caso el tope sube a `Max(Altura Armado, Alto * HI)`. El ejemplo del comentario es
Coronita: la fórmula da 1.672 m y el catálogo dice 1.67 m, y tomar el máximo evita que una
diferencia de 2 mm marque la tarima como excedida.

Las unidades mixtas nunca reciben la excepción.

### La bandera de altura

   450|Cuando `AP` supera el tope, se agrega a `AV` (`tms_fg14/modulo2.vba:9388-9392`):

```
REVISION MANUAL: altura ><tope> ...
```

La comparación lleva un margen de 0.0005 m para no marcar diferencias de redondeo.

## El peso: `AT`, `AU` y la tara

### `AT` — el peso de la fila
   460|
Nace del `Weight` de la PCA (`tms_fg14/modulo2.vba:1162`). Pero seis cadenas lo **recalculan**
desde el peso bruto por caja de TI HI (`LBS_ChainAppliesTihiAT`,
`tms_fg14/modulo2.vba:6336-6339`): Walmart, Soriana, City Club, Alsuper, Go Mart y Europea.

Y lo calculan con fórmulas distintas (`tms_fg14/modulo2.vba:8869-8872`):

```
' Refresh AT from TI-HI Peso Bruto (kg/case).
' Walmart: AT = (T*Armado + U) * pesoCase. If T>0 and Armado blank, use TI*HI.
' Soriana/City Club: AT = S * pesoCase (cartonaje) so Cont1/_S and sandwich sisters
   470|' get proportional kg; fallback (T*Armado)+U when S blank. No hit / no peso -> leave PCA AT.
```

La diferencia importa. Walmart reconstruye las cajas desde tarimas y restos; Soriana y City
Club usan el cartonaje directo de `S`, y el comentario da la razón: así las filas `Cont1`,
las `_S` y las hermanas de sándwich reciben kilos proporcionales en lugar de que una sola
cargue con todo el peso de la unidad.

Cuando no hay coincidencia en TI HI o no hay peso por caja, se deja el valor de la PCA.

### `AU` — el peso del camión
   480|
```
AU = SUM(AT del folio) + LBS_PesoTarimaKg() * Z
```

(`tms_fg14/modulo2.vba:1648-1651`, donde `Z` es el total de tarimas del folio.)

La tara por tarima es 30 kg para los tres tipos, y el código lo marca como pendiente
(`tms_fg14/modulo2.vba:13-17`):

```
   490|' Tarima tare (kg). All 30 until tech specs; LBS_PesoTarimaKg picks by type.
Private Const SK_PESO_TARIMA_DEFAULT As Double = 30#
Private Const SK_PESO_TARIMA_CHEP As Double = 30#      ' TODO tech specs
Private Const SK_PESO_TARIMA_PLASTICA As Double = 30#  ' TODO tech specs
Private Const SK_PESO_TARIMA_MADERA As Double = 30#    ' TODO tech specs
```

La infraestructura para diferenciar por tipo ya existe (`LBS_PesoTarimaKg` distingue `CHEP`,
`PLASTIC` y `MADER`, `tms_fg14/modulo2.vba:11490-11502`), pero los tres valores son iguales
hasta que lleguen las especificaciones técnicas.

   500|### El techo de peso

| Constante | Valor | Aplica a |
|---|---|---|
| `SK_MAX_PESO_KG` | `29000` | Sencillo y el resto de las cadenas |
| `LBS_FULL_MAX_PESO_KG` | `52500` | Full de Alsuper, Go Mart y Europea |

El comentario del segundo lo ubica: `' Catalogo Mode Mix Full weight ceiling (52.5 t) for
Alsuper/Go Mart Full lanes.'` (`tms_fg14/modulo2.vba:11-12`).

   510|`LBS_MaxPesoKgForRow` (`tms_fg14/modulo2.vba:11506-11520`) elige entre los dos: 52.5 t solo
si `Y` contiene `Full` **y** la cadena es de Mode Mix autoservicio. Todo lo demás va a 29 t.

Al exceder el techo se agrega a `AV` (`LBS_PesoReviewFlagText`,
`tms_fg14/modulo2.vba:11522-11531`):

```
REVISION MANUAL: peso >29 ton (30.1 ton)
REVISION MANUAL: peso >52.5 ton (54.2 ton)
```

   520|Además se registra en la hoja de fallos con el texto
`REVISION MANUAL: peso supera 29 ton` (`tms_fg14/modulo2.vba:1716`) y se guarda en
`weightReviewLog`, un registro por folio (`tms_fg14/modulo2.vba:1663`).

Cinco familias de cadena permiten **descargar** tarimas cuando el camión pasa de peso
(`LBS_IsPesoSalvageChain`, `tms_fg14/modulo2.vba:6312-6318`): Walmart, la familia
`CLUBCITY` (Soriana + City Club), Alsuper, Go Mart y Europea. El comentario lo llama
*overweight peel*. Para las demás cadenas, el exceso de peso solo genera la bandera; hay que
resolverlo a mano.

## Cómo se cuenta el espacio ocupado
   530|
Tres funciones que parecen iguales pero no lo son:

| Función | Cita | Qué devuelve |
|---|---|---|
| `LBS_RowTarimas` | `tms_fg14/modulo2.vba:17562` | Tarimas de una fila |
| `LBS_RowSumTW` | `tms_fg14/modulo2.vba:5964` | `T + W` de la fila |
| `LBS_GroupEffectiveTarimas` | `tms_fg14/modulo2.vba:16261` | Tarimas efectivas de un conjunto de filas, sin contar dos veces las unidades sándwich |

El comentario de `LBS_RowSumTW` recuerda la convención
(`tms_fg14/modulo2.vba:5963`):

   540|```
' Col Z por fila = SUM(T)+SUM(W). LBS_SANDWICH lleva T=0 (tarima en W).
```

`LBS_GroupEffectiveTarimas` es la que hay que usar al comparar contra el cupo, porque una
unidad sándwich con cinco filas ocupa una tarima, no cinco.

## Los patrones defensivos del código

Al leer el módulo se repiten tres advertencias. No son estilo: cada una corresponde a un
cierre inesperado de Excel que hubo que diagnosticar.
   550|
### 1. Nunca leer un rango grande de un tirón

```
' Filas en PCA (G preferred, A fallback). Never load A1:Z in one shot — hard-closes Excel.
```

Todos los volcados van en trozos de 1 000 filas.

### 2. Nunca `For Each` sobre `Collection` en los ciclos calientes

   560|```
' Indexed loops — For Each on Collections from Repack densify hard-closes Excel.
' Indexed — called from Repack densify; For Each on row Collections hard-closes Excel.
```

Aparece dos veces, en `LBS_WalmartConfCompatibleRows` (`tms_fg14/modulo2.vba:9520`) y en
`LBS_UnitHeightCapM` (`tms_fg14/modulo2.vba:11973`). Los recorridos usan índice explícito.

### 3. Cachear las hojas de parámetros

```
   570|' EFICIENCIA POR CADENA lookup cache. LBS_WalmartMinTarimas used to open/scan the
' sheet on every call (layer/pack/enforce) and hard-close Excel.
```

(`tms_fg14/modulo2.vba:139-140`.) Lo mismo para TI HI, `Catalogo Mode Mix`, `Consolida`,
`Cadenas 35 Tarimas` y el mapa de plantas. Todas se leen una vez por sesión.

La consecuencia operativa: **modificar una hoja de parámetros con el libro abierto no surte
efecto hasta cerrar y volver a abrir**, salvo que se llame explícitamente al reset
correspondiente (`LBS_ResetConsMap` para `Consolida`, `tms_fg14/modulo2.vba:5256`).

   580|Hay una defensa adicional contra hojas con rango usado contaminado
(`tms_fg14/modulo2.vba:2600-2601`):

```
' Used-range pollution in col A used to scan ~1M rows per lookup and kill Excel.
If lastRowEfic > 500 Then lastRowEfic = 500
```

`EFICIENCIA POR CADENA` se lee como máximo hasta la fila 500. **Una cadena listada más abajo
de la fila 500 se ignora sin aviso.**
