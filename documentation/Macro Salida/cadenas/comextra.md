[Volver al índice de cadenas](README.md)

# Comextra

Comextra tiene una regla que ninguna otra cadena tiene: **una fila programada por
confirmación, pedido y SKU.** El motor colapsa todo lo demás. Junto con eso, cada
confirmación es su propio camión (no se unen las cajas `a` y `b` del mismo pedido) y el piso
de llenado es del 90 %, empatado con OXXO como el más exigente.

## 1. Identificación

`LBS_IsComextraChain` (`tms_fg14/modulo2.vba:11885`), comparación exacta:

```
LBS_IsComextraChain = (Replace(UCase$(Trim$(CStr(m))), " ", "") = "COMEXTRA")
```

Sin familia en `LBS_ChainFamily`. Sí participa en dos listas transversales por mención
explícita: `LBS_IsOpenTruckChain` (`tms_fg14/modulo2.vba:6321`) y
`LBS_IsTruckMinFillChain` (`tms_fg14/modulo2.vba:4434`).

## 2. Parámetros

| Parámetro | Valor | Origen |
|---|---|---|
| Cupo Full | 40 tarimas, tope duro | `LBS_FULL_TRUCK_CAP` (`modulo2.vba:25`) |
| Cupo caja `a`/`b` | 20 tarimas | `LBS_FULL_BOX_CAP` (`modulo2.vba:29`) |
| Aviso de división | 20 tarimas por confirmación | `LBS_COMEXTRA_SPLIT_CAP` (`modulo2.vba:27`) |
| Cupo sencillo | 26 tarimas | `LBS_METRO_TRUCK_CAP` (`modulo2.vba:19`) |
| Piso de llenado | Full: 90 % del cupo (36 de 40 / 18 de 20). Sencillo: 90 % de 29 t (~26.1 t) | `LBS_COMEXTRA_MIN_FILL` / `LBS_SENCILLO_MIN_PESO_FRAC` |
| Altura máxima de unidad | 1.60 m | `LBS_ChainMaxUnitHeightM` (`modulo2.vba:11908-11912`) |
| Peso máximo Sencillo | 29 t | `SK_MAX_PESO_KG` (`modulo2.vba:10`) |
| Peso máximo Full | 52.5 t | `LBS_FULL_MAX_PESO_KG`; fallback si el catálogo no tiene fila F |

Comentarios originales:

```
' LBS - capacidad maxima por confirmacion COMEXTRA Full caja seca (catalogo mode mix = 40).
Private Const LBS_FULL_TRUCK_CAP As Long = 40
' LBS - aviso blando: LBS suele armar splits a/b de ~20 tarimas por confirmacion.
Private Const LBS_COMEXTRA_SPLIT_CAP As Long = 20
```

```
' LBS - COMEXTRA: piso de llenado post-consolidacion (90% del cap).
' Full shipment / unsuffixed leftover: shipCap 40 -> piso 36. Full a/b caja: 20 -> piso 18.
' Sencillo: min-fill is 90% of 29 t (~26.1 t), not tarimas.
```

Full sigue midiendo el piso en tarimas. Sencillo abre o se queda Programado si el peso
llega a 90 % de 29 t. Comextra prefiere camiones llenos a camiones repartidos
(`Fill toward 40 when possible` en Full).

Un Full **sin** sufijo `a`/`b` es su propio camión, aunque exista `P-####a`/`b` del mismo
número. El recorte a 40 puede dejar un resto de 1 tarima en el AD base (`P-1428` junto a
`P-1428a`/`b`). `LBS_EnforceWalmartMinFill` lo agrupa con `|loose` para no sumarlo al
embarque a+b; si queda bajo 36, pasa a `No planeado`.

El tope de 40 es duro, con su propia verificación
(`tms_fg14/modulo2.vba:5749-5751`):

```
' COMEXTRA Full: hard max 40 (Mode Mix / lane catalog must not exceed equipment cupo).
If LBS_IsComextraChain(mVal) Then
    If cap > LBS_FULL_TRUCK_CAP Or cap < 1 Then cap = LBS_FULL_TRUCK_CAP
```

Igual que en OXXO, el catálogo Mode Mix no puede subir el cupo por encima del límite del
equipo.

El peso sigue la lane de este origen|dest, igual que OXXO
(`LBS_ComextraLaneMode` / `PF_ComextraLaneMode`). Torreon PC29 F es 52.5 t aunque
LBS haya escrito `Y=Sencillo`. Torreon PC01 S es 28.9 t. Un dest-level Full de
otra planta no convierte un Sencillo. Si la lane no existe (KALTEX) y `Y` tiene
Full, el fallback sigue siendo 52.5 t (`LBS_FULL_MAX_PESO_KG`). El peel de
peso suma `a`+`b` como un embarque (igual que AU); el cupo sigue midiendo cada confirmación
por separado.

`LBS_COMEXTRA_SPLIT_CAP` se describe como "aviso blando" y produce la marca
`REVISION MANUAL: tarimas >20` (`tms_fg14/modulo2.vba:16284`): no bloquea nada, solo señala
que la confirmación es más grande de lo que LBS suele armar.

## 3. Reglas de negocio

**Una fila por confirmación, pedido y SKU.** Es la regla central. La función
`LBS_ConsolidarComextraSkuPerTruck` es compartida: en todas las cadenas colapsa
tarimas completas (`T>0`, `U=0`) del mismo `AD|pedido|SKU`. Comextra además une
restos y anclas de sándwich del mismo SKU. El comentario
(`tms_fg14/modulo2.vba`):

```
' One Programado row per confirmacion (AD) + pedido + SKU.
' No planeado discards are ConsolidarNoPlaneados (origen|dest|pedido|SKU), not this fn.
```

Los detalles que importan:

- El cartonaje se colapsa y las tarimas (`T`) y unidades (`U`) se recalculan desde las cajas
  (`S`) y el cartonaje (`V`). No se suman las columnas de las filas originales.
- La columna `W` se queda con el **máximo** de las filas unidas, nunca la suma. Sumarla
  contaría el mismo espacio de piso varias veces.
- **No** se unen las cajas `a` y `b`. Los discards `No planeado` del mismo
  origen|dest|pedido|SKU los une `ConsolidarNoPlaneados`, sin mezclarse con Programado.
- Solo corre después de `PartirTarimasFULL`. Correrlo antes uniría filas que la división
  todavía va a separar.

**Una confirmación por camión.** El comentario de `LBS_ConsolidaGroup`
(`tms_fg14/modulo2.vba:5443`):

```
' COMEXTRA: una confirmacion (AD) por camion; no unir splits a/b del mismo pedido.
```

Y en `SummaryFallo` (`tms_fg14/modulo3.vba:836`):

```
' COMEXTRA: cada confirmacion AD es su camion; cap 40/26 por AD, no merge entre a/b.
```

Es lo contrario del multi-stop: en Comextra el folio define el camión de forma rígida.

**Cupo de charolas.** Cada charola con más de una unidad cuenta como un espacio de piso. El
comentario de `LBS_ComextraCharolaCupoSlots` (`tms_fg14/modulo2.vba:11736-11738`):

```
' COMEXTRA OpenTrucks/Fill seed: each charola U>1 starts as own AI/W=1 cupo slot.
' Fallos then runs Walmart/Soriana densify (RepackRestosSandwiches + MergeThin) under
' 1.6 m so thin AIs collapse into real sandwiches (one W=1 anchor). U=1 rides along.
```

El proceso tiene dos etapas: primero cada charola se cuenta como espacio propio, de forma
conservadora, y después la densificación de `SummaryFallo` (reutilizando la maquinaria de
Walmart y Soriana) las colapsa en sándwiches reales bajo el límite de 1.60 m. Las charolas de
una sola unidad viajan con la que las hospede.

**Relleno desde `No planeado`.** `LBS_ComextraFillSpareFromNP`
(`tms_fg14/modulo2.vba:18505`) llena el espacio libre de los camiones con filas descartadas.
Es la consecuencia práctica del `Fill toward 40 when possible`.

**Apertura de camiones nuevos.** Comextra está en `LBS_IsOpenTruckChain` por mención
explícita, junto con OXXO. `LBS_IsOpenTruckGapAG` acepta cualquier `AG` (incluido en
blanco) para ambas cadenas.

`LBS_OpenTruckInventTemplateAD` inventa el folio plantilla cuando el destino solo tiene
`No planeado`. El alias `LBS_ComextraInventTemplateAD` delega ahí. El cupo del Sencillo
inventado sale de `LBS_SencilloCapForRow` (p. ej. Torreón 20), no del metro 26. Un pack
Sencillo que no llega a 26.1 t sin pasarse de 29 t se queda `No planeado`.

**Un SKU por camión, con degradación.** `LBS_ComextraDemoteOneCartonTarimas`
(`tms_fg14/modulo2.vba:11825`) degrada tarimas de un solo cartón, y
`LBS_ComextraHostAIOnAD` (`tms_fg14/modulo2.vba:11801`) /
`LBS_ComextraOtherHostAIOnAD` (`tms_fg14/modulo2.vba:11857`) buscan la unidad anfitriona donde colocar una charola.

**Altura, camas compartidas y sándwich.** Comextra está en las tres listas:
`LBS_ChainEnforcesUnitHeight` (`tms_fg14/modulo2.vba:11895`),
`LBS_ChainAllowsSharedCamas` (`tms_fg14/modulo2.vba:6342`) y
`LBS_IsSandwichWChain` (`tms_fg14/modulo2.vba:6353`). Pero **no** en
`LBS_ChainAppliesTihiAT` (`tms_fg14/modulo2.vba:6336`): el peso no se refresca desde el
catálogo.

## 4. Procedimientos

| Procedimiento | Línea | Qué hace |
|---|---|---|
| `LBS_IsComextraChain` | `11885` | Reconoce `COMEXTRA` |
| `LBS_ComextraCharolaCupoSlots` | `11739` | Cuenta el espacio de piso de una charola |
| `LBS_TagComextraCharolaEachAsUnit` | `11752` | Marca cada charola como unidad propia |
| `LBS_ComextraHostAIOnAD` | `11801` | Busca la unidad anfitriona en el folio |
| `LBS_ComextraDemoteOneCartonTarimas` | `11825` | Degrada tarimas de un solo cartón |
| `LBS_ComextraOtherHostAIOnAD` | `11857` | Busca una unidad anfitriona alterna |
| `LBS_ApplyComextraTarimaCapFlags` | `16278` | Escribe las marcas de `REVISION MANUAL` por cupo |
| `LBS_ConsolidarComextraSkuPerTruck` | `20040` | Colapsa a una fila por confirmación, pedido y SKU (todas las cadenas: solo tarimas completas; Comextra también restos) |
| `LBS_ComextraFillSpareFromNP` | `18505` | Rellena el espacio libre desde `No planeado` |
| `LBS_ComextraInventTemplateAD` | `20380` | Inventa el folio plantilla para un camión nuevo |
| `LBS_IsOpenTruckGapAG` | `6327` | Elegibilidad de remonte, con la excepción de `AG` vacía |

La fase de `SummaryOptimizar` correspondiente es `consolidar_comextra_sku`
(`tms_fg14/modulo2.vba:2537`).

## 5. Cómo validarlo

Fixtures en `tms_fg14/comextra/`:

| Archivo | Contenido |
|---|---|
| `sample.tsv` | Muestra de `Pedidos Surtidos` |
| `plan.tsv` | Plan de prueba |

Script:

| Script | Qué comprueba |
|---|---|
| `scripts/validate_comextra_merge.py` | La consolidación a una fila por confirmación, pedido y SKU; el contrato VBA de all-chains (`T>0 U=0`) |
| `scripts/validate_peso_all_chains.py` | WARN si el sample sigue partiendo tarimas completas del mismo SKU (p. ej. `P-1273`) |
| `scripts/validate_cartonaje_conservation.py` | Plan vs sample (bajo y sobre); llaves 156-gap COMEXTRA/GO MART |

## 6. Problemas conocidos y síntomas

| Síntoma | Causa probable | Dónde mirar |
|---|---|---|
| Varias filas del mismo SKU en la misma confirmación | La consolidación no corrió, o corrió antes de `PartirTarimasFULL` | `LBS_ConsolidarComextraSkuPerTruck` solo debe correr después de partir |
| Columna `W` inflada | Se sumó en lugar de tomar el máximo | El comentario dice `never the sum`. Si aparece, es una regresión |
| `REVISION MANUAL: tarimas >20` | Aviso blando: la confirmación es más grande de lo que LBS suele armar | No bloquea nada. `LBS_COMEXTRA_SPLIT_CAP` |
| `Descartado por baja eficiencia` con 33 o 34 tarimas | Bajo el piso de 36, que es el 90 % de 40 | `LBS_COMEXTRA_MIN_FILL`. Es el piso más exigente junto con OXXO |
| Cajas `a` y `b` unidas en un camión | No debería pasar: cada confirmación es su camión | Comentarios de `LBS_ConsolidaGroup` y `SummaryFallo` |
| Sobrantes de cupo que no se remontan | La columna `AG` quedó vacía **y** la excepción de Comextra no aplicó | `LBS_IsOpenTruckGapAG`; verificar que `Pedidos Surtidos!M` diga `Comextra` |
| Charolas contadas de más | La densificación de `SummaryFallo` no corrió | Es el diseño en dos etapas: primero conservador, después denso. Correr `SummaryFallo` |
| Camiones a 25 tarimas cuando cabían 40 | El relleno desde `No planeado` no encontró qué subir | `LBS_ComextraFillSpareFromNP` |
| Full de más de 40 | El catálogo Mode Mix declara un cupo mayor | El código lo topa. Si aparece, revisar el tope de `modulo2.vba:5749-5751` |
| Leftover `U` (S=58) y pallet id `214_Tarima` sin la tarima de 156 | El merge o el peel de altura tiró `T*V` y dejó solo restos | `LBS_ConsolidarComextraSkuPerTruck` reencola el excedente a NP; `LBS_RestorePlanCartonajeGaps` corre al final. `CompararCartonajes` al inicio no cierra un hueco que Fallos abre después |
| Full `P-1428` de 1 tarima Programado junto a `P-1428a`/`b` | El min-fill sumaba el resto al embarque a+b (ya en 40) y lo daba por lleno | `|loose` en `LBS_EnforceWalmartMinFill`; `Fallos: min-fill after RecalcPeso` |
| Full AW 29000 / AV `peso >29 ton` con Y Full | Catálogo sin fila F (KALTEX) caía al Sencillo 29 t | `LBS_MaxPesoKgForRow` / `PF_MaxPesoKgForRow`: COMEXTRA Full → 52.5 t |
| Full `a`+`b` Programado sobre 52.5 t | El peel medía cada caja (~27 t) y AU el embarque | `LBS_WalmartDiscardOverWeightTruck` usa `LBS_TarimaTotalGroupKey` |
