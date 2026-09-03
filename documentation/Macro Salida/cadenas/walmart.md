[Volver al índice de cadenas](README.md)

# Walmart (Bodega Aurrera y SuperCenter)

Es la cadena con más lógica propia del motor: cerca de 100 procedimientos con `Walmart` en el
nombre. Casi todo gira alrededor de dos cosas que ninguna otra cadena combina con la misma
intensidad: el **sándwich** (apilar tarimas parciales en una sola unidad física) y el **tope
de altura de 1.60 m** que limita cuánto se puede apilar.

## 1. Identificación

`LBS_IsWalmartChain` (`tms_fg14/modulo2.vba:6179`) delega en la familia:

```
LBS_IsWalmartChain = (LBS_ChainFamily(CStr(m)) = "WALMART")
```

Y la familia (`LBS_ChainFamily`, `tms_fg14/modulo2.vba:5203-5204`) reconoce exactamente dos
literales, normalizados a mayúsculas y sin espacios:

| Valor en `Pedidos Surtidos!M` | Reconocido |
|---|---|
| `Walmart BA` → `WALMARTBA` | Sí |
| `Walmart SC` → `WALMARTSC` | Sí |
| `Walmart` a secas | **No** |
| `WalMart BA` | Sí (se normaliza a mayúsculas) |

Para el catálogo TI HI hay un clasificador aparte,
`LBS_WalmartTihiIsWalmartCadena` (`tms_fg14/modulo2.vba:8047`), que acepta los mismos dos
literales. No se usa para filtrar la carga del catálogo —eso cambió para incluir todas las
cadenas— pero sigue ahí para las estadísticas que reporta `LBS_SyncTihiSheet`.

Los grupos de consolidación de Walmart en la tabla embebida
(`LBS_LoadConsDefaults`, `tms_fg14/modulo2.vba:5354`) son ocho:

`WALMARTCHALCO`, `WALMARTCULIACAN`, `WALMARTGDL`, `WALMARTMERIDA`, `WALMARTMTY`,
`WALMARTMX`, `WALMARTNORTE`, `WALMARTVILLAHERMOSA`.

## 2. Parámetros

| Parámetro | Valor | Origen |
|---|---|---|
| Cupo sencillo | 28 tarimas | `LBS_WALMART_TRUCK_CAP` (`modulo2.vba:21`) |
| Cupo Full | 40 tarimas | `LBS_FULL_TRUCK_CAP` (`modulo2.vba:25`) |
| Cupo caja `a`/`b` | 20 tarimas | `LBS_FULL_BOX_CAP` (`modulo2.vba:29`) |
| Piso de llenado | 40 % del cupo = 12 de 28 | `LBS_WALMART_MIN_FILL` (`modulo2.vba:65`) |
| Rescate de grupo | 50 % del cupo = 14 de 28 | `LBS_WALMART_RESCUE_MIN_FILL` (`modulo2.vba:60`) |
| Altura máxima de unidad | 1.60 m | `LBS_WALMART_MAX_HEIGHT_M` (`modulo2.vba:23`) |
| Peso máximo | 29 t | `SK_MAX_PESO_KG` (`modulo2.vba:10`) |
| Tolerancia de cama compartida | 0.2 cm | `LBS_WALMART_LAYER_ALTO_TOL_CM` (`modulo2.vba:99`) |

Comentarios originales, que traen la justificación:

```
' LBS - capacidad de camion (tarimas) para los grupos WALMART (Bodega Aurrera + SuperCenter).
Private Const LBS_WALMART_TRUCK_CAP As Long = 28
' LBS - WALMART: altura maxima de tarima/unidad (suma AO -> AP) en metros.
Private Const LBS_WALMART_MAX_HEIGHT_M As Double = 1.6
```

```
' LBS - WALMART: piso de llenado de camion (tarimas). El % de "EFICIENCIA POR CADENA" es
' fill rate de PEDIDO (col AR, gate de eficiencia) para Walmart/Alsuper — truck floor is
' this constant (cap 28 -> piso 12). ClubCity (Soriana/City Club) truck floor is
' LBS_CLUBCITY_MIN_FILL (70% of metro cap 26 -> piso 19).
```

```
' LBS - WALMART: rescate de grupo todo-fallido si la carga combinada >= este % del cap (28 tarimas).
' Solo para el gate de grupo (LBS_WalmartGroupEfficiencyGate); el piso post-consolidacion es
' LBS_WALMART_MIN_FILL.
```

**Dos controles, dos números.** El porcentaje de `EFICIENCIA POR CADENA` se compara contra
`AR` (fill rate del pedido). El 40 % de `LBS_WALMART_MIN_FILL` se compara contra las tarimas
físicas del camión. Un camión con 15 tarimas pasa el piso aunque su pedido tenga fill rate
bajo, y un pedido con buen fill rate puede terminar en un camión de 8 tarimas que se
descarta. Es la confusión más frecuente en operación.

## 3. Reglas de negocio

**Cupo por cadena, no por grupo.** Un camión de Walmart lleva 28 tarimas siempre, aunque el
destinatario no esté mapeado en la hoja `Consolida`. Sin esta regla el motor usaría el cupo
genérico de 26 y dejaría dos tarimas sin usar en cada camión
(`LBS_TruckCapForRow`, `tms_fg14/modulo2.vba:5887-5896`).

**Sándwich.** Una unidad física puede llevar varias tarimas parciales apiladas: una tarima
base y capas encima. La columna `AI` identifica la unidad y la columna `W` cuenta el slot: la
tarima ancla lleva `W=1` y las capas `W=0`, para que no se cuente el mismo espacio de piso
dos veces. `LBS_IsSandwichWChain` (`modulo2.vba:6353`) incluye a Walmart.

**Tope de 1.60 m.** La suma de alturas `AO` de una unidad se acumula en `AP`, y si pasa de
1.60 m la unidad se rompe o se descarta la capa que sobra. Es un límite físico de la caja del
camión, no una preferencia comercial.

**Camas compartidas.** Dos SKU distintos pueden ir en la misma cama si su altura de caja
difiere en 2 mm o menos. En un clúster mezclado se usan `TI-1` cajas por cama en lugar de
`TI`, para dejar holgura (`modulo2.vba:97-98`).

**Una confirmación por unidad.** Dos confirmaciones distintas no deben quedar sobre la misma
tarima física. `LBS_WalmartSeparateConfPerUnit` (`modulo2.vba:7752`) las separa, y cuando no
puede deja la marca `REVISION MANUAL: conf 33/20 en misma tarima AI (cap 28)`
(`modulo2.vba:7870`).

**Gate por grupo, no por pedido.** A diferencia de Chedraui, Walmart evalúa la eficiencia por
grupo de consolidación completo, con promedio ponderado por tarimas. El comentario describe
el rescate multi-camión (`modulo2.vba:2898-2902`):

```
' WALMART BA/SC (y ALSUPER via chainKind): eficiencia por grupo de consolidacion antes del
' descarte folio-a-folio. Reglas: promedio ponderado por tarimas; si ningun folio pasa, rescate
' multi-camion: arma camiones folio a folio (mayor primero, con donacion parcial de tarimas
' completas) y descarta solo los camiones bajo LBS_WALMART_RESCUE_MIN_FILL del cap; merge de
' folios fallidos sobre folios que pasan cuando hay cupo.
```

Es decir: si ningún folio del grupo pasa por sí solo, el motor no tira el grupo. Arma
camiones empezando por el folio más grande, dona tarimas completas de los folios chicos, y
solo descarta los camiones que quedan por debajo del 50 % del cupo. Después intenta subir los
folios que reprobaron a los camiones que pasaron y tienen espacio.

**Canibalización y reconstrucción.** Cuando un camión queda sobre el cupo,
`LBS_WalmartCannibalizeToCap` (`modulo2.vba:3511`) le quita tarimas. Cuando un camión queda
bajo el piso y se descarta, `LBS_WalmartRebuildPalletsFromBajoLlenadoDiscards`
(`modulo2.vba:19574`) intenta reconstruir tarimas completas con las piezas descartadas, para
no perder producto por un armado desafortunado.

**Apertura de camiones nuevos.** Si `CompararCartonajes` encontró cartonaje del Plan que LBS
no reportó, `LBS_WalmartOpenTrucksForNoReportadoGaps` (`modulo2.vba:20449`) abre folios
nuevos para colocarlo, en lugar de dejarlo en `No planeado`.

**Marcadores de exportación.** `LBS_ApplyWalmartSandwichExportMarkers` (`modulo2.vba:6520`)
escribe las marcas que necesita el archivo de TMS para que el sándwich se entienda del otro
lado.

## 4. Procedimientos

Son cerca de 100. Agrupados por función, los que importan:

### Gate y empaque

| Procedimiento | Línea | Qué hace |
|---|---|---|
| `LBS_WalmartGroupEfficiencyGate` | `2903` | Gate de eficiencia por grupo, con rescate multi-camión. También sirve a Alsuper vía `chainKind` |
| `LBS_PackWalmartOneCkey` | `3184` | Empaca una llave de consolidación en camiones |
| `LBS_PackWalmartToCapacity` | `3372` | Recorre todas las llaves empacando hasta el cupo |
| `LBS_WalmartAnchorBefore` | `3169` | Elige la tarima ancla de un sándwich |
| `LBS_WalmartCannibalizeToCap` | `3511` | Quita tarimas de camiones sobre cupo |
| `LBS_WalmartSplitBareFullKeepProgramado` | `3442` | Parte un Full pelado conservando lo programado |
| `LBS_WalmartMinTarimas` | `4404` | Calcula el piso en tarimas (cupo por `LBS_WALMART_MIN_FILL`) |
| `LBS_EnforceWalmartMinFill` | `4441` | Aplica el piso: descarta bajo el piso, marca en la franja intermedia |
| `LBS_EnforceWalmartFolioCapAfterLayering` | `16370` | Revalida el cupo después de acomodar charolas |

### Unidades físicas (columna `AI`)

| Procedimiento | Línea | Qué hace |
|---|---|---|
| `LBS_WalmartAssignTarimaUnitAI` | `7186` | Asigna el identificador de unidad a cada tarima |
| `LBS_WalmartNextUnitAI` | `7131` | Genera el siguiente identificador libre del folio |
| `LBS_WalmartDedupDuplicateUnitAI` | `7266` | Quita identificadores repetidos |
| `LBS_WalmartSplitNonContiguousUnitAI` | `7342` | Separa una unidad cuyas filas no son contiguas |
| `LBS_WalmartSplitSharedUnitAI` | `7442` | Separa unidades compartidas indebidamente |
| `LBS_WalmartSplitBareFullsSharingAI` | `7478` | Separa Fulls pelados que comparten unidad |
| `LBS_WalmartSplitMixedDestUnitAI` | `7528` | Separa una unidad con destinos distintos |
| `LBS_WalmartSeparateConfPerUnit` | `7752` | Garantiza una confirmación por unidad |
| `LBS_WalmartResetMaxAICache` / `LBS_WalmartEnsureMaxAICache` | `217` / `224` | Caché O(1) del máximo identificador por folio |

### Catálogo TI HI y altura

| Procedimiento | Línea | Qué hace |
|---|---|---|
| `LBS_WalmartEnsureTihiDict` | `8337` | Carga el catálogo en memoria desde la hoja `TI HI` |
| `LBS_WalmartTihiLookupRec` | `8369` | Busca `Array(TI, HI, AltoCm, AltArmCm, PesoCase)` por cadena y material |
| `LBS_WalmartTihiFormulaHeightM` | `8403` | Altura por fórmula: capas por alto de caja |
| `LBS_WalmartTihiSotHeightM` | `8800` | Altura usando `Altura Armado` del catálogo |
| `LBS_WalmartMixedLayerCap` | `8437` | Cajas por cama en un clúster mezclado (`TI-1`) |
| `LBS_WalmartRowsSharedHeightM` | `8469` | Altura de filas que comparten cama |
| `LBS_WalmartComputeTihiAO` / `LBS_WalmartTryApplyTihiAO` | `8840` / `8813` | Calcula y escribe `AO` |
| `LBS_WalmartTryApplyTihiAT` / `LBS_WalmartRefreshAllTihiAT` | `8873` / `8930` | Escribe y refresca el peso `AT` |
| `LBS_WalmartRebuildAPTouched` | `8942` | Recalcula `AP` (suma de `AO` de la unidad) |
| `LBS_RecalcWalmartHeightsAOAP` | `9117` | La pasada completa de alturas |
| `LBS_WalmartRowFallbackAO` | `8449` | Altura de respaldo cuando el catálogo no tiene el SKU |
| `LBS_WalmartFillMissingSandwichAO` | `9461` | Rellena `AO` faltante desde la hoja de PCA |

### Caché de `Pallet Container Association`

| Procedimiento | Línea | Qué hace |
|---|---|---|
| `LBS_WalmartEnsurePcaHeightCache` | `8999` | Construye la caché plana de alturas por bloques |
| `LBS_WalmartPcaLastRow` | `8986` | Último renglón de la hoja de PCA |
| `LBS_WalmartPcaResetCache` | `8978` | Limpia la caché |
| `LBS_WalmartPCAHeightByMaterial` | `9422` | Altura por material y cantidad |
| `LBS_WalmartAnchorPCAHeight` | `11426` | Altura de la tarima ancla |

### Sándwich: acomodo y rescate

| Procedimiento | Línea | Qué hace |
|---|---|---|
| `LBS_WalmartRepackRestosSandwiches` | `9624` | Reacomoda restos en sándwiches |
| `LBS_WalmartTryStackSandwichUnitOntoUnderHeight` | `9890` | Intenta apilar sobre una unidad que aún tiene altura libre |
| `LBS_WalmartFindRemountAI` | `10149` | Busca una unidad donde remontar una tarima |
| `LBS_WalmartMergeThinProgramadoSandwiches` | `17898` | Une sándwiches delgados |
| `LBS_WalmartPrepSandwichFill` / `LBS_WalmartFillSandwichesFromDiscards` | `18730` / `18756` | Rellena sándwiches con producto descartado |
| `LBS_WalmartRebuildPalletsFromBajoLlenadoDiscards` | `19574` | Reconstruye tarimas desde descartes por bajo llenado |
| `LBS_WalmartMountRebuiltUnitsOntoAD` | `20278` | Monta las unidades reconstruidas en un folio |
| `LBS_WalmartSandwichAnchorTAdjust` | `7043` | Ajusta el conteo `T` de la tarima ancla |
| `LBS_WalmartCraftPalletWAdjust` | `7064` | Ajuste de `W` para tarimas craft |

### Descartes

| Procedimiento | Línea | Qué hace |
|---|---|---|
| `LBS_WalmartDiscardOverHeightSandwich` | `10467` | Descarta lo que excede 1.60 m |
| `LBS_WalmartDiscardOverWeightTruck` | `11068` | Descarta lo que excede 29 t |
| `LBS_WalmartRefreshAUFromAT` | `11373` | Recalcula el peso del camión desde el peso por SKU |
| `LBS_ClearWalmartDismountedInfo` | `20316` | Limpia la información de lo desmontado |

### Apertura de camiones

| Procedimiento | Línea | Qué hace |
|---|---|---|
| `LBS_IsWalmartOpenTruckGapAG` | `20418` | Decide si un motivo de `AG` es remontable |
| `LBS_WalmartOpenTrucksForNoReportadoGaps` | `20449` | Abre folios nuevos para el cartonaje no reportado |
| `LBS_WalmartTemplateADForCkey` | `20357` | Elige el folio plantilla para el nuevo camión |

## 5. Cómo validarlo

Fixtures en `tms_fg14/walmart/`:

| Archivo | Contenido |
|---|---|
| `sample.tsv` | Muestra de `Pedidos Surtidos` |
| `summaryok.tsv` | Salida esperada de `SummaryOK` |
| `plan.tsv`, `ogplan.tsv`, `testplan.tsv` | Planes de prueba |
| `shipmets.tsv` | Muestra de `Shipments` |
| `1302.tsv` | Caso puntual del folio 1302 |
| `cartonaje diff.tsv` | Diferencias de cartonaje contra el Plan |
| `consolidadiag.tsv` | Salida de `LBS_DiagnosticarConsolidacion` |
| `excepciones_sku.tsv` | Los SKU de la excepción EXC28 |
| `itemconections.tsv` | Conexiones de item del export |
| `condicionescadenas.tsv` | Condiciones de cadena |
| `small pallets.tsv`, `small_pallets_analysis.txt` | Análisis de tarimas chicas |

Scripts:

| Script | Qué comprueba |
|---|---|
| `scripts/validate_walmart_restos.py` | El acomodo de restos y sándwiches |
| `scripts/validate_walmart_heights.py` | El cálculo de `AO` y `AP` contra el tope de 1.60 m |
| `scripts/validate_walmart_conf_unit.py` | Una confirmación por unidad `AI` |
| `scripts/validate_walmart_exc_groupid.py` | La excepción EXC28 y el `Group Id` |
| `scripts/validate_plan_walmart_metro_volumes.py` | Volúmenes por metro en el Plan |
| `scripts/validate_tihi_ao_all_chains.py` | `AO` desde TI HI, todas las cadenas |

## 6. Problemas conocidos y síntomas

| Síntoma | Causa probable | Dónde mirar |
|---|---|---|
| `REVISION MANUAL: altura >1.6` | El TI HI del catálogo no corresponde al armado real, o se apilaron SKU que no debían compartir tarima | Hoja `TI HI`, columnas `J`, `K`, `O`, `Q` para ese material |
| `REVISION MANUAL: conf 33/20 en misma tarima AI (cap 28)` | Dos confirmaciones sobre la misma unidad que el motor no pudo separar | `LBS_WalmartSeparateConfPerUnit`; revisar el sándwich del folio |
| `AO` en cero con altura esperada | El SKU no está en la hoja `TI HI`, o está con `TI` o `Alto (CM)` en cero | Correr `LBS_SyncTihiSheet` y revisar el conteo de claves del mensaje |
| Camiones de 26 tarimas en lugar de 28 | El destinatario no está en la hoja `Consolida` **y** la fila no se reconoció como Walmart | Revisar el literal exacto de `Pedidos Surtidos!M`: tiene que ser `Walmart BA` o `Walmart SC` |
| `Descartado: bajo llenado` en camiones que parecen razonables | Quedaron bajo 12 tarimas | `LBS_EnforceWalmartMinFill`; revisar si la reconstrucción desde descartes alcanzó |
| Grupo completo descartado | Ningún folio pasó el gate y el rescate no llegó al 50 % del cupo | `LBS_WalmartGroupEfficiencyGate`; revisar el umbral en `EFICIENCIA POR CADENA` |
| Folios que deberían ir juntos, separados | Grupos distintos en la hoja `Consolida`, o vigencias distintas | `LBS_DiagnosticarConsolidacion`, columna `A` de `ConsolidaDiag` |
| Excel se cierra sin aviso durante `SummaryOK` | Hoja de PCA muy grande | `LBS_WALMART_PCA_MAX_ROWS` (250 000) en [09-parametros-y-catalogos.md](../09-parametros-y-catalogos.md) |

Una nota sobre depuración: `LBS_WalmartSalvageTrace` (`modulo2.vba:17853`) imprime las
decisiones de rescate en la ventana Inmediato, pero solo si `LBS_WALMART_SALVAGE_TRACE`
(`modulo2.vba:76`) está en `True`. El comentario advierte que activarla **puede congelar
Excel**, así que se usa únicamente con pocos datos y para un caso concreto.
