[Volver al índice de cadenas](README.md)

# OXXO y Neto

OXXO es la cadena con los cupos más documentados de toda la macro, porque los suyos **no
vienen del catálogo Mode Mix sino del maestro de equipos**, con una fecha de acuerdo con el
cliente escrita en el código. Neto comparte con OXXO la división por orígenes mezclados y la
llave de totales, pero sus cupos salen del catálogo y, como mayorista de la lista blanca,
también le aplica el cupo de 35 o 36. Aquí se documentan juntas porque el código las trata en
pareja en varios puntos.

## 1. Identificación

| Función | Línea | Literal |
|---|---|---|
| `LBS_IsOxxoChain` | `tms_fg14/modulo2.vba:16230` | `OXXO` |
| `LBS_IsNetoChain` | `tms_fg14/modulo2.vba:16234` | `NETO` |
| `PF_IsOxxoChain` | `tms_fg14/modulo5.vba:379` | `OXXO` (copia para `PartirTarimasFULL`) |

Ninguna tiene familia en `LBS_ChainFamily`, así que quedan fuera de las listas por familia
(peso excedente, apertura de camiones).

`Neto` aparece además en la lista blanca de mayoristas (`LBS_IsCap35AllowChain`,
`tms_fg14/modulo2.vba:5276`), y eso le agrega el tratamiento de
[mayoristas-cap35.md](mayoristas-cap35.md).

Un destinatario tiene nombre propio en el código: `400101621` es Monterrey
(`LBS_OXXO_MTY_DEST`, `tms_fg14/modulo2.vba:38`).

## 2. Parámetros

### OXXO

| Parámetro | Valor | Origen |
|---|---|---|
| Cupo sencillo general | 24 tarimas | `LBS_OXXO_SENCILLO_CAP` (`modulo2.vba:35`) |
| Cupo sencillo `Z4290_OXX` | 22 tarimas | `LBS_OXXO_Z4290_SENCILLO_CAP` (`modulo2.vba:36`) |
| Cupo sencillo Monterrey | 28 tarimas | `LBS_OXXO_MTY_SENCILLO_CAP` (`modulo2.vba:37`) |
| Cupo Full (tope) | 36 tarimas | `LBS_OXXO_SHIPMENT_CAP` (`modulo2.vba:32`) |
| Cupo caja `a`/`b` | 18 tarimas | `LBS_OXXO_BOX_CAP` (`modulo2.vba:31`) |
| Piso de llenado | 90 % del cupo = 33 de 36 | `LBS_OXXO_MIN_FILL` (`modulo2.vba:69`) |
| Altura máxima | Sin tope | `LBS_ChainEnforcesUnitHeight` no la incluye |
| Peso máximo | 29 t | `SK_MAX_PESO_KG` (`modulo2.vba:10`) |

El comentario que fecha el acuerdo (`tms_fg14/modulo2.vba:33-34`):

```
' OXXO Sencillo from equipment Max Pallet Count (client 29/07/2026):
' Z5290_OXX=24 (not catalog S/26 / generic Z5290), Z4290_OXX=22, Z5290_OXX_MTY=28.
```

Y el del tope de Full (`tms_fg14/modulo2.vba:5700-5702`):

```
' Catalogo Mode Mix Full Pallets Max; fallback OXXO=36 else 40.
' OXXO ceiling = LBS_OXXO_SHIPMENT_CAP (Z3500_OXX/Z2550_OXX Max=36) even when catalog
' has Full desenganche 40 (e.g. Pto Vallarta) — keep *_OXX until client changes equipment.
```

Estos dos comentarios son la razón por la que **editar el catálogo Mode Mix no cambia los
cupos de OXXO**. El código topa el cupo a 36 aunque el catálogo diga 40, y usa los cupos de
equipo para sencillo aunque el catálogo diga 26. El comentario dice hasta cuándo: `keep *_OXX
until client changes equipment`. Si el cliente cambia de equipo, hay que editar las
constantes.

El piso del 90 % es el más alto de la macro junto con Comextra:

```
' LBS - OXXO Full: piso de llenado post-consolidacion (90% del shipCap 36 -> piso 33).
Private Const LBS_OXXO_MIN_FILL As Double = 0.9
```

### Neto

| Parámetro | Valor | Origen |
|---|---|---|
| Cupo sencillo | Catálogo Mode Mix, respaldo 26 | `LBS_SencilloCapForRow` (`modulo2.vba:5841`) |
| Cupo Full | 35 si el SKU está en `Cadenas 35 Tarimas`, si no 36 | `LBS_CatalogFullShipmentCap` (`modulo2.vba:5754-5761`) |
| Cupo caja `a`/`b` | Cupo de embarque entre 2, salvo camión único de mayorista | `LBS_FullBoxCapForRow` (`modulo2.vba:5513`) |
| Piso de llenado | 90 % (comparte el mecanismo de OXXO) | `LBS_IsTruckMinFillChain` no lo incluye; su piso viene del empaque de mayoristas |
| Peso máximo | 29 t | `SK_MAX_PESO_KG` (`modulo2.vba:10`) |

## 3. Reglas de negocio

**Cómo se resuelve el cupo de sencillo de OXXO.** El comentario de
`LBS_ApplyOxxoSencilloEquipCap` (`tms_fg14/modulo2.vba:5767-5769`) describe el mapeo:

```
' Map OXXO Sencillo catalog/lane pallets to equipment Max:
' MTY (400101621) -> Z5290_OXX_MTY 28; catalog <=22 (often 20) -> Z4290_OXX 22;
' else Z5290_OXX 24 (never catalog S/26).
```

El orden es: si el destinatario es Monterrey, 28. Si el catálogo dice 22 o menos (casi
siempre 20), 22. En cualquier otro caso, 24. Nunca el 26 genérico del catálogo.

Lo que hace el catálogo aquí es solo **elegir entre dos equipos**: un carril con cupo bajo
significa que va en `Z4290_OXX` y por eso lleva 22. No es que el cupo del catálogo se use
directamente.

**Nunca se parte un sencillo.** Un sencillo de OXXO no se divide en cajas `a` y `b`. Los
recortes tienen procedimiento propio: `LBS_TrimOxxoSencilloOverCap`
(`tms_fg14/modulo2.vba:14199`), separado del recorte de Full
(`LBS_TrimOxxoFullOverCap`, `tms_fg14/modulo2.vba:14869`).

**Balanceo de cajas `a` y `b`.** Un Full de 36 se parte en 18 más 18. El comentario de
`LBS_RebalanceOxxoFullAb` (`tms_fg14/modulo2.vba:15965-15967`) trae el caso que lo motivó:

```
' After Full trim: a gets boxCap first, b gets the rest (max boxCap). Full 36 -> 18+18.
' Orphan a (or b) over boxCap: mint the missing twin and move excess before rebalance
' (P-1025a TW=36 with no b was box-trimmed to 18 and discarded 18 cupo).
```

El defecto era grave: el folio `P-1025a` tenía 36 tarimas y ninguna caja `b`. El recorte por
caja lo bajaba a 18 y **descartaba las otras 18**. La corrección crea la caja gemela que
falta y le mueve el excedente, antes de balancear.

**Llave de totales por embarque base.** OXXO y Neto agrupan el total de la columna `Z` por
embarque, no por folio. El comentario (`tms_fg14/modulo2.vba:16237`):

```
' Clave para total col Z: OXXO/NETO Full y splits PartirFulles a/b agrupan por shipment base.
```

Sin esto, un Full partido en `P-1025a` y `P-1025b` reportaría dos totales de 18 en lugar de
uno de 36, y el cliente vería dos embarques donde hay uno.

**Un origen por folio.** OXXO y Neto están en `LBS_IsMixedOriginSplitChain`
(`tms_fg14/modulo2.vba:13730`) junto con Chedraui y Walmart. Un folio con dos plantas se
divide en folios `P-####` nuevos.

**División de restos al partir.** `PartirTarimasFULL` tiene un caso especial para OXXO:
`PF_SplitOxxoRestosRow` (`tms_fg14/modulo5.vba:1325`) divide una fila de restos entre las dos
cajas, insertando renglones. Detalle en
[06-partir-tarimas-full.md](../06-partir-tarimas-full.md).

**Sin tope de altura.** OXXO y Neto no están en `LBS_ChainEnforcesUnitHeight`
(`tms_fg14/modulo2.vba:11895`) ni en `LBS_ChainAllowsSharedCamas`
(`tms_fg14/modulo2.vba:6342`). No se les calcula altura de unidad ni se comparten camas.

**Sin sándwich por unidad.** Tampoco están en `LBS_IsSandwichWChain`
(`tms_fg14/modulo2.vba:6353`).

**Bandera de presencia.** `LBS_HasOxxo` (`tms_fg14/modulo2.vba:6275`) permite al
consolidador de fallos saltarse las fases de OXXO cuando no hay ninguna fila de la cadena.

## 4. Procedimientos

| Procedimiento | Línea | Qué hace |
|---|---|---|
| `LBS_IsOxxoChain` | `16230` | Reconoce `OXXO` |
| `LBS_IsNetoChain` | `16234` | Reconoce `NETO` |
| `LBS_HasOxxo` | `6275` | Bandera de presencia en la hoja |
| `LBS_ApplyOxxoSencilloEquipCap` | `5770` | Traduce el cupo de catálogo al cupo de equipo (24 / 22 / 28) |
| `LBS_CatalogSencilloShipmentCap` | `5784` | Cupo de sencillo desde el catálogo, antes del ajuste por equipo |
| `LBS_CatalogFullShipmentCap` | `5703` | Cupo de Full, con el tope de 36 para OXXO |
| `LBS_TrimOxxoOverCap` | `14179` | Recorte general sobre cupo |
| `LBS_TrimOxxoSencilloOverCap` | `14199` | Recorte de sencillos |
| `LBS_TrimOxxoFullOverCap` | `14869` | Recorte de Fulls por caja |
| `LBS_RebalanceOxxoFullAb` | `15968` | Balancea las cajas `a` y `b`, creando la gemela faltante |
| `LBS_MoveOxxoFullTarimasToAd` | `16094` | Mueve tarimas entre folios de Full |
| `LBS_ApplyOxxoTarimaCapFlags` | `16256` | Escribe las marcas de `REVISION MANUAL` por cupo |
| `LBS_TarimaTotalGroupKey` | `16238` | Llave del total `Z` por embarque base |
| `LBS_IsMixedOriginSplitChain` | `13730` | Lista de cadenas con división por origen |
| `PF_IsOxxoChain` | `modulo5.vba:379` | Copia del clasificador en `PartirTarimasFULL` |
| `PF_SplitOxxoRestosRow` | `modulo5.vba:1325` | Divide una fila de restos entre las dos cajas |

La fase de `SummaryOK` correspondiente es `totales:trim_oxxo_cap`
(`tms_fg14/modulo2.vba:1575`).

## 5. Cómo validarlo

Fixtures en `tms_fg14/oxxo/`:

| Archivo | Contenido |
|---|---|
| `sample.tsv` | Muestra de `Pedidos Surtidos` |
| `summaryok.tsv` | Salida esperada de `SummaryOK` |
| `plan.tsv` | Plan de prueba |
| `pca.tsv` | Muestra de `Pallet Container Association` |
| `equipment.tsv` | Los equipos `*_OXX` con su `Max Pallet Count` |

El archivo `equipment.tsv` es el que respalda los cupos de 24, 22 y 28: ahí están los
`Z5290_OXX`, `Z4290_OXX` y `Z5290_OXX_MTY` con sus máximos.

Para Neto, `tms_fg14/neto/sample.tsv`.

Scripts:

| Script | Qué comprueba |
|---|---|
| `scripts/validate_tms_oxxo.py` | La salida a TMS de OXXO |
| `scripts/validate_oxxo_sample_z.py` | El total de la columna `Z` por embarque base |
| `scripts/validate_oxxo_body_type.py` | El tipo de caja asignado |
| `scripts/validate_partir_fulles_blocks.py` | La división en cajas `a` y `b` |

## 6. Problemas conocidos y síntomas

| Síntoma | Causa probable | Dónde mirar |
|---|---|---|
| Sencillos de 26 tarimas en OXXO | La fila no se reconoció como OXXO | Revisar el literal exacto de `Pedidos Surtidos!M`; el comentario dice `never catalog S/26` |
| Cambiar el catálogo Mode Mix no mueve el cupo | Es deliberado: OXXO usa cupos de equipo, topados a 36 en Full | Constantes `LBS_OXXO_*` en `modulo2.vba:31-38` |
| Full de 40 en OXXO | No debería pasar: el tope es 36 | `LBS_CatalogFullShipmentCap`, líneas `5744-5746` |
| Caja `a` con 36 tarimas y sin caja `b` | El caso `P-1025a` del comentario | `LBS_RebalanceOxxoFullAb` debería crear la gemela. Si no, es una regresión |
| Dos embarques donde debería haber uno | El total `Z` se agrupó por folio en lugar de por embarque base | `LBS_TarimaTotalGroupKey` |
| `Descartado por baja eficiencia` con 30 o 32 tarimas | Bajo el piso de 33, que es el 90 % de 36 | `LBS_OXXO_MIN_FILL`. Es el piso más exigente de la macro |
| Monterrey con cupo 24 en lugar de 28 | El destinatario de la columna `O` no es exactamente `400101621` | `LBS_OXXO_MTY_DEST` |
| `AO` vacío en OXXO | Es lo esperado: OXXO no calcula altura de unidad | `LBS_ChainEnforcesUnitHeight` no lo incluye |
| SKU de Neto que no llega a 35 | No está marcado `YES` en `Cadenas 35 Tarimas`, o `Neto` no está en la lista blanca | Ver [mayoristas-cap35.md](mayoristas-cap35.md) |
| Folios `P-####` nuevos | División por orígenes mezclados | `LBS_SplitChedrauiMixedOriginFolios`, que también sirve a OXXO y Neto |
