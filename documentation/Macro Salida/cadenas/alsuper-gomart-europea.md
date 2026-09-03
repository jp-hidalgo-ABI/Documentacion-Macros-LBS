[Volver al índice de cadenas](README.md)

# Alsuper, Go Mart y Europea

Las tres se agrupan bajo el nombre "Mode Mix autoservicio" porque comparten una
característica que ninguna otra cadena tiene: sus cupos **y** su techo de peso salen del
catálogo `Catalogo Mode Mix`, incluyendo el techo de 52.5 toneladas que solo aplica a sus
carriles Full. Alsuper además reutiliza el gate de eficiencia de Walmart.

## 1. Identificación

| Función | Línea | Literales |
|---|---|---|
| `LBS_IsAlsuperChain` | `tms_fg14/modulo2.vba:6183` | `ALSUPER` |
| `LBS_IsGoMartChain` | `tms_fg14/modulo2.vba:6187` | `GOMART` |
| `LBS_IsEuropeaChain` | `tms_fg14/modulo2.vba:6191` | `EUROPEA`, `LAEUROPEA` |
| `LBS_IsModeMixAutoservicioChain` | `tms_fg14/modulo2.vba:6198` | Las tres juntas |

Europea es la única de las tres que acepta dos literales, con y sin artículo:

```
u = Replace(UCase$(Trim$(CStr(m))), " ", "")
LBS_IsEuropeaChain = (u = "EUROPEA" Or u = "LAEUROPEA")
```

El comentario del agrupador (`tms_fg14/modulo2.vba:6197`) resume el criterio:

```
' Alsuper / Go Mart / Europea: Mode Mix caps (26 S / 40 F) and catalog weight ceilings.
```

Las tres tienen familia propia en `LBS_ChainFamily` (`tms_fg14/modulo2.vba:5207-5212`):
`ALSUPER`, `GOMART` y `EUROPEA`. Eso las mete en las listas por familia de peso excedente
(`LBS_IsPesoSalvageChain`, `tms_fg14/modulo2.vba:6313`) y apertura de camiones
(`LBS_IsOpenTruckChain`, `tms_fg14/modulo2.vba:6321`).

## 2. Parámetros

| Parámetro | Valor | Origen |
|---|---|---|
| Cupo sencillo | Catálogo Mode Mix, respaldo 26 | `LBS_SencilloCapForRow` (`modulo2.vba:5841`) |
| Cupo Full | Catálogo Mode Mix, respaldo 40 | `LBS_CatalogFullShipmentCap` (`modulo2.vba:5703`) |
| Cupo caja `a`/`b` | Cupo de embarque entre 2, normalmente 20 | `LBS_FullBoxCapForRow` (`modulo2.vba:5513`) |
| Piso de llenado | 40 % del cupo | `LBS_WALMART_MIN_FILL` (`modulo2.vba:65`) |
| Rescate de grupo | 50 % del cupo | `LBS_WALMART_RESCUE_MIN_FILL` (`modulo2.vba:60`) |
| Altura máxima de unidad | 1.60 m | `LBS_ChainMaxUnitHeightM` (`modulo2.vba:11908-11912`) |
| Peso máximo Full | **52.5 t** | `LBS_FULL_MAX_PESO_KG` (`modulo2.vba:12`) |
| Peso máximo sencillo | 29 t | `SK_MAX_PESO_KG` (`modulo2.vba:10`) |

El techo de peso es lo que las distingue:

```
' Catalogo Mode Mix Full weight ceiling (52.5 t) for Alsuper/Go Mart Full lanes.
Private Const LBS_FULL_MAX_PESO_KG As Double = 52500#
```

`LBS_MaxPesoKgForRow` (`tms_fg14/modulo2.vba:11506`) es quien decide por fila, y la condición
es doble (`tms_fg14/modulo2.vba:11515`): el equipo de la columna `Y` tiene que decir `Full`
**y** la cadena tiene que ser una de las tres. Un sencillo de Alsuper lleva 29 t como
cualquier otro.

El piso y el rescate son los de Walmart. El comentario de la constante
(`tms_fg14/modulo2.vba:61-62`) lo dice al pasar:

```
' LBS - WALMART: piso de llenado de camion (tarimas). El % de "EFICIENCIA POR CADENA" es
' fill rate de PEDIDO (col AR, gate de eficiencia) para Walmart/Alsuper — truck floor is
' this constant (cap 28 -> piso 12).
```

Con la salvedad de que en estas cadenas el cupo no es 28 sino el del catálogo, así que el
piso efectivo es el 40 % de ese número: 10 sobre un sencillo de 26, 16 sobre un Full de 40.

## 3. Reglas de negocio

**El cupo viene del catálogo, no de Walmart.** Es una distinción explícita del código
(`tms_fg14/modulo2.vba:5889`):

```
' Alsuper/GoMart/Europea use Mode Mix above (26 S / 40 F) — not Walmart 28.
```

Aunque comparten piso y gate con Walmart, el cupo es distinto. Un camión de Alsuper lleva 26,
no 28.

**Corrección de equipo en destinos solo-Full.** El comentario de `LBS_ForceModeMixFullY`
(`tms_fg14/modulo2.vba:12419-12420`) describe el problema y la solución:

```
' Alsuper/Go Mart: LBS sometimes emits Sencillo on a Full-only Mode Mix dest
' (Chihuahua 400002423). Rewrite Y so cupo/min-fill/weight use Full 40 / 52.5 t.
```

LBS a veces asigna un sencillo a un destino que en el catálogo solo tiene carriles Full. La
macro reescribe la columna `Y` a `Full caja seca` para que el cupo, el piso y el techo de
peso sean los correctos. El destino de Chihuahua (`400002423`) es el caso que lo detonó.

Detalles de la corrección: solo aplica a filas `Programado`, solo a Alsuper y Go Mart (**no**
a Europea), solo si el equipo dice `Sencillo`, y solo si el destino es solo-Full según
`LBS_CatalogDestIsFullOnly` (`tms_fg14/modulo2.vba:12399`). Corre tres veces: al final de
`SummaryOK` (`tms_fg14/modulo2.vba:1577`), al inicio del filtro de eficiencia
(`tms_fg14/modulo2.vba:4864`) y en el consolidador de restos
(`tms_fg14/modulo2.vba:21167`).

**Colapso de sufijos antes de consolidar.** Es la regla más delicada de Alsuper. El
comentario (`tms_fg14/modulo2.vba:4765-4769`):

```
' ALSUPER Full caja seca: LBS puede traer las cajas a/b ya partidas (P-1390a/P-1390b). La
' consolidacion (merge/gate/min-fill) trabaja por SHIPMENT (cap 40 = a+b), no por caja (20):
' tratar cada caja como camion de 40 duplica cupo y rompe el balance a/b (p.ej. a=31 sin b).
' Se colapsan los sufijos al folio base antes del merge; PartirFulles (paso 3) rehace el
' split a/b ~20/20 sobre el shipment ya consolidado.
```

En prosa: LBS puede entregar los Fulls ya divididos en `P-1390a` y `P-1390b`. Si la
consolidación tratara cada caja como un camión de cupo 40, duplicaría el cupo real y
produciría desbalances como una caja `a` con 31 tarimas y ninguna `b`. Por eso el motor pega
las cajas al folio base antes de consolidar, y deja que `PartirTarimasFULL` las vuelva a
dividir 20 y 20 sobre el embarque ya armado.

El orden de las tres fases importa: colapsar, consolidar, volver a partir. Alterarlo
reintroduce el desbalance.

**Gate compartido con Walmart.** `LBS_WalmartGroupEfficiencyGate`
(`tms_fg14/modulo2.vba:2903`) tiene un parámetro `chainKind` que por omisión vale `WALMART` y
para estas cadenas se pasa como Alsuper. El comentario
(`tms_fg14/modulo2.vba:2898`) lo declara:

```
' WALMART BA/SC (y ALSUPER via chainKind): eficiencia por grupo de consolidacion antes del
' descarte folio-a-folio.
```

Es el mismo algoritmo: promedio ponderado por tarimas, rescate multi-camión desde el 50 % del
cupo, y unión de folios fallidos sobre los que pasan.

**Peso desde el catálogo TI HI.** Las tres están en `LBS_ChainAppliesTihiAT`
(`tms_fg14/modulo2.vba:6336`): el peso `AT` se refresca desde `Peso Bruto de material`.

**Camas compartidas, altura y sándwich por unidad.** Las tres están en las tres listas:
`LBS_ChainAllowsSharedCamas` (`6342`), `LBS_ChainEnforcesUnitHeight` (`11895`) y
`LBS_IsSandwichWChain` (`6353`). Comparten con Walmart casi toda la maquinaria física.

**Corrección de comentarios de tarima cero.** `LBS_FixModeMixZeroTarimaComment`
(`tms_fg14/modulo2.vba:12442`) reescribe los comentarios `+0_Tarima ...+` de las columnas
`AE` a `AG` cuando las cajas son mayores que cero, para que las auditorías de material
excedente dejen de marcar tarimas fantasma. Aplica a las tres cadenas.

## 4. Procedimientos

| Procedimiento | Línea | Qué hace |
|---|---|---|
| `LBS_IsAlsuperChain` | `6183` | Reconoce `ALSUPER` |
| `LBS_IsGoMartChain` | `6187` | Reconoce `GOMART` |
| `LBS_IsEuropeaChain` | `6191` | Reconoce `EUROPEA` y `LAEUROPEA` |
| `LBS_IsModeMixAutoservicioChain` | `6198` | Las tres juntas |
| `LBS_HasWalmartOrAlsuper` | `6267` | Bandera de presencia de Walmart o Alsuper |
| `LBS_CollapseAlsuperFullSuffixFolios` | `4770` | Pega las cajas `a`/`b` al folio base antes de consolidar |
| `LBS_CatalogDestIsFullOnly` | `12399` | Detecta destinos que solo tienen carriles Full |
| `LBS_ForceModeMixFullY` | `12421` | Reescribe el equipo a `Full caja seca` en destinos solo-Full |
| `LBS_FixModeMixZeroTarimaComment` | `12442` | Corrige los comentarios `+0_Tarima` |
| `LBS_MaxPesoKgForRow` | `11506` | Devuelve 52.5 t en Full de estas cadenas, 29 t en el resto |
| `LBS_WalmartGroupEfficiencyGate` | `2903` | El gate compartido, vía `chainKind` |

La fase de `SummaryOptimizar` correspondiente es `filtrar:alsuper_gate`
(`tms_fg14/modulo2.vba:4967`), y el colapso de sufijos ocurre en `filtrar:merge_metro`
(`tms_fg14/modulo2.vba:4880-4883`).

## 5. Cómo validarlo

Fixtures:

| Archivo | Contenido |
|---|---|
| `tms_fg14/alsuper/sample.tsv` | Muestra de `Pedidos Surtidos` |
| `tms_fg14/alsuper/plan.tsv` | Plan de prueba |
| `tms_fg14/alsuper/RERUN_CHECKLIST.md` | Lista de verificación para volver a correr el caso |
| `tms_fg14/gomart/sample.tsv` | Muestra de Go Mart |
| `tms_fg14/gomart/itempackage.tsv` | Empaque por item |

`RERUN_CHECKLIST.md` es el único fixture con instrucciones escritas: conviene leerlo antes de
tocar la lógica de Alsuper.

Scripts:

| Script | Qué comprueba |
|---|---|
| `scripts/validate_plant_sencillo_chains.py` | Qué cadenas van en sencillo por planta |
| `scripts/validate_7e_itempackage_armado.py` | El armado contra el empaque por item |
| `scripts/validate_ae_pallet_mix.py` | La mezcla de tarimas |
| `scripts/validate_tihi_ao_all_chains.py` | `AO` desde TI HI |

## 6. Problemas conocidos y síntomas

| Síntoma | Causa probable | Dónde mirar |
|---|---|---|
| Caja `a` con 31 tarimas y sin caja `b` | Es exactamente el defecto que el colapso de sufijos previene | `LBS_CollapseAlsuperFullSuffixFolios`. Si aparece, verificar el orden colapsar / consolidar / partir |
| Camiones de 28 tarimas en Alsuper | Se aplicó el cupo de Walmart en lugar del catálogo | El comentario dice `not Walmart 28`. Es una regresión |
| Sencillo en un destino que solo tiene carriles Full | La corrección de equipo no aplicó | `LBS_ForceModeMixFullY`; recordar que **no** cubre a Europea |
| Marca de peso a 29 t en un Full de Alsuper | La columna `Y` no dice `Full` | `LBS_MaxPesoKgForRow` exige las dos condiciones |
| Marca de peso a 52.5 t en un sencillo | No debería: el techo alto solo aplica a Full | Igual |
| Europea con equipo sencillo en destino solo-Full | Es el comportamiento actual: la corrección excluye a Europea | Si el cliente lo pide, hay que agregarla en `modulo2.vba:12431` |
| Tarimas fantasma en auditorías de material | Comentarios `+0_Tarima` sin corregir | `LBS_FixModeMixZeroTarimaComment` |
| `Descartado: bajo llenado` con 9 tarimas en sencillo | Bajo el 40 % de 26, que son 10 | `LBS_WALMART_MIN_FILL` aplicado sobre el cupo del catálogo |
| `Europea` escrita como `La Europea` sin reconocerse | Debería funcionar: acepta los dos literales | Revisar espacios o acentos en `Pedidos Surtidos!M` |
