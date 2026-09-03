[Volver a cadenas](README.md) · [Macro de entrada](../README.md) · [Índice general](../../README.md)

# Otras cadenas — macro de entrada

Las cadenas que no tienen página propia. Casi todas se distinguen únicamente por su sufijo de
item; las que tienen lógica adicional se detallan abajo.

## Resolución del `Item Id`: el orden real

Antes de entrar en cada cadena conviene fijar el orden completo con que `dp` resuelve el
`Item Id` de una fila del STR, porque las excepciones de esta página se ubican en distintos
puntos de esa cascada.

`merged/modulo1.vba:391-427`

| Paso | Condición | Resultado |
|---|---|---|
| 1 | `Plan!G = "Mercado Libre"` | `<SKU>_MER_LIBRE` (valor fijo en el código) |
| 2 | Cualquier otra cadena | `<SKU>_<LBS_CadenaItemSuffix(cadena)>` |
| 3 | El tipo de tarima es plástica o madera | **`<SKU>` desnudo**, se descarta el sufijo |
| 4 | El tipo de tarima es CHEP y existe la fila en `ArmadoChep` | La columna `E` de `ArmadoChep`, tal cual |

Los pasos 3 y 4 son sobreescrituras: se aplican después de haber calculado el sufijo. Por eso
un item de tarima de madera nunca lleva sufijo, sin importar la cadena.

**Mercado Libre no está en la tabla `LBS_CadenaItemSuffix`.** Su sufijo `_MER_LIBRE` está
escrito directamente en el bucle de resolución (`merged/modulo1.vba:395-397`), y hay una
segunda copia de la misma condición en `LBS_ResolveItemIdFromPlanRow`
(`merged/modulo1.vba:1067`), que es la función que usa `ValidarExportMEXKA` para verificar.
Ambas tienen que coincidir; si alguien cambiara solo una, el validador reportaría un
desajuste falso.

Otro detalle del paso 1: la comparación de Mercado Libre es **sensible a mayúsculas**
(`cadenaLB = "Mercado Libre"`, sin `UCase$`), a diferencia de todas las demás. Un `Plan!G`
con `MERCADO LIBRE` en mayúsculas caería en el paso 2 y produciría el sufijo genérico
`LEFT` = `MER`.

## Rabbit

Rabbit es la cadena con más lógica de esta página. Sufijo: `RAB`
(`merged/modulo1.vba:981`). Destino FG14: `FG14_RABBIT` o `FG14_RABBIT_FULL`.

### Los SKU de latón

`LBS_IsLatonSkuBare` (`merged/modulo2.vba:210-220`)

```
3003697  3003911  3007678  3006134  3009209  3018106  3018557  3018636  3017708
```

El comentario los describe (`merged/modulo2.vba:210`):
*"Laton madera SKUs (Rabbit FG14): full lane when tarimas >= 35 (manual build 4200 @ 120)"*.

`LBS_IsLatonSkuPlanRow` (`merged/modulo2.vba:222-231`) es la versión que recibe el número de
fila del Plan y lee el SKU de `Plan!T`.

### El carril `RABBIT_FULL`

`LBS_Fg14RabbitLaneSlug` (`merged/modulo2.vba:251-272`)

Un renglón de Rabbit va al carril `RABBIT_FULL` en lugar de `RABBIT` cuando se cumplen las
dos condiciones:

1. El SKU está en la lista de latón.
2. **`Plan!N >= 35` tarimas, o bien** el cartonaje (el mayor entre `Plan!M` y `Plan!AA`)
   `>= 4200` cajas.

Las dos constantes son locales a la función (`merged/modulo2.vba:257-258`):

```vba
Const FULL_TARIMAS As Long = 35
Const FULL_CARTONS As Long = 4200
```

Los dos umbrales son la misma cantidad expresada de dos maneras: 35 tarimas de 120 cajas son
4 200 cajas. El comentario lo anota como *"manual build 4200 @ 120"*. Se evalúan con un `Or`
para tolerar que el armado registrado no sea exactamente 120.

La función lee las cantidades con `LBS_PlanLongValue` (`merged/modulo2.vba:234`), que maneja
valores como `" 4,200 "` con comas y espacios no separables. Es un detalle práctico: el Plan
suele traer las cantidades con formato de miles.

### Consecuencias del carril `RABBIT_FULL`

| Efecto | Cita |
|---|---|
| El STR **nunca se parte**. El comentario dice *"never split STR - stays on RABBIT_FULL lane"* | `merged/modulo1.vba:230-231` |
| `EqByLane` emite **solo `Z3500`**, sin fila de sencillo | `merged/modulo1.vba:1497-1498` |
| `itemConnections` marca el item como tarima completa sin mezcla | `merged/modulo1.vba:1828-1829` |

Para los renglones de Rabbit que **no** son latón o no alcanzan el umbral, el carril es
`FG14_RABBIT` y `EqByLane` emite `Z5290` con banderas de tarima completa
(`merged/modulo1.vba:1515-1516`). El comentario aclara que esas banderas no son las de OXXO
aunque se parezcan.

## Go Mart

Sufijo: `GOMC` (`merged/modulo1.vba:976`).

`LBS_IsGoMartCadena` (`merged/modulo2.vba:461-465`) reconoce exactamente `GO MART`, con
espacio.

Su particularidad está anotada en el comentario (`merged/modulo2.vba:460`):

```
' GO MART ships Tarima Plástica (client plan col "Armado de Tarimas" -> Plan!P).
```

Go Mart siempre embarca en tarima plástica, así que por el paso 3 de la cascada de arriba sus
items **salen con el SKU desnudo, sin el sufijo `GOMC`**. El sufijo está registrado por
completitud, pero en la práctica casi nunca se aplica.

La segunda parte del comentario indica que el armado viene de la columna del Plan del
cliente, no del catálogo CHEP. Es coherente: el catálogo TI HI está construido para armados
CHEP.

## Amazon, Mercado Libre y Rapiturbo

Los tres son destinos FG14 de comercio electrónico.

| Cadena | Sufijo | Slug FG14 | Cita del slug |
|---|---|---|---|
| `AMAZON` | `AMA` | `AMAZON` | `merged/modulo2.vba:198` |
| `MERCADO LIBRE` | `_MER_LIBRE` (fijo en el código) | `MERCADO_LIBRE` | `merged/modulo2.vba:199` |
| `RAPITURBO` / `RAPPITURBO` / `RAPPI TURBO` | `Rappi Turbo` | `RAPPITURBO` | `merged/modulo2.vba:197` |

Rapiturbo acepta **tres** grafías distintas en el slug pero solo **dos** en el sufijo
(`RAPITURBO` y `RAPPI TURBO`, `merged/modulo1.vba:979`). Un `Plan!G` con `RAPPITURBO` sin
espacio recibe el slug FG14 correcto pero el sufijo genérico `LEFT` = `Rap`.

El sufijo `Rappi Turbo` es el único de la tabla que conserva mayúsculas y minúsculas mixtas y
que contiene un espacio. No es un error: así lo espera el maestro de items.

Los tres se registran en la hoja `FG14 Destinos` y `LBS_AplicarFg14DestinosAMaestros`
(`merged/modulo4.vba:1085`) siembra sus filas en `Handling` y `TarimaPorDestino`.

## Super Ofertas, Super Bara y Super Kompras

| Cadena | Sufijo | Cita |
|---|---|---|
| `SUPER OFERTAS` | `SUP_OFE` | `merged/modulo1.vba:972` |
| `SUPER BARA` | `SUP` | `merged/modulo1.vba:971` |
| `SUPER KOMPRAS` | `SUP_KOM` | `merged/modulo1.vba:973` |

`SUP` a secas pertenece a Super Bara, no es un prefijo genérico. Vale la pena tenerlo
presente al agregar una cadena nueva que empiece con "Super".

### El partido de Super Ofertas en bloques de 35

Super Ofertas es la única cadena cuyos renglones se parten en varias filas de STR. Ver el
detalle en
[../05-generacion-mexka.md](../05-generacion-mexka.md#el-partido-de-super-ofertas).

Resumen: cuando el material está listado en `Condiciones Cadenas` y `Plan!N > 35`, el renglón
se convierte en varias filas de hasta 35 tarimas cada una
(`merged/modulo1.vba:265-286`). La cantidad de cada parte se reparte proporcionalmente
(`merged/modulo1.vba:271`).

Las partes que salen de exactamente 35 tarimas reciben `Consolidation Class` igual al ID del
STR (`merged/modulo1.vba:546-548`), la misma protección que los camiones EXC28 de Walmart:
un camión completo no se mezcla ni se parte. Las partes con resto caen en la lógica normal
(`merged/modulo1.vba:551-556`).

El ejemplo de la hoja de muestra [merged/plan.tsv](../../../merged/plan.tsv) es precisamente
un renglón de Super Ofertas: 8 400 cajas con armado de 120 dan 70 tarimas, que se parten en
dos bloques de 35.

## Cadenas con sufijo dedicado y sin otra lógica

| Cadena | Sufijo | Cita |
|---|---|---|
| `SMART AND FINAL` | `SMA_FIN` | `merged/modulo1.vba:970` |
| `CABRITO ABARROTERO` | `CAB_ABA` | `merged/modulo1.vba:974` |
| `TOBE DISTRIBUTIONS` | `TOB_DIS` | `merged/modulo1.vba:975` |
| `ASTURIANO` | `ASTURIAN` | `merged/modulo1.vba:977` |
| `CONASUPER` | `CONASUPE` | `merged/modulo1.vba:978` |
| `EUROPEA` | `EUROPEA` | `merged/modulo1.vba:982` |

`ASTURIAN` y `CONASUPE` son truncamientos a ocho caracteres de `ASTURIANO` y `CONASUPER`, no
abreviaturas arbitrarias. `EUROPEA` es el nombre completo.

## Cadenas sin sufijo dedicado

Todas las demás usan `LEFT(cadena, 3)` (`merged/modulo1.vba:983-984`). Entre ellas:

| Cadena | Sufijo resultante |
|---|---|
| `CHEDRAUI` | `CHE` |
| `OXXO` | `OXX` |
| `SORIANA` | `Sor` |
| `CITY CLUB` | `Cit` |
| `CITY FRESKO` | `Cit` |
| `HEB` | `HEB` |
| `NETO` | `NET` |
| `MERCO` | `Mer` |
| `MERZA` | `Mer` |
| `CASA LEY` | `Cas` |
| `SAMS` | `Sam` |

Nótese que el atajo **conserva las mayúsculas y minúsculas originales** de `Plan!G`, porque
opera sobre el texto sin normalizar (`merged/modulo1.vba:984`). Por eso la tabla muestra
`Sor` y no `SOR`: depende de cómo esté capturado el nombre en el Plan. Si un día alguien
captura `SORIANA` en mayúsculas y otro día `Soriana`, se generan dos items distintos para el
mismo producto.

Ese es un argumento adicional para mantener la columna `E` de `ArmadoChep` completa: al venir
de una tabla, el `Item Id` es estable sin importar cómo se capturó la cadena.

### Colisiones conocidas del atajo

| Sufijo | Cadenas que lo comparten |
|---|---|
| `Cit` | City Club y City Fresko |
| `Mer` | Merco y Merza |

En ambos casos, si falta la fila de `ArmadoChep`, dos cadenas distintas producen el mismo
`Item Id` y LBS las consolida como si fueran la misma.

## Destino HEB `400170996`

`ValidarExportMEXKA` tiene tres reglas específicas para este destino
(`merged/modulo6.vba:2266-2270`):

| Mensaje | Significado |
|---|---|
| `HEB 400170996: itemConnections OK pero connections vacio` | Falta el carril en `connections` |
| `HEB 400170996: en STR pero sin itemConnections` | Falta la fila de item |
| `HEB 400170996: STR + itemConnections + connections alineados` | Mensaje OK |

Este destino no tiene lógica especial en la generación; las validaciones existen porque fue
el caso donde se detectó el problema de `connections` vacío que produce
`Invalid Connection Id`. Se conservan como canario.

## Los 15 items de NETO

`merged/modulo6.vba:2316-2328`. `ValidarExportMEXKA` verifica que los 15 `Item Id` de NETO
estén en la hoja `items` del export:

| Mensaje | Significado |
|---|---|
| `items NETO: no hay pedidos NETO en plan (N/A)` | No aplica en esta corrida |
| `items: 15 Item Id NETO en catalogo items` | Mensaje OK |
| `items: faltan N NETO en hoja items` | Correr `Sync_Step4_ExportItems` |

Hay una validación relacionada sobre el sufijo (`merged/modulo6.vba:3717`):
`STR: N Item Id *_NET con SKU base en catalogo items`, que detecta items con sufijo `_NET`
cuando el maestro tiene el SKU base.

## En la macro de salida

Varias de estas cadenas tienen reglas de armado propias del lado de salida:

| Cadena | Página |
|---|---|
| Go Mart, Alsuper, Europea | [../../tms_fg14/cadenas/alsuper-gomart-europea.md](../../tms_fg14/cadenas/alsuper-gomart-europea.md) |
| ToBe, Conasuper, Neto, Cabrito Abarrotero, Super Ofertas | [../../tms_fg14/cadenas/mayoristas-cap35.md](../../tms_fg14/cadenas/mayoristas-cap35.md) |
| Sams | [../../tms_fg14/cadenas/sams.md](../../tms_fg14/cadenas/sams.md) |
| HEB, Smart, Merco, Merza, Casa Ley | [../../tms_fg14/cadenas/otras-cadenas.md](../../tms_fg14/cadenas/otras-cadenas.md) |
