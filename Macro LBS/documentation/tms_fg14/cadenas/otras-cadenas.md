[Volver al índice de cadenas](README.md)

# HEB, Smart, Merco, Merza, Casa Ley y el comportamiento por omisión

Cinco cadenas tienen clasificador propio pero **ninguna lógica de armado propia**: aparecen
solo en las listas transversales de comportamiento. Y todo lo que no cae en ningún
clasificador usa el comportamiento por omisión, que también se documenta aquí porque es lo
que reciben la mayoría de las cadenas del Plan.

## 1. Identificación

| Función | Línea | Literal aceptado |
|---|---|---|
| `LBS_IsHebChain` | `tms_fg14/modulo2.vba:6287` | `HEB` |
| `LBS_IsSmartChain` | `tms_fg14/modulo2.vba:6291` | `SMART` |
| `LBS_IsMercoChain` | `tms_fg14/modulo2.vba:6295` | `MERCO` |
| `LBS_IsMerzaChain` | `tms_fg14/modulo2.vba:6299` | `MERZA` |
| `LBS_IsCasaLeyChain` | `tms_fg14/modulo2.vba:6303` | `CASALEY` |

Las cuatro primeras normalizan igual que el resto: mayúsculas, sin espacios, comparación
exacta. `Smart and Final` **no** se reconoce como `Smart`: la normalización no recorta
palabras.

Casa Ley es la única de toda la macro que además quita guiones:

```
LBS_IsCasaLeyChain = (Replace(Replace(UCase$(Trim$(CStr(m))), " ", ""), "-", "") = "CASALEY")
```

Así que acepta `Casa Ley`, `CASA-LEY`, `Casa-Ley` y `CASALEY`.

Ninguna de las cinco tiene familia en `LBS_ChainFamily` (`tms_fg14/modulo2.vba:5198`).

## 2. Parámetros

Los mismos para las cinco cadenas y para el comportamiento por omisión, salvo la altura.

| Parámetro | Valor | Origen |
|---|---|---|
| Cupo sencillo | 26 tarimas | `LBS_TruckCapForGroup` (`modulo2.vba:5497`) |
| Cupo Full | 40 tarimas si el equipo dice `Full` | `LBS_TruckCapForGroup` (`modulo2.vba:5499-5502`) |
| Cupo caja `a`/`b` | 20 tarimas | `LBS_FULL_BOX_CAP` (`modulo2.vba:29`) |
| Piso de llenado | Ninguno propio | No están en `LBS_IsTruckMinFillChain` (`modulo2.vba:4434`) |
| Altura máxima | 1.60 m **solo HEB** | `LBS_ChainMaxUnitHeightM` (`modulo2.vba:11908-11912`) |
| Peso máximo | 29 t | `SK_MAX_PESO_KG` (`modulo2.vba:10`) |

El cupo por grupo tiene tres casos (`tms_fg14/modulo2.vba:5496`):

```
' LBS - Capacidad de camion por grupo: 40 Full caja seca (col Y), 28 WALMART, 26 sencillo.
```

En orden: si el equipo de la columna `Y` contiene `Full`, 40. Si el grupo de consolidación
empieza con `WALMART`, 28. En cualquier otro caso, 26.

Un detalle que puede sorprender: **una cadena desconocida cuyo destinatario esté mapeado a un
grupo que empiece con `WALMART` recibe cupo 28.** El cupo depende del grupo, no de la cadena,
en este último paso. Es poco probable pero posible si alguien mapea mal la hoja `Consolida`.

## 3. Reglas de negocio

### Qué comportamientos tiene cada una

| Comportamiento | HEB | Smart | Merco | Merza | Casa Ley | Por omisión |
|---|---|---|---|---|---|---|
| Camas compartidas (`LBS_ChainAllowsSharedCamas`, `6342`) | Sí | Sí | Sí | Sí | Sí | No |
| Tope de altura de 1.60 m (`LBS_ChainEnforcesUnitHeight`, `11895`) | Sí | No | No | No | No | No |
| Altura `AO` desde TI HI (`LBS_ChainComputesTihiAO`, `6308`) | Si está en la hoja | Si está en la hoja | Si está en la hoja | Si está en la hoja | Si está en la hoja | Si está en la hoja |
| Peso `AT` desde TI HI (`LBS_ChainAppliesTihiAT`, `6336`) | No | No | No | No | No | No |
| Sándwich por unidad `W` (`LBS_IsSandwichWChain`, `6353`) | No | No | No | No | No | No |
| Piso de llenado de camión (`LBS_IsTruckMinFillChain`, `4434`) | No | No | No | No | No | No |
| Descarga de peso excedente (`LBS_IsPesoSalvageChain`, `6313`) | No | No | No | No | No | No |
| Apertura de camiones nuevos (`LBS_IsOpenTruckChain`, `6321`) | No | No | No | No | No | No |
| División por orígenes mixtos (`LBS_IsMixedOriginSplitChain`, `13730`) | No | No | No | No | No | No |

**HEB es el caso intermedio.** Tiene camas compartidas y tope de altura de 1.60 m, pero no
tiene piso de llenado ni apertura de camiones. En la práctica se le calcula y respeta la
altura, pero el armado es el genérico. El comentario que lo agrupa
(`tms_fg14/modulo2.vba:11911`) lo pone junto a las cadenas de mezcla de autoservicio:

```
' Chep/Ultrapallet autoservicio mix (+ COMEXTRA): same 1.6 m unit stack cap.
```

Hay una referencia a HEB en otro comentario, sobre el consolidador de restos
(`tms_fg14/modulo3.vba:834-835`):

```
' LBS - Tras agregar los order failures (No planeado), intentar montarlos sobre un camion
' del mismo grupo (Soriana/City Club, Walmart BA/SC, CHEDRAUI) con cupo, como paridad HEB.
```

"Paridad HEB" significa que el remonte de restos se construyó tomando como referencia lo que
ya se hacía con HEB. HEB fue el modelo, no el beneficiario.

**Smart, Merco, Merza y Casa Ley solo tienen camas compartidas.** Es su única regla propia:
dos SKU con altura de caja dentro de 2 mm pueden ir en la misma cama, usando `TI-1` cajas
cuando el clúster es mezclado. Nada más. Ni altura, ni piso, ni sándwich.

Tener camas compartidas sin tope de altura es una combinación deliberada: se permite
optimizar la cama, pero no se limita cuánto se apila, porque no hay una restricción
declarada por el cliente para esas cadenas.

### El comportamiento por omisión, en detalle

Una cadena que no reconoce ningún clasificador:

**Cupo.** 26 de sencillo, 40 si el equipo dice `Full`, 20 por caja `a`/`b`. El cupo se resuelve
por el último paso de `LBS_TruckCapForRow` (`tms_fg14/modulo2.vba:5898`), que delega en el
grupo de consolidación.

**Consolidación.** Su destinatario es su propio grupo salvo que esté en la hoja `Consolida`.
La llave de camión sigue siendo origen, grupo y vigencia, con la misma regla rígida de no
mezclar vigencias.

**Eficiencia.** Solo el gate de `AR` contra la hoja `EFICIENCIA POR CADENA`. Sin piso de
llenado de camión, un camión con 5 tarimas se emite si el fill rate del pedido pasa el
umbral. Es lo más importante de entender del comportamiento por omisión: **la única defensa
contra camiones vacíos es la hoja `EFICIENCIA POR CADENA`.**

**Altura.** `AO` se calcula si la cadena aparece en la hoja `TI HI`, porque
`LBS_ChainComputesTihiAO` (`tms_fg14/modulo2.vba:6308`) devuelve verdadero para cualquier
cadena mapeada. Pero **no se aplica ningún tope**: la altura se informa y no se limita.

**Peso.** 29 t, con marca de `REVISION MANUAL` si se excede. Sin descarga de excedente: la
marca se queda y el operador decide.

**Restos.** Sin sándwich por unidad y sin apertura de camiones nuevos. Lo que no cabe se
queda en `No planeado`.

En resumen: llena hasta 26, informa altura si puede, marca el peso si se pasa, y descarta el
resto. Es conservador por diseño.

### Cadenas que están en el Plan pero no aquí

La macro de entrada reconoce muchas más cadenas que la de salida: Amazon, Mercado Libre,
Rapiturbo, Asturiano, Conasuper, Rabbit, Go Mart, Super Bara, Super Kompras, ToBe y otras
(ver [merged/cadenas/otras-cadenas.md](../../merged/cadenas/otras-cadenas.md)). De esas, en
esta macro solo tienen tratamiento propio las que están en la lista blanca de mayoristas.

Las demás llegan al motor de salida y reciben el comportamiento por omisión. No es un
descuido: significa que su armado no requirió reglas especiales.

Hay fixtures de dos de ellas sin página propia:

- `tms_fg14/asturiano/` (`sample.tsv`, `plan.tsv`) — Asturiano.
- `tms_fg14/smart/` (`sample.tsv`, `equipment.tsv`) — Smart.

## 4. Procedimientos

Ninguna de estas cadenas tiene procedimientos propios. Solo clasificadores:

| Procedimiento | Línea | Qué hace |
|---|---|---|
| `LBS_IsHebChain` | `6287` | Reconoce `HEB` |
| `LBS_IsSmartChain` | `6291` | Reconoce `SMART` |
| `LBS_IsMercoChain` | `6295` | Reconoce `MERCO` |
| `LBS_IsMerzaChain` | `6299` | Reconoce `MERZA` |
| `LBS_IsCasaLeyChain` | `6303` | Reconoce `CASALEY`, con o sin guion |

Y los procedimientos genéricos que las atienden:

| Procedimiento | Línea | Qué hace |
|---|---|---|
| `LBS_TruckCapForRow` | `5847` | Resuelve el cupo por fila |
| `LBS_TruckCapForGroup` | `5497` | El último paso: 40 / 28 / 26 |
| `LBS_ConsolidaKey` | `5463` | La llave de camión |
| `LBS_ChainAllowsSharedCamas` | `6342` | Camas compartidas |
| `LBS_ChainComputesTihiAO` | `6308` | Cálculo de `AO` si la cadena está en `TI HI` |
| `FiltrarPorEficiencia` | `4828` | El gate de `AR` |

## 5. Cómo validarlo

Fixtures:

| Archivo | Contenido |
|---|---|
| `tms_fg14/heb/sample.tsv` | Muestra de HEB |
| `tms_fg14/soriana/heb.tsv` | HEB comparado contra Soriana |
| `tms_fg14/smart/sample.tsv` | Muestra de Smart |
| `tms_fg14/smart/equipment.tsv` | Equipos de Smart |
| `tms_fg14/asturiano/sample.tsv` | Muestra de Asturiano |
| `tms_fg14/asturiano/plan.tsv` | Plan de Asturiano |

Scripts:

| Script | Qué comprueba |
|---|---|
| `scripts/validate_tihi_ao_all_chains.py` | `AO` desde TI HI en todas las cadenas, incluidas estas |
| `scripts/validate_sample_tsv.py` | Estructura general de las muestras |
| `scripts/validate_ae_pallet_mix.py` | La mezcla de tarimas y las camas compartidas |
| `scripts/validate_lbs_regression.py` | Regresión general del motor |

## 6. Problemas conocidos y síntomas

| Síntoma | Causa probable | Dónde mirar |
|---|---|---|
| Camiones muy vacíos que se emiten igual | Sin piso de llenado propio, solo aplica el gate de `AR` | Subir el umbral de la cadena en `EFICIENCIA POR CADENA` |
| Unidades de más de 1.60 m sin marca | Solo HEB tiene tope entre estas cadenas | Es lo esperado. Si el cliente lo pide, agregarla en `LBS_ChainEnforcesUnitHeight` (`modulo2.vba:11895`) |
| `AO` vacío | La cadena no está en la hoja `TI HI` | Agregar el material a la hoja y correr `LBS_SyncTihiSheet` |
| Peso `AT` distinto del catálogo | Ninguna de estas cadenas refresca `AT` | `LBS_ChainAppliesTihiAT` |
| `Smart and Final` sin camas compartidas | El clasificador solo acepta `SMART` exacto | Normalizar el valor en el Plan, o ampliar `LBS_IsSmartChain` |
| Casa Ley con guion sin reconocerse | Debería funcionar: es la única cadena que quita guiones | Revisar espacios extra o acentos |
| Cupo 28 en una cadena desconocida | Su destinatario está mapeado a un grupo que empieza con `WALMART` | Corregir el grupo en la hoja `Consolida` |
| Restos que se quedan en `No planeado` | Sin apertura de camiones nuevos, es el comportamiento esperado | `LBS_IsOpenTruckChain` no las incluye |
| `REVISION MANUAL: peso >29 ton` que no se resuelve | Sin descarga de excedente automática | `LBS_IsPesoSalvageChain` no las incluye. Requiere acción manual |
