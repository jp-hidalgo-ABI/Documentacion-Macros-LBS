[Volver al índice de la macro de salida](../README.md)

# Reglas por cadena — macro de salida

El motor de armado tiene una capa común (documentada en
[04-motor-armado-cargas.md](../04-motor-armado-cargas.md)) y encima una capa por cadena que
cambia cupos, pisos, alturas y, en varios casos, el algoritmo completo de empaque. Esta
carpeta describe esa segunda capa.

Cada página sigue la misma estructura, para que se puedan comparar entre sí:

1. **Identificación** — los literales exactos con los que el clasificador reconoce la cadena.
2. **Parámetros** — cupos, pisos, altura y peso, con el comentario original del código.
3. **Reglas de negocio** — en prosa, sin código.
4. **Procedimientos** — nombre, línea y qué hace.
5. **Cómo validarlo** — los fixtures y scripts que ya existen en el repositorio.
6. **Problemas conocidos** — qué se ve en `AG` y `AV` cuando esta cadena falla.

## Índice

| Página | Cadenas que cubre |
|---|---|
| [walmart.md](walmart.md) | `Walmart BA`, `Walmart SC` |
| [soriana-city-club.md](soriana-city-club.md) | `Soriana`, `City Club` (familia `CLUBCITY`) |
| [chedraui.md](chedraui.md) | `Chedraui` |
| [la-comer.md](la-comer.md) | `La Comer` |
| [oxxo-neto.md](oxxo-neto.md) | `OXXO`, `Neto` |
| [comextra.md](comextra.md) | `Comextra` |
| [alsuper-gomart-europea.md](alsuper-gomart-europea.md) | `Alsuper`, `Go Mart`, `Europea` / `La Europea` |
| [sams.md](sams.md) | `Sams` |
| [mayoristas-cap35.md](mayoristas-cap35.md) | `ToBe Distributions`, `Conasuper`, `Neto`, `Cabrito Abarrotero`, `Super Ofertas` |
| [otras-cadenas.md](otras-cadenas.md) | `HEB`, `Smart`, `Merco`, `Merza`, `Casa Ley` y el comportamiento por omisión |

## Cómo se identifica una cadena

Todo sale de `Pedidos Surtidos!M`. Los clasificadores normalizan igual en todos los casos:
mayúsculas, sin espacios, comparación exacta. `LBS_IsSorianaChain`
(`tms_fg14/modulo2.vba:6279`) es representativo:

```
LBS_IsSorianaChain = (Replace(UCase$(Trim$(CStr(m))), " ", "") = "SORIANA")
```

La consecuencia práctica de la comparación exacta: **si el Plan trae la cadena escrita
distinto, la macro no la reconoce y la fila cae en el comportamiento por omisión.** No hay
coincidencia parcial ni tolerancia a variantes. `City Fresko`, por ejemplo, no es
`CITYCLUB`.

Dos clasificadores agrupan cadenas en familias (`LBS_ChainFamily`,
`tms_fg14/modulo2.vba:5198`):

| Familia | Literales |
|---|---|
| `CLUBCITY` | `SORIANA`, `CITYCLUB` |
| `WALMART` | `WALMARTBA`, `WALMARTSC` |
| `CHEDRAUI` | `CHEDRAUI` |
| `ALSUPER` | `ALSUPER` |
| `GOMART` | `GOMART` |
| `EUROPEA` | `EUROPEA`, `LAEUROPEA` |

`Casa Ley` es el único clasificador que además quita guiones
(`LBS_IsCasaLeyChain`, `tms_fg14/modulo2.vba:6303`), así que acepta `CASA LEY`, `CASA-LEY` y
`CASALEY`.

## Tabla comparativa

Cupos en tarimas. El piso es fracción del cupo que aplica; el número entre paréntesis es el
piso efectivo.

| Cadena | Cupo sencillo | Cupo Full | Cupo caja `a`/`b` | Piso de llenado | Altura máx. | Peso máx. |
|---|---|---|---|---|---|---|
| Walmart BA/SC | 28 | 40 | 20 | 40 % (12 de 28) | 1.6 m | 29 t |
| Soriana / City Club | 26 | catálogo | catálogo / 2 | 70 % (19 de 26) | 1.6 m | 29 t |
| Chedraui | 26 | catálogo | catálogo / 2 | 80 % (21 de 26) | 1.6 m | 29 t |
| La Comer | 26 | catálogo | catálogo / 2 | 80 % (21 de 26) | 1.6 m | 29 t |
| OXXO | 24 / 22 / 28 | 36 (tope) | 18 | 90 % (33 de 36) | sin tope | 29 t |
| Neto | catálogo | 35 / 36 | catálogo / 2 | 90 % | sin tope | 29 t |
| Comextra | 26 | 40 | 20 | 90 % (36 / 18 / 24) | 1.6 m | 29 t |
| Alsuper / Go Mart / Europea | 26 (catálogo) | 40 (catálogo) | 20 | 40 % | 1.6 m | 52.5 t |
| Sams | 26 | — | — | — | sin tope | 29 t |
| Mayoristas Cap35 | 26 | 35 o 36, camión único | no aplica | — | sin tope | 29 t |
| HEB | 26 | 40 | 20 | — | 1.6 m | 29 t |
| Smart / Merco / Merza / Casa Ley | 26 | 40 | 20 | — | sin tope | 29 t |
| Por omisión | 26 | 40 | 20 | — | sin tope | 29 t |

Los tres cupos de sencillo de OXXO dependen del equipo y del destino: 24 general, 22 para
`Z4290_OXX` y 28 para el destinatario `400101621` (Monterrey). Detalle en
[oxxo-neto.md](oxxo-neto.md).

"Catálogo" significa que el cupo lo dicta `Catalogo Mode Mix` por carril, con 26 (sencillo) o
40 (Full) como respaldo cuando el carril no aparece.

### Cómo se resuelve el cupo, en orden

`LBS_TruckCapForRow` (`tms_fg14/modulo2.vba:5847`) evalúa en este orden y **el primero que
aplica gana**:

1. Sencillo de OXXO, Neto, Alsuper, Go Mart o Europea: cupo de catálogo, respaldo 26.
2. Full de mayorista de la lista blanca: cupo de embarque (35 o 36), **nunca** media caja.
3. Full de OXXO, Neto, Soriana, City Club, Go Mart, Alsuper, o con sufijo de camión, o SKU
   de la hoja `Cadenas 35 Tarimas`: si tiene sufijo `a`/`b`, media caja; si no, cupo de
   embarque completo.
4. Walmart: 28 si es sencillo, 40 si es Full. **Siempre por cadena**, nunca por grupo.
5. Cualquier otro: cupo por grupo de consolidación (`LBS_TruckCapForGroup`,
   `tms_fg14/modulo2.vba:5497`), que da 40 si el equipo dice `Full`, 28 si el grupo empieza
   con `WALMART` y 26 en cualquier otro caso.

El paso 4 existe por un defecto concreto, documentado en el código
(`tms_fg14/modulo2.vba:5887-5889`):

```
' Walmart: always 28 (Sencillo) / 40 (Full) by chain — never metro 26 when
' Consolida group is a bare destinatario id (not "WALMART*").
' Alsuper/GoMart/Europea use Mode Mix above (26 S / 40 F) — not Walmart 28.
```

Un destinatario de Walmart que no esté mapeado en la hoja `Consolida` tiene como grupo su
propio identificador, que no empieza con `WALMART`, así que el paso 5 le daría 26 en lugar de
28. El paso 4 lo evita mirando la cadena directamente.

## La llave de consolidación

`LBS_ConsolidaKey` (`tms_fg14/modulo2.vba:5463`) decide qué filas pueden compartir camión:

```
origen (col L) | grupo de consolidación [ | R/C de La Comer ] | vigencia (col R)
```

El comentario declara la regla más rígida de todo el motor
(`tms_fg14/modulo2.vba:5461-5462`):

```
' + vigencia (col R). Hard rule: never mix distinct Plan vigencias on one truck.
```

Dos vigencias distintas **nunca** comparten camión. No hay excepción por cadena, y todo lo
que parece contradecirlo (como la segunda pasada de Soriana / City Club) en realidad opera
sobre un carril sin vigencia y después reparte, sin mezclar dentro de un camión.

La Comer es la única cadena que agrega un cuarto componente: la marca `R` o `C` de la
columna `AE`. Ver [la-comer.md](la-comer.md).

## Qué comportamientos aplica cada cadena

Las banderas transversales, que se combinan de forma distinta que los cupos.

| Comportamiento | Función | Cadenas |
|---|---|---|
| Calcula altura `AO` desde TI HI | `LBS_ChainComputesTihiAO` (`modulo2.vba:6308`) | Todas las que aparezcan en la hoja `TI HI` |
| Aplica tope de altura de 1.6 m | `LBS_ChainEnforcesUnitHeight` (`modulo2.vba:11895`) | Walmart, La Comer, Soriana, City Club, HEB, Alsuper, Chedraui, Go Mart, Europea, Comextra |
| Refresca el peso `AT` desde TI HI | `LBS_ChainAppliesTihiAT` (`modulo2.vba:6336`) | Walmart, Soriana, City Club, Alsuper, Go Mart, Europea |
| Permite camas compartidas | `LBS_ChainAllowsSharedCamas` (`modulo2.vba:6342`) | Walmart, Soriana, City Club, La Comer, HEB, Alsuper, Go Mart, Europea, Chedraui, Smart, Merco, Merza, Casa Ley, Comextra |
| Cuenta sándwich por unidad `W` | `LBS_IsSandwichWChain` (`modulo2.vba:6353`) | Walmart, Alsuper, Go Mart, Europea, Soriana, City Club, Comextra |
| Tiene piso de llenado de camión | `LBS_IsTruckMinFillChain` (`modulo2.vba:4434`) | Walmart, Alsuper, Go Mart, Europea, Soriana, City Club, OXXO, Comextra |
| Descarga peso excedente | `LBS_IsPesoSalvageChain` (`modulo2.vba:6313`) | Familias `WALMART`, `CLUBCITY`, `ALSUPER`, `GOMART`, `EUROPEA` |
| Abre camiones nuevos | `LBS_IsOpenTruckChain` (`modulo2.vba:6321`) | Las anteriores más Comextra |
| Divide folios con orígenes mixtos | `LBS_IsMixedOriginSplitChain` (`modulo2.vba:13730`) | Chedraui, OXXO, Neto, Walmart |

Dos de estas listas tienen ausencias deliberadas.

**La Comer no cuenta sándwich por `W`.** El comentario lo dice
(`tms_fg14/modulo2.vba:6350-6352`):

```
' Cadenas con conteo sandwich por unidad AI (W=1 en ancla, W=0 si bare full paga el slot):
' Walmart BA/SC, Alsuper, Go Mart, Europea, Soriana, City Club, COMEXTRA.
' LA COMER stays outside — FixLaComerSandwichAnchors owns its ANCHOR/W recovery.
```

La Comer tiene su propio mecanismo (`LBS_FixLaComerSandwichAnchors`,
`tms_fg14/modulo2.vba:13301`) y meterla en el mecanismo genérico rompería el suyo.

**Chedraui permite camas compartidas y tope de altura, pero no refresca `AT`.** Es decir, se
le calcula la altura pero el peso viene de donde venga, no del catálogo TI HI.

## El comportamiento por omisión

Una cadena que ningún clasificador reconoce:

- Cupo 26 de sencillo, 40 si el equipo dice `Full`, y 20 por caja `a`/`b`.
- Sin piso de llenado propio: solo el gate de `AR` contra `EFICIENCIA POR CADENA`.
- Sin tope de altura: `AO` se calcula solo si la cadena está en la hoja `TI HI`.
- Peso máximo de 29 t.
- Sin camas compartidas ni conteo de sándwich por unidad.
- Sin remonte a camiones nuevos: lo que no cabe se queda en `No planeado`.

En la práctica es un armado conservador: llena hasta 26 y descarta el resto. Ver
[otras-cadenas.md](otras-cadenas.md).

## Cómo se valida

Los fixtures viven en `tms_fg14/<cadena>/*.tsv` y los scripts en `scripts/validate_*.py`.
Cada página de cadena lista los suyos en la sección "Cómo validarlo". El inventario general:

| Carpeta de fixtures | Scripts relacionados |
|---|---|
| `tms_fg14/walmart/` | `validate_walmart_restos.py`, `validate_walmart_heights.py`, `validate_walmart_conf_unit.py`, `validate_walmart_exc_groupid.py` |
| `tms_fg14/soriana/`, `tms_fg14/sample pc01 soriana/` | `validate_soriana_sencillo.py`, `validate_soriana_weight_tihi.py` |
| `tms_fg14/chedraiu/` | `validate_chedraui_restos.py` |
| `tms_fg14/lacomer/` | `validate_lacomer_rc.py`, `validate_lacomer_restos.py`, `validate_lacomer_weight_tihi.py`, `validate_tihi_pkg_lacomer.py` |
| `tms_fg14/oxxo/` | `validate_tms_oxxo.py`, `validate_oxxo_sample_z.py`, `validate_oxxo_body_type.py` |
| `tms_fg14/comextra/` | `validate_comextra_merge.py` |
| `tms_fg14/conasuper/`, `tms_fg14/neto/` | `validate_partir_fulles_blocks.py` |
| `tms_fg14/alsuper/`, `tms_fg14/gomart/` | `validate_plant_sencillo_chains.py`, `validate_7e_itempackage_armado.py` |
| `tms_fg14/sams/`, `tms_fg14/heb/`, `tms_fg14/smart/`, `tms_fg14/asturiano/` | `validate_sample_tsv.py`, `validate_tihi_ao_all_chains.py` |

El nombre de la carpeta de Chedraui está escrito `chedraiu`. Es un error de dedo que ya
quedó en los scripts, así que se conserva.
