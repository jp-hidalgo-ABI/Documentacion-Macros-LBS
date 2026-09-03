[Volver al índice de cadenas](README.md)

# Mayoristas con cupo de 35 y 36

Cinco cadenas mayoristas tienen un tratamiento distinto de todo lo demás: sus Fulls van en
**un camión único sin cajas `a` y `b`**, con cupo de 35 tarimas para los SKU listados en la
hoja `Cadenas 35 Tarimas` y 36 para el resto. Es la única familia donde un Full no se parte.

## 1. Identificación

La lista blanca está en el código, no en la hoja
(`LBS_IsCap35AllowChain`, `tms_fg14/modulo2.vba:5276`):

```
' Client allowlist: mayorista chains that use Cap35 SKUs and/or Full 36 single-truck.
Private Function LBS_IsCap35AllowChain(ByVal chainKey As String) As Boolean
    Select Case chainKey
        Case "TOBEDISTRIBUTIONS", "CONASUPER", "NETO", "CABRITOABARROTERO", "SUPEROFERTAS"
            LBS_IsCap35AllowChain = True
```

Cinco literales, normalizados a mayúsculas y sin espacios
(`LBS_Cap35ChainKey`, `tms_fg14/modulo2.vba:5271`):

| Valor esperado en `Pedidos Surtidos!M` | Llave normalizada |
|---|---|
| `ToBe Distributions` | `TOBEDISTRIBUTIONS` |
| `Conasuper` | `CONASUPER` |
| `Neto` | `NETO` |
| `Cabrito Abarrotero` | `CABRITOABARROTERO` |
| `Super Ofertas` | `SUPEROFERTAS` |

`LBS_IsMayoristaSingleFullTruck` (`tms_fg14/modulo2.vba:5286`) usa la misma lista:

```
' Mayorista Full: one unsuffixed truck (no a/b). Same allowlist as Cap35 sheet chains.
```

**La lista blanca es la puerta.** Agregar una cadena a la hoja `Cadenas 35 Tarimas` no hace
nada si la cadena no está en esta lista del código. Y hay que editarla en **dos** lugares,
porque `PartirTarimasFULL` tiene su propia copia (`PF_IsCap35AllowChain`,
`tms_fg14/modulo5.vba:538`).

`Neto` está en esta lista **y** tiene su propio clasificador. Su otra mitad de tratamiento
está en [oxxo-neto.md](oxxo-neto.md).

Ninguna de las cinco tiene familia en `LBS_ChainFamily`, así que quedan fuera de las listas
por familia.

## 2. Parámetros

| Parámetro | Valor | Origen |
|---|---|---|
| Cupo Full con SKU listado | 35 tarimas | `LBS_CAP35_SHIPMENT` (`modulo2.vba:40`) |
| Cupo Full con SKU no listado | 36 tarimas | `LBS_MAYORISTA_FULL_SHIPMENT` (`modulo2.vba:43`) |
| Cupo caja `a`/`b` | **No aplica**: camión único | `LBS_FullBoxCapForRow` (`modulo2.vba:5513`) |
| Cupo sencillo | 26 tarimas o catálogo | `LBS_SencilloCapForRow` (`modulo2.vba:5841`) |
| Piso de llenado | Ninguno propio | No están en `LBS_IsTruckMinFillChain` |
| Altura máxima | Sin tope | No están en `LBS_ChainEnforcesUnitHeight` |
| Peso máximo | 29 t | `SK_MAX_PESO_KG` (`modulo2.vba:10`) |

Comentarios originales:

```
' SKU-level Full shipment cap (sheet "Cadenas 35 Tarimas": allowlisted cadena + SKU, Viajes a 35=YES).
Private Const LBS_CAP35_SHIPMENT As Long = 35
Private Const LBS_CAP35_SHEET As String = "Cadenas 35 Tarimas"
' Allowlisted mayorista Full (non-listed SKU): single-truck ceiling 36 (no a/b).
Private Const LBS_MAYORISTA_FULL_SHIPMENT As Long = 36
```

Las tres constantes están duplicadas en `modulo5` (`PF_CAP35_SHIPMENT`, `PF_CAP35_SHEET`,
`PF_MAYORISTA_FULL_SHIPMENT`, `tms_fg14/modulo5.vba:6-8`). Cambiar una sin la otra produce
que `SummaryOptimizar` y `PartirTarimasFULL` trabajen con cupos distintos.

## 3. Reglas de negocio

**La hoja `Cadenas 35 Tarimas`.** Cuatro columnas, datos desde la fila 2:

| Columna | Encabezado | Uso |
|---|---|---|
| `A` | `ID Cliente` | **Se ignora** |
| `B` | `Cadena` | Se normaliza y se compara contra la lista blanca |
| `C` | `SKU` | Se canoniza con `SF_CanonMaterial` |
| `D` | `Viajes a 35` | Debe ser `YES`, `SI` o `SÍ` |

El comentario del cargador (`tms_fg14/modulo2.vba:5299`) dice que la primera columna se
ignora:

```
' Load allowlisted cadena+SKU with Viajes a 35=YES (dest col A ignored).
```

Es decir, el cupo de 35 es **por cadena y SKU, no por destino**. Un SKU marcado para
Conasuper lleva 35 en cualquier destino de Conasuper.

`LBS_EnsureCap35` (`tms_fg14/modulo2.vba:5300`) aplica los dos filtros en cascada: primero la
bandera de la columna `D`, después la lista blanca del código. La llave del diccionario es
`cadena|SKU` (`LBS_Cap35DictKey`, `tms_fg14/modulo2.vba:5290`), y
`LBS_IsCap35Row` (`tms_fg14/modulo2.vba:5333`) la busca leyendo la cadena de
`Pedidos Surtidos!M` y el material de `Pedidos Surtidos!AA`.

**Camión único, sin cajas.** Es la regla distintiva. En la resolución del cupo por fila
(`LBS_TruckCapForRow`, `tms_fg14/modulo2.vba:5865-5868`) los mayoristas se evalúan **antes**
que las cadenas que usan media caja:

```
If LBS_IsMayoristaSingleFullTruck(mVal) Then
    ' Single truck: always shipment cap (35 Cap35 SKU / 36 other), never caja half.
    LBS_TruckCapForRow = LBS_CatalogFullShipmentCap(ws, r)
```

Y si un folio llega con sufijo, se le quita: `LBS_CollapseMayoristaFullSuffixFolios`
(`tms_fg14/modulo2.vba:15112`), con este comentario:

```
' Mayorista Full: strip a/b suffixes so one AD = one truck (no caja half-cap).
```

Un folio, un camión. Sin `a` ni `b`.

**Cómo se decide entre 35 y 36.** `LBS_CatalogFullShipmentCap`
(`tms_fg14/modulo2.vba:5753-5762`) resuelve en este orden:

```
' Cap35 SKU (sheet Viajes a 35=YES): hard max 35.
' Other allowlisted mayorista Full: ceiling 36 (single truck, no a/b).
If LBS_IsCap35Row(ws, r) Then
    If cap > LBS_CAP35_SHIPMENT Or cap < 1 Then cap = LBS_CAP35_SHIPMENT
ElseIf LBS_IsMayoristaSingleFullTruck(mVal) Then
    If InStr(1, CStr(ws.Cells(r, "Y").Value), "Full", vbTextCompare) > 0 Then
        If cap > LBS_MAYORISTA_FULL_SHIPMENT Or cap < 1 Then cap = LBS_MAYORISTA_FULL_SHIPMENT
```

Los dos son **topes**, no valores fijos. Si el catálogo Mode Mix declara 30 para el carril, se
usan 30. Si declara 40, se topa a 35 o 36 según corresponda.

Nótese que el 36 exige que la columna `Y` diga `Full`, y el 35 no. Un SKU listado lleva 35
aunque el equipo venga en blanco.

**Remonte de sobrantes.** El comentario de `LBS_PackCap35Leftovers`
(`tms_fg14/modulo2.vba:14344-14346`) describe el empaque:

```
' Mayorista Full: pack No planeado leftovers into trucks at Catalogo Mode Mix max
' (Cap35 SKU=35, other mayorista Full=36). Same plant|dest; Cap35 SKUs never mix onto 36.
' Single unsuffixed truck per AD (collapse a/b). Skips OXXO.
```

Tres restricciones importantes:

- Solo se empaca dentro de la misma planta y destino.
- **Un SKU de cupo 35 nunca se mezcla en un camión de 36.** Si se mezclara, el camión de 36
  tendría producto que solo puede ir a 35, y el cupo dejaría de significar algo.
- Se omite OXXO, que tiene su propio empaque.

`LBS_MayoristaFullPackCap` (`tms_fg14/modulo2.vba:14314`) calcula el cupo objetivo del
empaque. Tiene un detalle curioso, documentado en el propio código
(`tms_fg14/modulo2.vba:14331`):

```
' Temporarily treat as Full so CatalogFullShipmentCap applies mayorista ceiling.
ySaved = ws.Cells(r, "Y").Value
ws.Cells(r, "Y").Value = "Full encortinado"
cap = LBS_CatalogFullShipmentCap(ws, r)
ws.Cells(r, "Y").Value = ySaved
```

Las filas `No planeado` suelen traer la columna `Y` en blanco, y el tope de 36 exige que diga
`Full`. La función escribe `Full encortinado` temporalmente, calcula el cupo y restaura el
valor original. Es un rodeo, pero está documentado y es intencional.

`LBS_Cap35MergeIntoFolioProduct` (`tms_fg14/modulo2.vba:14822`) une el producto remontado
dentro del folio destino.

**Sin altura, sin camas, sin sándwich.** Ninguna de las cinco está en
`LBS_ChainEnforcesUnitHeight` (`tms_fg14/modulo2.vba:11895`),
`LBS_ChainAllowsSharedCamas` (`6342`) ni `LBS_IsSandwichWChain` (`6353`). Son cadenas de
tarima completa: no hay restos que apilar.

## 4. Procedimientos

### En `modulo2` (armado y consolidación)

| Procedimiento | Línea | Qué hace |
|---|---|---|
| `LBS_Cap35ChainKey` | `5271` | Normaliza la cadena para la búsqueda |
| `LBS_IsCap35AllowChain` | `5276` | La lista blanca de cinco cadenas |
| `LBS_IsMayoristaSingleFullTruck` | `5286` | Camión único sin cajas, misma lista |
| `LBS_Cap35DictKey` | `5290` | Llave `cadena|SKU` |
| `LBS_EnsureCap35` | `5300` | Carga la hoja `Cadenas 35 Tarimas` en memoria |
| `LBS_IsCap35Row` | `5333` | Decide si una fila tiene cupo 35 |
| `LBS_CatalogFullShipmentCap` | `5703` | Resuelve el cupo, con los topes de 35 y 36 |
| `LBS_MayoristaFullPackCap` | `14314` | Cupo objetivo del empaque de sobrantes |
| `LBS_PackCap35Leftovers` | `14347` | Empaca las filas `No planeado` en camiones |
| `LBS_Cap35MergeIntoFolioProduct` | `14822` | Une el producto remontado en el folio |
| `LBS_CollapseMayoristaFullSuffixFolios` | `15112` | Quita los sufijos `a`/`b` |

### En `modulo5` (división de Fulls)

| Procedimiento | Línea | Qué hace |
|---|---|---|
| `PF_Cap35ChainKey` | `533` | Copia del normalizador |
| `PF_IsCap35AllowChain` | `538` | Copia de la lista blanca |
| `PF_IsMayoristaSingleFullTruck` | `547` | Copia del clasificador de camión único |
| `PF_Cap35DictKey` | `551` | Copia de la llave |
| `PF_EnsureCap35` | `561` | Copia del cargador |
| `PF_IsCap35Row` | `593` | Copia del decisor por fila |

La duplicación completa entre módulos es deliberada (cada módulo tiene su propia caché) pero
es la fuente más probable de inconsistencias. **Cualquier cambio va en los dos.**

La fase de `SummaryOptimizar` correspondiente es `filtrar:pack_cap35`
(`tms_fg14/modulo2.vba:5056`).

## 5. Cómo validarlo

Fixtures:

| Archivo | Contenido |
|---|---|
| `tms_fg14/conasuper/Cadenas 35 Tarimas.tsv` | La hoja de la lista blanca de SKU |
| `tms_fg14/conasuper/sample.tsv` | Muestra de `Pedidos Surtidos` |
| `tms_fg14/neto/sample.tsv` | Muestra de Neto |

El archivo `Cadenas 35 Tarimas.tsv` es el fixture de referencia de la hoja: su primer renglón
de datos es `400101621 / Super Ofertas / 3007678 / YES`, que sirve para verificar el formato
esperado de las cuatro columnas.

Scripts:

| Script | Qué comprueba |
|---|---|
| `scripts/validate_partir_fulles_blocks.py` | Que los Fulls de mayorista no se partan en cajas |
| `scripts/validate_surtido_partir.py` | El resultado de la división sobre `Pedidos Surtidos` |
| `scripts/validate_sample_tsv.py` | Estructura general de las muestras |

## 6. Problemas conocidos y síntomas

| Síntoma | Causa probable | Dónde mirar |
|---|---|---|
| Un SKU marcado `YES` que no llega a 35 | La cadena no está en la lista blanca del código | `LBS_IsCap35AllowChain` (`modulo2.vba:5276`). La hoja sola no basta |
| Cupos distintos entre `SummaryOptimizar` y `PartirTarimasFULL` | Se cambió una constante en `modulo2` y no en `modulo5` | Comparar `modulo2.vba:40-43` con `modulo5.vba:6-8` |
| Folios de mayorista con sufijo `a` o `b` | El colapso de sufijos no corrió | `LBS_CollapseMayoristaFullSuffixFolios` |
| Camión de 36 con producto de cupo 35 | La restricción de no mezclar falló | El comentario dice `Cap35 SKUs never mix onto 36`. Es una regresión |
| Cupo menor a 35 en un SKU listado | El catálogo Mode Mix declara menos para ese carril | Es correcto: 35 es un tope, no un mínimo |
| Sobrantes `No planeado` de mayorista sin remontar | La columna `Y` en blanco y el rodeo de `Full encortinado` no aplicó | `LBS_MayoristaFullPackCap` (`modulo2.vba:14331`) |
| Cadena nueva que el cliente pidió agregar y no funciona | Falta editarla en los dos módulos | `LBS_IsCap35AllowChain` **y** `PF_IsCap35AllowChain` |
| `Viajes a 35` con otro texto (por ejemplo `X`) | Solo se aceptan `YES`, `SI` y `SÍ` | `LBS_EnsureCap35` (`modulo2.vba:5322`) |
| `Neto` con reglas que no coinciden con esta página | Neto tiene tratamiento doble | Ver también [oxxo-neto.md](oxxo-neto.md) |
