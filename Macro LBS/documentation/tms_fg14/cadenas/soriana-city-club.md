[Volver al índice de cadenas](README.md)

# Soriana y City Club (familia `CLUBCITY`)

Son dos cadenas que el motor trata como una sola familia porque comparten CEDIS y viajan
juntas. El rasgo distintivo es la **segunda pasada cruzando vigencias**: un mecanismo para
compartir sobrantes entre camiones del mismo carril cuando el empaque por vigencia dejó
camiones a medias, sin romper la regla de que un camión no mezcla vigencias.

## 1. Identificación

Dos clasificadores exactos:

| Función | Línea | Literal |
|---|---|---|
| `LBS_IsSorianaChain` | `tms_fg14/modulo2.vba:6279` | `SORIANA` |
| `LBS_IsCityClubChain` | `tms_fg14/modulo2.vba:6283` | `CITYCLUB` |

Y la familia (`LBS_ChainFamily`, `tms_fg14/modulo2.vba:5201-5202`):

```
If u = "SORIANA" Or u = "CITYCLUB" Then
    LBS_ChainFamily = "CLUBCITY"
```

`City Fresko` **no** entra: normalizado da `CITYFRESKO`, que no coincide con ninguno de los
dos literales. En la macro de entrada sí tiene tratamiento propio (ver
[merged/cadenas/soriana-city-club.md](../../merged/cadenas/soriana-city-club.md)); aquí cae
en el comportamiento por omisión.

Los grupos de la tabla embebida (`LBS_LoadConsDefaults`, `tms_fg14/modulo2.vba:5354-5362`)
son cuatro, por ciudad:

| Grupo | Destinatarios |
|---|---|
| `CLUBCITYGDL` | `400006407`, `400001207` |
| `CLUBCITYMTY` | `400006526`, `400093198` |
| `CLUBCITYQRO` | `400087899`, `400053595` |
| `CLUBCITYTUL` | `400006528`, `400001205` |

El comentario de la tabla lo confirma (`tms_fg14/modulo2.vba:5349-5351`):

```
' LBS - Tabla por defecto Destinatario -> Grupo (semilla para la hoja "Consolida" y
' fallback cuando la hoja no existe). Incluye CLUBCITY (Soriana + City Club) y WALMART
' (Bodega Aurrera + SuperCenter) con sus excepciones (7482 MX / 400087097 -> NORTE).
```

## 2. Parámetros

| Parámetro | Valor | Origen |
|---|---|---|
| Cupo sencillo | 26 tarimas | `LBS_METRO_TRUCK_CAP` (`modulo2.vba:19`) |
| Cupo Full | Catálogo Mode Mix, normalmente 40 | `LBS_CatalogFullShipmentCap` (`modulo2.vba:5703`) |
| Cupo caja `a`/`b` | Cupo de embarque entre 2, normalmente 20 | `LBS_FullBoxCapForRow` (`modulo2.vba:5513`) |
| Piso de llenado | 70 % del cupo = 19 de 26 | `LBS_CLUBCITY_MIN_FILL` (`modulo2.vba:67`) |
| Altura máxima de unidad | 1.60 m | `LBS_ChainMaxUnitHeightM` (`modulo2.vba:11908-11912`) |
| Peso máximo | 29 t | `SK_MAX_PESO_KG` (`modulo2.vba:10`) |

Comentarios originales:

```
' LBS - capacidad de camion (tarimas) para montar restos por metro (Soriana/City Club).
Private Const LBS_METRO_TRUCK_CAP As Long = 26
```

```
' LBS - CLUBCITY (Soriana + City Club): piso de llenado post-consolidacion (70% del cap).
Private Const LBS_CLUBCITY_MIN_FILL As Double = 0.7
```

El nombre `METRO` de la constante viene de aquí: el cupo de 26 se introdujo para el armado
por metro de Soriana y City Club, y después se convirtió en el respaldo genérico de toda la
macro.

La altura de 1.60 m no tiene constante propia. `LBS_ChainMaxUnitHeightM`
(`tms_fg14/modulo2.vba:11908-11912`) reutiliza la de Walmart, con este comentario:

```
' Chep/Ultrapallet autoservicio mix (+ COMEXTRA): same 1.6 m unit stack cap.
LBS_ChainMaxUnitHeightM = LBS_WALMART_MAX_HEIGHT_M
```

Es decir, es el mismo límite físico porque es el mismo tipo de tarima y de caja de camión, no
un acuerdo comercial distinto.

## 3. Reglas de negocio

**Solo sencillo en la macro de entrada, ambos modos aquí.** La macro de entrada restringe
Soriana y City Club a sencillo; en esta macro el motor sí maneja Fulls, porque LBS los puede
armar y hay que procesarlos.

**Carril sin vigencia.** Además de la llave de consolidación normal (que incluye vigencia),
esta familia tiene una llave de carril que la omite (`LBS_ClubCityLaneKey`,
`tms_fg14/modulo2.vba:5478`):

```
' ClubCity lane without vigencia: plant|CLUBCITY* (MTY/GDL/QRO/TUL).
' Pass-1 trucks still use LBS_ConsolidaKey (includes vigencia); pass-2 merges by lane.
```

La llave de carril exige que el grupo empiece con `CLUBCITY`. Un destinatario de Soriana
mapeado a un grupo con otro nombre no participa de la segunda pasada.

**Segunda pasada cruzando vigencias.** Es el mecanismo distintivo. El comentario
(`tms_fg14/modulo2.vba:15205-15208`) enumera las cuatro restricciones:

```
' ClubCity pass-2: after same-vigencia packing, share underfilled Programado and
' No planeado leftovers across Plan vigencia on the same plant|CLUBCITY* lane.
' Never remix full (cap) trucks. Never mix Destinatario on the same AI/pallet.
' Never rewrite col R. Same-dest + two vigencias stay split (pass-1 rule).
```

En prosa: después de empacar por vigencia, si quedan camiones a medias y sobrantes en `No
planeado`, se reparten entre los camiones del mismo carril aunque sean de vigencias
distintas. Pero **nunca** se toca un camión que ya está lleno, **nunca** se mezclan
destinatarios en la misma tarima física, **nunca** se reescribe la vigencia de la columna
`R`, y dos vigencias del mismo destinatario siguen separadas.

La última restricción es la que evita que la segunda pasada sea una puerta trasera para
mezclar vigencias. Lo que se comparte son sobrantes entre carriles, no producto dentro de un
mismo camión.

**Identificador de unidad limpio al remontar.** Cuando se remonta una fila,
`LBS_ClubCityTagRemountDestPureAI` (`tms_fg14/modulo2.vba:15678`) le asigna un identificador
`AI` nuevo. El comentario dice por qué:

```
' Assign a fresh AI to remounted rows so they never share a pallet with another dest.
```

Sin esto, una tarima física podría terminar con producto de dos destinatarios, que es
justamente lo que la segunda pasada promete no hacer.

**Unión de Fulls sin sufijo.** Dos folios Full sin sufijo en la misma llave de consolidación
se unen si juntos llegan cerca del cupo de embarque. El comentario trae el caso y la
corrección (`tms_fg14/modulo2.vba:15707-15711`):

```
' ClubCity: two+ unsuffixed Full Programado ADs on the same ConsolidaKey with
' combined TW within 2 of shipCap (40) - merge onto the largest AD, then
' SplitUnsuffixed emits a/b (P-1120=20 + P-1121=18 -> one Full 40).
' Do NOT merge weak orphans (20+6=26) - that created broken a/b (20+6) when
' SplitUnsuffixed always ran after this sub.
```

La condición "a menos de 2 del cupo" no es arbitraria: solo se unen folios que juntos den un
Full completo. Unir un 20 con un 6 producía un Full partido en cajas de 20 y 6, que no sirve.

**Rellenado de sobrantes débiles.** Un Full sin sufijo, o una caja `a`/`b` cuya hermana no
existe, se considera sobrante débil (`LBS_ClubCityFullIsWeakLeftover`,
`tms_fg14/modulo2.vba:3959`) y se rellena hacia el cupo de embarque. El comentario
(`tms_fg14/modulo2.vba:3982-3984`):

```
' Fill Full weak leftovers (orphan box / unsuffixed) toward shipment cap
' (ClubCity CatalogFullShipmentCap 40 / OXXO 36). Pack two half-boxes into one Full;
' donors may empty. Complete a/b pairs are skipped.
```

Los pares `a`/`b` completos se dejan en paz. Solo se toca lo que está incompleto.

**Unión de restos del mismo SKU.** Varias filas de resto del mismo pedido y material se unen
en una sola por folio, si la altura lo permite. El comentario tiene una advertencia de
rendimiento importante (`tms_fg14/modulo2.vba:18099-18104`):

```
' ClubCity (Soriana + City Club): coalesce same pedido|material restos onto one
' Programado row per folio when height still fits. Same AD+AI always merges;
' same folio / different AI only when allowCrossAI and trial AP <= unit cap.
' Never merges across folios. SummaryOK must pass allowCrossAI:=False - cross-AI
' height trials (SharedHeight/TIHI) hard-close Excel on large OK sheets; Fallos
' runs the full pass. Returns donors cleared.
```

Durante `SummaryOK` la unión solo opera dentro de la misma tarima física. La versión
completa, que cruza tarimas y prueba alturas, corre en `SummaryFallo`, porque hacerla en
`SummaryOK` sobre una hoja grande cierra Excel.

**Folios multi-destino protegidos.** Un folio del carril con dos o más destinatarios está
protegido contra divisiones (`LBS_ClubCityIsProtectedMultiDestFolio`,
`tms_fg14/modulo2.vba:13919`): es un multi-stop legítimo de la segunda pasada, no un error de
mezcla.

**Peso desde el catálogo.** La familia refresca el peso `AT` desde `Peso Bruto de material`
del catálogo TI HI (`LBS_ChainAppliesTihiAT`, `tms_fg14/modulo2.vba:6336`).

## 4. Procedimientos

| Procedimiento | Línea | Qué hace |
|---|---|---|
| `LBS_IsSorianaChain` | `6279` | Reconoce `SORIANA` |
| `LBS_IsCityClubChain` | `6283` | Reconoce `CITYCLUB` |
| `LBS_HasClubCity` | `6271` | Bandera de presencia de la familia en la hoja |
| `LBS_ClubCityLaneKey` | `5478` | Llave de carril sin vigencia: `origen|CLUBCITY*` |
| `LBS_ClubCityFullIsWeakLeftover` | `3959` | Detecta Full sin sufijo o caja huérfana |
| `LBS_CannibalizeClubCityFullShipCap` | `3985` | Rellena los sobrantes débiles hacia el cupo de embarque |
| `LBS_ClubCityCrossVigenciaSecondPass` | `15209` | La segunda pasada que cruza vigencias por carril |
| `LBS_ClubCityTagRemountDestPureAI` | `15678` | Identificador `AI` nuevo para las filas remontadas |
| `LBS_ClubCityMergeUnsuffixedFullSameCkey` | `15713` | Une Fulls sin sufijo de la misma llave |
| `LBS_ClubCityCoalesceSameSkuRestos` | `18105` | Une restos del mismo pedido y material |
| `LBS_ClubCityIsProtectedMultiDestFolio` | `13919` | Protege los folios multi-destino del carril |

La familia también usa buena parte de la maquinaria de Walmart: el cálculo de alturas
(`LBS_RecalcWalmartHeightsAOAP`), la asignación de unidades (`LBS_WalmartAssignTarimaUnitAI`)
y el refresco de peso (`LBS_WalmartRefreshAllTihiAT`). El prefijo `Walmart` en esos nombres
es histórico, no una restricción de cadena.

La fase de `SummaryOptimizar` correspondiente es `filtrar:clubcity_cross_vig`
(`tms_fg14/modulo2.vba:4891`), y la de unión de restos es `coalesce_clubcity_sku`
(`tms_fg14/modulo2.vba:1492`) dentro de `SummaryOK`.

## 5. Cómo validarlo

Fixtures en `tms_fg14/soriana/`:

| Archivo | Contenido |
|---|---|
| `sample.tsv` | Muestra de `Pedidos Surtidos` |
| `plan.tsv` | Plan de prueba |
| `cityclub baja.tsv` | Caso de City Club en Baja |
| `pallet contaier.tsv` | Muestra de `Pallet Container Association` |
| `ti hi over 1.6.tsv` | Caso de unidades que exceden 1.60 m |
| `heb.tsv` | Comparación contra HEB |

Y en `tms_fg14/sample pc01 soriana/cityclub.tsv`, una muestra desde la planta `PC01`.

Scripts:

| Script | Qué comprueba |
|---|---|
| `scripts/validate_soriana_sencillo.py` | Que Soriana y City Club vayan en sencillo donde corresponde |
| `scripts/validate_soriana_weight_tihi.py` | El peso `AT` refrescado desde el catálogo TI HI |
| `scripts/validate_multi_origen_consolidation.py` | La consolidación con orígenes distintos |
| `scripts/validate_consolidar_split_mismo_pedido.py` | Divisiones dentro del mismo pedido |

## 6. Problemas conocidos y síntomas

| Síntoma | Causa probable | Dónde mirar |
|---|---|---|
| `Descartado por baja eficiencia` en camiones de 15 a 18 tarimas | Quedaron bajo el piso de 19 | `LBS_CLUBCITY_MIN_FILL`; revisar si la segunda pasada pudo compartir sobrantes |
| Camiones a medias en el mismo carril y con vigencias distintas | La segunda pasada no operó | Verificar que el grupo de `Consolida` empiece con `CLUBCITY`: `LBS_ClubCityLaneKey` lo exige |
| Una tarima con dos destinatarios | Falló la asignación de `AI` al remontar | `LBS_ClubCityTagRemountDestPureAI`; es un defecto, no un comportamiento esperado |
| Full partido en cajas desiguales (por ejemplo 20 y 6) | Se unieron dos sobrantes débiles | Es exactamente lo que el comentario de `LBS_ClubCityMergeUnsuffixedFullSameCkey` dice que ya no debe pasar. Si aparece, es una regresión |
| `City Fresko` con cupo 26 y sin reglas de familia | No lo reconoce el clasificador | Es el comportamiento esperado en esta macro. Ver [otras-cadenas.md](otras-cadenas.md) |
| Restos del mismo SKU sin unir después de `SummaryOK` | La unión cruzando tarimas solo corre en `SummaryFallo` | Es deliberado, por rendimiento. Correr `SummaryFallo` |
| `REVISION MANUAL: altura >1.6` | Igual que en Walmart: catálogo TI HI desalineado del armado real | Hoja `TI HI` para ese material |
