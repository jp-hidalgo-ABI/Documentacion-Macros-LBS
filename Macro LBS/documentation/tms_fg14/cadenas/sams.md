[Volver al índice de cadenas](README.md)

# Sams

Sams es la cadena con el tratamiento más corto del motor: tres procedimientos. Pero resuelve
un problema muy concreto y muy visible, que es que LBS le cruza los pedidos entre camiones.
El motor deshace ese cruce y vuelve a armar un camión por pedido.

## 1. Identificación

`LBS_IsSamsChain` (`tms_fg14/modulo2.vba:12021`), comparación exacta:

```
LBS_IsSamsChain = (Replace(UCase$(Trim$(CStr(m))), " ", "") = "SAMS")
```

Acepta `Sams` y `SAMS`. **No** acepta `Sam's` ni `Sams Club`: el apóstrofo y la palabra
adicional no se eliminan en la normalización, que solo quita espacios.

Sin familia en `LBS_ChainFamily`, y ausente de todas las listas transversales de altura, peso
y sándwich. En términos del motor, Sams es una cadena de comportamiento por omisión más un
re-armado por pedido.

Sí tiene bandera de presencia (`mLBS_HasSams`, poblada por `LBS_ScanChainPresence` en
`tms_fg14/modulo2.vba:6246-6248`), para que el consolidador de fallos pueda saltarse su fase
cuando no hay filas de la cadena.

## 2. Parámetros

| Parámetro | Valor | Origen |
|---|---|---|
| Cupo de camión | 26 tarimas | `LBS_METRO_TRUCK_CAP` (`modulo2.vba:19`) |
| Piso de llenado | Ninguno propio | No está en `LBS_IsTruckMinFillChain` (`modulo2.vba:4434`) |
| Altura máxima | Sin tope | No está en `LBS_ChainEnforcesUnitHeight` (`modulo2.vba:11895`) |
| Camas compartidas | No | No está en `LBS_ChainAllowsSharedCamas` (`modulo2.vba:6342`) |
| Peso máximo | 29 t | `SK_MAX_PESO_KG` (`modulo2.vba:10`) |

El cupo de 26 está escrito directamente en el comentario del re-armado, no se toma del
catálogo. Ver abajo.

Sin piso de llenado propio significa que a Sams solo le aplica el gate de `AR` contra la hoja
`EFICIENCIA POR CADENA`. Un camión de Sams con 8 tarimas no se descarta por bajo llenado,
porque no hay piso que lo descarte.

## 3. Reglas de negocio

**El problema.** El comentario de `LBS_UnmixSamsPedidoFolios`
(`tms_fg14/modulo2.vba:12024-12026`) lo describe con el patrón exacto:

```
' SAMS: LBS a menudo cruza pedidos entre trucks (23+3 ciclos). Por plant|dest, reempaca
' Programado por pedido en camiones de cap 26 (un pedido por truck salvo volumen >cap).
```

El patrón `23+3` es característico: LBS arma un camión con 23 tarimas de un pedido y 3 de
otro, y luego otro camión con el resto. El resultado son camiones mezclados sin razón
operativa: los pedidos cabían cada uno en su camión.

**La solución: re-armado por pedido.** El motor agrupa por planta y destino, y vuelve a
empacar las filas `Programado` de forma que **cada pedido tenga su propio camión**. La única
excepción es un pedido cuyo volumen exceda el cupo de 26, que necesariamente ocupa más de un
camión.

Es la única cadena donde el motor descarta por completo el armado de LBS y lo rehace. En
todas las demás lo ajusta.

**El pool de folios.** Al rehacer el armado hacen falta números de folio. En lugar de generar
siempre folios nuevos, `LBS_SamsTakePoolOrNextAD` (`tms_fg14/modulo2.vba:12205`) primero
reutiliza los folios que quedaron libres al deshacer el cruce, y solo cuando el pool se agota
genera uno nuevo:

```
Do While poolIdx <= adPool.Count
    cand = CStr(adPool(poolIdx))
    poolIdx = poolIdx + 1
    If Not usedAD.Exists(cand) Then
        usedAD.Add cand, True
        LBS_SamsTakePoolOrNextAD = cand
```

La razón es operativa: si el camión que llevaba 23 tarimas del pedido A ahora lleva las 26,
conviene que conserve su folio original. Cambiarle el número por gusto obliga a reconciliar
folios que nadie pidió cambiar.

Cuando el pool se agota, cae en `LBS_NextLaComerFolioAD`
(`tms_fg14/modulo2.vba:12288`) para generar el siguiente número, verificando que no esté ya
en uso. El nombre de esa función es de La Comer porque ahí se escribió primero; la usan
ambas cadenas.

**Herencia de fecha y equipo.** Al montar las filas en el folio nuevo se heredan la fecha y
el equipo de la plantilla de la llave de consolidación
(`tms_fg14/modulo2.vba:12192-12195`), para que el camión rearmado no quede sin esos datos.

## 4. Procedimientos

| Procedimiento | Línea | Qué hace |
|---|---|---|
| `LBS_IsSamsChain` | `12021` | Reconoce `SAMS` |
| `LBS_UnmixSamsPedidoFolios` | `12027` | Deshace el cruce de pedidos y rearma un camión por pedido |
| `LBS_SamsTakePoolOrNextAD` | `12205` | Toma un folio del pool de liberados, o genera uno nuevo |

Y la bandera de presencia, poblada en `LBS_ScanChainPresence`
(`tms_fg14/modulo2.vba:6246`).

## 5. Cómo validarlo

Fixture:

| Archivo | Contenido |
|---|---|
| `tms_fg14/sams/sample.tsv` | Muestra de `Pedidos Surtidos` con el cruce de pedidos |

Es el único fixture de la cadena. Sirve para verificar el patrón `23+3` antes del re-armado y
un pedido por camión después.

Scripts:

| Script | Qué comprueba |
|---|---|
| `scripts/validate_sample_tsv.py` | Estructura general de la muestra |
| `scripts/validate_consolidar_split_mismo_pedido.py` | Que no queden divisiones dentro del mismo pedido |
| `scripts/validate_multi_order_single.py` | Sencillos con varios pedidos |

## 6. Problemas conocidos y síntomas

| Síntoma | Causa probable | Dónde mirar |
|---|---|---|
| Camiones con dos pedidos mezclados (patrón `23+3`) | El re-armado no corrió | `LBS_UnmixSamsPedidoFolios`; verificar que `Pedidos Surtidos!M` diga exactamente `Sams` |
| `Sam's Club` o `Sams Club` sin reglas de la cadena | El clasificador solo acepta `SAMS` | La normalización quita espacios pero no apóstrofos ni palabras extra |
| Folios que cambiaron de número sin razón | El pool se agotó y se generaron folios nuevos | `LBS_SamsTakePoolOrNextAD`. Es esperado si hay más camiones después del re-armado que antes |
| Camiones muy vacíos que no se descartan | Sams no tiene piso de llenado propio | Solo aplica el gate de `AR`. Ajustar el umbral en `EFICIENCIA POR CADENA` si el cliente lo quiere más estricto |
| `AO` y `AP` vacíos | Sams no está en las listas de altura | Es lo esperado |
| Camión con más de 26 tarimas | Un solo pedido excede el cupo | Es la excepción prevista: `un pedido por truck salvo volumen >cap` |
| Camión rearmado sin fecha o sin equipo | Falló la herencia desde la plantilla de la llave | `tms_fg14/modulo2.vba:12192-12195` |
