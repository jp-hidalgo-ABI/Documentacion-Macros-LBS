[Volver a la macro de salida](README.md) · [Índice general](../README.md)

# Eficiencia y descartes

Un camión puede estar bien armado y aun así no salir. Este documento explica los dos
mecanismos que deciden qué se despacha:

1. El **gate de eficiencia** — compara el fill rate que reportó LBS (`AR`) contra el umbral
   de la cadena.
    10|2. El **piso de llenado** — compara las tarimas del camión contra un porcentaje de su cupo.

Son independientes y se aplican en momentos distintos. Una fila puede pasar el gate y
después caer por el piso, o al revés.

## `AR` — el fill rate de LBS

`AR` no lo calcula la macro. Lo copia del `Load Efficiency` de la hoja `Shipments`, cruzando
por el `Shipment Id` que extrae del folio `AD` (`tms_fg14/modulo2.vba:4912-4931`):

```vba
    20|shipmentCode = SK_ShipmentIdFromAD(SK_CellStr(wsPS.Cells(i, "AD").Value))
If Len(shipmentCode) > 0 And shipEff.Exists(shipmentCode) Then
    valorEff = shipEff(shipmentCode)
    ...
    wsPS.Cells(i, "AR").Value = CDbl(valorEff)
```

El mapa lo arma `LBS_BuildShipEffMap` (`tms_fg14/modulo2.vba:4814-4826`), que registra el
**primer** valor visto por `Shipment Id`.

Tres consecuencias prácticas:
    30|
- **`AR` es un valor por embarque, no por fila.** Todas las filas del mismo folio comparten el
  mismo `AR`.
- **Un folio sin guion en `AD` no obtiene `AR`.** La condición
  `InStr(AD, "-") > 0` (`tms_fg14/modulo2.vba:4914`) es lo que distingue un folio real de un
  valor residual.
- **`AR` se borra al final.** El último recorrido de `FiltrarPorEficiencia` limpia `AR` en
  **todas** las filas, no solo las descartadas (`tms_fg14/modulo2.vba:5040`). Después de una
  corrida completa la columna está vacía; para verla hay que detenerse en `SummaryOptimizar`.

## La hoja `EFICIENCIA POR CADENA`
    40|
Dos columnas:

| Col | Contenido |
|---|---|
| `A` | El nombre de la cadena |
| `B` | El umbral de fill rate, en porcentaje |

La carga la hace `LBS_LoadEficRefCache` (`tms_fg14/modulo2.vba:2581-2618`), una vez por
sesión. Detalles que importan:

    50|- **La llave se normaliza** quitando espacios y pasando a mayúsculas:
  `Replace(UCase$(Trim$(A)), " ", "")`. `La Comer` y `LACOMER` son la misma llave.
- **Solo se aceptan valores numéricos** en `B`. Un umbral escrito como texto se ignora en
  silencio.
- **El umbral por omisión es 90** cuando la cadena no aparece
  (`LBS_EficienciaRefForCadena`, `tms_fg14/modulo2.vba:2624`). Una cadena nueva arranca
  exigiendo 90% de fill rate.
- **Solo se leen las primeras 500 filas** (`tms_fg14/modulo2.vba:2601`), con este comentario:

```
    60|' Used-range pollution in col A used to scan ~1M rows per lookup and kill Excel.
```

Una cadena listada más abajo de la fila 500 se ignora sin aviso. Es poco probable, pero es la
clase de detalle que explica un descarte inexplicable.

## El descarte por fila

Es el mecanismo más simple. Al final de `FiltrarPorEficiencia`
(`tms_fg14/modulo2.vba:4997-5018`):

    70|```vba
If eficienciaRef >= 0 And valorEffTem < eficienciaRef Then
    wsPS.Cells(i, "H").Value = "No planeado"
    wsPS.Cells(i, "AG").Value = "Descartado por baja eficiencia"
End If
```

**Cuatro cadenas quedan excluidas** de este paso porque tienen su propio gate
(`tms_fg14/modulo2.vba:5000-5003`): Walmart, Chedraui, Alsuper y La Comer. El comentario lo
aclara y agrega un dato importante sobre Soriana y City Club
    80|(`tms_fg14/modulo2.vba:4995-4997`):

```
'--- Segundo paso: comparar con EFICIENCIA POR CADENA (no Walmart/CHEDRAUI/ALSUPER/LA COMER;
' ya procesados por gate/folio). ClubCity (Soriana + City Club) uses this per-row AR
' discard only — not truck min-fill in LBS_EnforceWalmartMinFill. ---
```

## Los gates por grupo

El descarte por fila es demasiado crudo para las cadenas grandes: mata un pedido completo
    90|porque un embarque salió con fill rate bajo, aunque ese pedido combinado con otro llene un
camión. Los gates resuelven eso evaluando conjuntos.

### El gate de Chedraui — por pedido

`LBS_ChedrauiPedidoEfficiencyGate` (`tms_fg14/modulo2.vba:2717`). Su comentario resume las
tres reglas (`tms_fg14/modulo2.vba:2714-2716`):

```
' CHEDRAUI: eficiencia por destino (origen|dest) antes del descarte folio-a-folio.
' Reglas: promedio ponderado por tarimas; rescate todo-fallido si cabe en un camion;
   100|' merge de folios fallidos sobre folios que pasan cuando hay cupo.
```

1. **Promedio ponderado por tarimas.** No se promedian los `AR` a secas: cada folio pesa por
   su número de tarimas.
2. **Rescate todo-fallido.** Si *todos* los folios del destino fallan pero la carga combinada
   cabe en un camión y alcanza `LBS_CHEDRAUI_RESCUE_MIN_FILL = 0.8` del cupo, se rescatan.
3. **Merge parcial.** Si algunos pasan y otros no, los fallidos se suben a los que pasan
   mientras haya cupo.

La llave de agrupación es `LBS_ChedrauiDestPedidoKey` (`tms_fg14/modulo2.vba:2707-2712`):
   110|`grupo | vigencia`. El comentario cita el acuerdo que la definió:

```
' CHEDRAUI: dest (sin origen ni pedido) para anclas de charolas cross-planta.
' Minuta #20: incluye vigencia para no mezclar pedidos con distinta col R en un camion.
```

### El gate de Walmart y Alsuper — por grupo de consolidación

`LBS_WalmartGroupEfficiencyGate` (`tms_fg14/modulo2.vba:2903`) se llama **dos veces**, con el
mismo código y distinto parámetro (`tms_fg14/modulo2.vba:4966`, `4970`):

   120|```vba
' WALMART BA/SC: rescate por grupo de consolidacion (plant|metro) antes del descarte folio.
Call LBS_WalmartGroupEfficiencyGate(wsPS, lastRowPS, wsEfic, lastRowEfic)
' ALSUPER: mismo gate de grupo (plant|ALSUPERdest) con rescate multi-camion.
Call LBS_WalmartGroupEfficiencyGate(wsPS, lastRowPS, wsEfic, lastRowEfic, "ALSUPER")
```

El umbral de rescate es `LBS_WALMART_RESCUE_MIN_FILL = 0.5`, y su comentario advierte de un
punto que se confunde con facilidad (`tms_fg14/modulo2.vba:57-60`):

```
   130|' LBS - WALMART: rescate de grupo todo-fallido si la carga combinada >= este % del cap (28 tarimas).
' Solo para el gate de grupo (LBS_WalmartGroupEfficiencyGate); el piso post-consolidacion es
' LBS_WALMART_MIN_FILL.
```

Son dos umbrales distintos para Walmart: **50% para rescatar en el gate, 40% como piso del
camión.** No son el mismo número y no se aplican en el mismo momento.

### El gate de La Comer — por grupo

`LBS_LaComerGroupEfficiencyGate`, llamado en `tms_fg14/modulo2.vba:4946`. Agrupa por la llave
   140|de consolidación, que para La Comer incluye la marca R/C.

## El piso de llenado

El piso es el segundo filtro, y es el que descarta camiones bien armados pero poco llenos.
Lo aplica `LBS_EnforceWalmartMinFill` (`tms_fg14/modulo2.vba:4441-4588`), cuyo nombre engaña:
cubre **ocho cadenas**, no solo Walmart.

### Qué cadenas lo aplican

`LBS_IsTruckMinFillChain` (`tms_fg14/modulo2.vba:4434-4439`): Walmart, Alsuper, Go Mart,
   150|Europea, Soriana, City Club, OXXO y COMEXTRA.

Chedraui tiene su propia versión, `LBS_EnforceChedrauiPostConsolidation`
(`tms_fg14/modulo2.vba:4593`), y La Comer también
(`LBS_EnforceLaComerPostConsolidation`, llamada en `tms_fg14/modulo2.vba:4953`).

### Cómo se calcula el piso

`LBS_WalmartMinTarimas` (`tms_fg14/modulo2.vba:4404-4426`) redondea **hacia arriba** con el
idioma `-Int(-x)`:

   160|| Cadena | Fracción | Constante | Ejemplo |
|---|---|---|---|
| Soriana, City Club | 70% | `LBS_CLUBCITY_MIN_FILL` | cupo 26 → piso 19 |
| OXXO | 90% | `LBS_OXXO_MIN_FILL` | shipCap 36 → piso 33 |
| COMEXTRA | 90% | `LBS_COMEXTRA_MIN_FILL` | ship 40 → 36; caja 20 → 18; sencillo 26 → 24 |
| Todas las demás | 40% | `LBS_WALMART_MIN_FILL` | cupo 28 → piso 12 |

El comentario de `LBS_WALMART_MIN_FILL` distingue las dos cosas que la gente mezcla
(`tms_fg14/modulo2.vba:61-65`):

```
   170|' LBS - WALMART: piso de llenado de camion (tarimas). El % de "EFICIENCIA POR CADENA" es
' fill rate de PEDIDO (col AR, gate de eficiencia) para Walmart/Alsuper — truck floor is
' this constant (cap 28 -> piso 12). ClubCity (Soriana/City Club) truck floor is
' LBS_CLUBCITY_MIN_FILL (70% of metro cap 26 -> piso 19).
```

**El porcentaje de la hoja `EFICIENCIA POR CADENA` no es el piso de llenado.** Es el umbral
del fill rate por pedido. Subir el número de la hoja no hace que se descarten camiones poco
llenos, ni bajarlo los rescata.

   180|Aun así, para Soriana y City Club el código consulta la hoja **solo para calentar la caché**
(`tms_fg14/modulo2.vba:4411-4414`):

```vba
' Warm EFICIENCIA cache (pedido AR gate). Truck floor is always 70%.
effRef = LBS_EficienciaRefForCadena(cadena, Nothing, 0)
LBS_WalmartMinTarimas = -Int(-(cap * LBS_CLUBCITY_MIN_FILL))
```

El valor leído se descarta. El piso es siempre 70%.
   190|
Chedraui tiene la misma independencia, y el comentario explica por qué se hizo así
(`tms_fg14/modulo2.vba:4590-4592`):

```
' CHEDRAUI: descarta camiones Programado bajo effRef (AR) o bajo 80% de llenado.
' Truck floor uses LBS_CHEDRAUI_MIN_FILL (not EFICIENCIA sheet) so a missing/zero
' sheet cannot leave under-floor leftovers Programado (e.g. P-1031 at 3 tarimas).
```

Una hoja vacía o con cero habría dejado pasar un camión de 3 tarimas. El folio `P-1031` es la
   200|evidencia.

### Los tres tramos de resultado

Con `total` = tarimas efectivas del folio y `cap` = su cupo:

| Condición | Resultado |
|---|---|
| `total >= cap` | Nada. El camión está lleno |
| `minTar <= total < cap` | Bandera `REVISION MANUAL: tarimas <N (M tarimas)` en `AV` |
| `total < minTar` | **Todo el folio pasa a `No planeado`** |

   210|(`tms_fg14/modulo2.vba:4535`, `4555-4585`.)

El tramo intermedio es informativo: el camión sale, pero con aviso.

### Los Fulls partidos se evalúan juntos

Un detalle que evita romper los splits (`tms_fg14/modulo2.vba:4469-4471`):

```
' Fulls partidos a/b (PartirFulles): evaluar por shipment base (cap 40 = a+b),
' no por caja; si no, la caja b (~10-20 tarimas) se descarta sola y rompe el split.
   220|```

Los folios `...a` y `...b` se colapsan a su folio base antes de medir. De otro modo la caja
`b`, que suele llevar la mitad de la carga, caería sola y dejaría media carga sin camión.

Lo mismo por el lado del cupo (`tms_fg14/modulo2.vba:4518-4520`):

```
' ClubCity/OXXO/COMEXTRA Full min-fill is shipment-level (a+b collapsed to base AD).
' Never use box 20/18 here — orphan half-box leftovers must face shipCap
' (ClubCity 40 piso 70% / OXXO 36 piso 90% / COMEXTRA 40 piso 90%) or merge first.
   230|```

### Los saldos pendientes cuentan para el piso

Es la regla menos obvia y la más importante de entender. Las filas `No planeado` cuyo `AG`
contiene `No reportado por LBS` y que tienen `T > 0` **se suman al total del folio** al medir
el piso (`tms_fg14/modulo2.vba:4480-4492`, `4537-4548`).

El comentario da la razón (`tms_fg14/modulo2.vba:4432-4433`):

```
' Los saldos "No reportado por LBS" (CompararCartonajes) del mismo pedido cuentan para el piso:
   240|' se montan hasta Fallos, y descartar el folio antes los dejaria sin camion (pedido entero fuera).
```

La secuencia del problema: `CompararCartonajes` detecta cartonaje del Plan que LBS no reportó
y lo inserta como `No planeado`. Ese saldo se monta después, en `SummaryFallo`. Si el piso se
midiera sin contarlo, el folio se descartaría antes de que el saldo llegara, y el pedido
completo quedaría fuera.

Cada saldo se cuenta **una sola vez**: al usarlo se remueve del diccionario
(`pendPed.Remove ped`, `tms_fg14/modulo2.vba:4544`), y se limita a `cap - total` para no
inflar el total por encima del cupo.

   250|### Qué se limpia al descartar

Cuando un folio cae por el piso, no basta con cambiar `H`. Se borran diez columnas
(`tms_fg14/modulo2.vba:4563-4581`): `Y`, `X`, `Z`, `AD`, `AK`, `AT`, `AU`, `AI`, `AN`, `AP`,
`W`, y se pone `I = 0`. Además se quitan las banderas de tarimas de `AV` con
`SK_StripReviewFlagMatching`, porque ya no aplican.

El comentario justifica hacerlo de inmediato: `' Clear truck/unit cols immediately (Fallos may
rewrite Z next).'`

## Los motivos de `AG`
   260|
### Motivos de descarte

| Texto en `AG` | Quién lo pone | Qué significa |
|---|---|---|
| `Descartado por baja eficiencia` | `FiltrarPorEficiencia` (`tms_fg14/modulo2.vba:5014`) y el piso de Soriana, City Club, OXXO y COMEXTRA (`4559`) | El `AR` quedó bajo el umbral de la cadena, o el camión no llegó a su piso |
| `Descartado: bajo llenado post-consolidacion` | El piso, para las demás cadenas (`tms_fg14/modulo2.vba:4561`) | El camión no llegó al piso de llenado |
| `Descartado por exceso de cupo` | Las funciones de recorte de cupo | La tarima no cupo en el camión |
| `Descartado: peso >29 ton` | El descargue por sobrepeso | El camión pasaba del techo de peso |
| `Descartado: altura >1.6m` | El control de altura | La unidad pasaba del tope de altura |
| `No reportado por LBS` | `CompararCartonajes` (`tms_fg14/modulo3.vba:1058`) | Cartonaje del Plan que LBS no colocó |
   270|| `No reportado por LBS (cartonaje Plan)` | `CompararCartonajes` (`tms_fg14/modulo3.vba:1053`) | Variante con la fuente explícita |

Que las dos primeras existan es intencional: el texto distingue de qué mecanismo vino el
descarte, aunque el efecto sea el mismo.

### Marcadores internos

| Valor | Significado |
|---|---|
| `LBS_SANDWICH_ANCHOR` | Ancla de unidad sándwich |
| `LBS_SANDWICH` | Hermana de unidad sándwich |
   280|| `LBS_CHAROLA_MOUNT` | Resto de charola montado sobre otra tarima; luego se convierte en `LBS_SANDWICH` |

Estos aparecen en filas `Programado` y **no** son errores.

### El caso `Shippable qty is 0`

Es un motivo que viene de LBS, no de la macro. Por omisión esas filas **no** se pueden
remontar (`tms_fg14/modulo2.vba:84-85`):

```
' TEST MODE Plan!V1: allow remount of "Shippable qty is 0" AG rows (default off).
   290|Private Const LBS_IGNORE_SHIPPABLE_QTY_ZERO_DEFAULT As Boolean = False
```

Con el modo prueba activo en `Plan!V1` sí se remontan, y `SummaryFallo` lo reporta
explícitamente (`tms_fg14/modulo3.vba:855`):

```
TEST MODE (Plan!V1): se ignoro 'Shippable qty is 0' al remountar.
```

Ver la sección de modo prueba en [10-diagnostico-y-errores.md](10-diagnostico-y-errores.md).
   300|
### La limpieza de motivos obsoletos

Dos funciones evitan que queden motivos viejos en filas que cambiaron de estado:

| Función | Cita | Qué hace |
|---|---|---|
| `LBS_ClearStaleDescartadoAGOnProgramado` | `tms_fg14/modulo2.vba:4684` | Borra motivos de descarte en filas que volvieron a `Programado` |
| `LBS_ClearMountAGUnlessNoReportado` | `tms_fg14/modulo2.vba:4675-4682` | Lo mismo, pero preserva `No reportado por LBS` |

La segunda existe porque ese marcador tiene que sobrevivir: es lo que le dice al piso de
   310|llenado que hay un saldo pendiente.

## El orden completo de `FiltrarPorEficiencia`

`SummaryOptimizar` delega en `FiltrarPorEficiencia` (`tms_fg14/modulo2.vba:4828`), que ejecuta
alrededor de cuarenta pasos. El nombre de cada fase se escribe en `Plan!X1`, así que si la
macro se cae, ahí queda el punto exacto.

| Fase (`SK_Phase`) | Qué hace |
|---|---|
| `filtrar:clubcity_cross_vig` | Segunda pasada de Soriana y City Club por carril, ignorando vigencia |
   320|| (sin fase) | `LBS_UnmixSamsPedidoFolios`: reagrupa los pedidos cruzados de Sams |
| (sin fase) | `LBS_UnifyProgramadoFolioY`: un solo tipo de unidad por folio |
| (sin fase) | `LBS_DedupeFolioProductRows`: elimina filas duplicadas de folio + producto |
| (sin fase) | Asignación y dedupe de unidades `AI` de Walmart |
| `filtrar:recompute_totals` | Recalcula `Z` y `AU` por camión |
| `filtrar:lookup_ar` | Copia `AR` desde `Shipments` |
| `filtrar:chedraui_gate` | Gate de Chedraui |
| `filtrar:pack_chedraui` | Empaqueta Chedraui hasta el cupo |
| `filtrar:enforce_chedraui` | Integridad de folio de Chedraui |
| `filtrar:lacomer_gate` | Gate de La Comer |
   330|| `filtrar:pack_lacomer` | Empaqueta La Comer, apila restos de charolas, recorta y aplica el piso |
| `filtrar:walmart_gate` | Gate de Walmart |
| `filtrar:alsuper_gate` | El mismo gate para Alsuper |
| `filtrar:pack_walmart` | Empaqueta Walmart hasta el cupo |
| `filtrar:min_fill` | **El piso de llenado** de las ocho cadenas, más el de Chedraui |
| `filtrar:totales_finales` | Recalcula `Z` |
| `filtrar:descartar_cadena` | El descarte por fila contra `EFICIENCIA POR CADENA` |
| `filtrar:mover_no_planeado` | Ordena para mandar los `No planeado` al final |
| `filtrar:consolidar_no_planeado` | Llama a `ConsolidarNoPlaneados` |
| `filtrar:pack_cap35` | Remonta los sobrantes de mayoristas |
   340|| `filtrar:z_final` | Reescribe los totales `Z` por folio |

### Dos observaciones sobre este orden

**El piso se aplica antes del descarte por fila.** Un camión puede caer por piso de llenado y
después sus filas volver a evaluarse contra `EFICIENCIA POR CADENA`, pero ya son
`No planeado` y el segundo paso las salta (`tms_fg14/modulo2.vba:4999`).

**La Comer se procesa dos veces.** Después de abrir camiones nuevos desde `No planeado`, se
vuelve a aplicar el piso, y el comentario explica por qué
(`tms_fg14/modulo2.vba:4959`):

   350|```
' Open can leave trucks under 80% once Fix/PrepareTW collapses sister W; enforce again.
```

Abrir un camión y luego colapsar las `W` de las hermanas de sándwich puede dejarlo por debajo
del 80%, así que hay que volver a medir.

## Los restos descartados no se remontan aquí

Es una decisión explícita, con un comentario que registra la paridad que se buscaba
(`tms_fg14/modulo2.vba:5020-5021`):
   360|
```
' LBS - Paridad HEB: montar restos descartados solo en Fallos (LBS_ConsolidarRestos),
' no aqui, para no inflar camiones que ya vienen llenos desde LBS (p.ej. 1310=27, 1382=16).
```

El remonte de restos vive en `SummaryFallo`, no en `SummaryOptimizar`. Hacerlo aquí inflaba
camiones que LBS ya había armado llenos, y los folios 1310 y 1382 son los casos que lo
mostraron. Ver [07-fallos-y-remonte.md](07-fallos-y-remonte.md).

## Cómo diagnosticar un descarte
   370|
1. **Buscar el folio en `AD`** y ver si tiene filas `Programado`. Si todas son
   `No planeado`, el camión cayó completo.
2. **Leer `AG`.** El texto dice qué mecanismo lo descartó.
3. Si dice `Descartado por baja eficiencia`, hay dos posibilidades: falló el gate de `AR` o
   falló el piso de llenado (Soriana, City Club, OXXO, COMEXTRA comparten el texto). Contar
   las tarimas del folio y compararlas contra el piso de la cadena resuelve la duda.
4. Si dice `Descartado: bajo llenado post-consolidacion`, fue el piso, sin ambigüedad.
5. **Revisar `AV`.** Un `REVISION MANUAL: tarimas <N` indica que el camión quedó en el tramo
   intermedio, no descartado.
   380|6. **Si el descarte parece injusto**, verificar si hay saldos `No reportado por LBS` del mismo
   pedido: deberían haber contado para el piso.

Para el caso de dos filas que se esperaba que consolidaran y no lo hicieron, existe
`LBS_DiagnosticarConsolidacion` (`tms_fg14/modulo2.vba:22403`). Ver
[10-diagnostico-y-errores.md](10-diagnostico-y-errores.md).
