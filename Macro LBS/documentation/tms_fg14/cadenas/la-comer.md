[Volver al índice de cadenas](README.md)

# La Comer

La Comer es la única cadena que agrega un componente propio a la llave de consolidación: la
marca `R` o `C` de la columna `AE`. Un camión de La Comer es de refrigerado o de seco, nunca
mezclado, y eso atraviesa todo su tratamiento. Es también la única cadena que se sacó a
propósito del mecanismo genérico de sándwich, porque tiene el suyo.

## 1. Identificación

`LBS_IsLaComerChain` (`tms_fg14/modulo2.vba:11889`), comparación exacta:

```
LBS_IsLaComerChain = (Replace(UCase$(Trim$(CStr(m))), " ", "") = "LACOMER")
```

Acepta `La Comer`, `LA COMER` y `LaComer`. No tiene familia propia en `LBS_ChainFamily`, así
que en las listas por familia (peso, apertura de camiones) no participa.

Tampoco aparece en la tabla embebida de grupos de consolidación: cada destinatario es su
propio grupo salvo que se mapee a mano en la hoja `Consolida`.

## 2. Parámetros

| Parámetro | Valor | Origen |
|---|---|---|
| Cupo sencillo | 26 tarimas | `LBS_METRO_TRUCK_CAP` (`modulo2.vba:19`) |
| Cupo Full | Catálogo Mode Mix | `LBS_CatalogFullShipmentCap` (`modulo2.vba:5703`) |
| Cupo caja `a`/`b` | Cupo de embarque entre 2 | `LBS_FullBoxCapForRow` (`modulo2.vba:5513`) |
| Piso de llenado | 80 % del cupo = 21 de 26 | `LBS_LACOMER_MIN_FILL` (`modulo2.vba:78`) |
| Altura máxima de unidad | 1.60 m | `LBS_LACOMER_MAX_HEIGHT_M` (`modulo2.vba:83`) |
| Peso máximo | 29 t | `SK_MAX_PESO_KG` (`modulo2.vba:10`) |

Comentarios originales:

```
' LBS - LA COMER: piso de llenado post-consolidacion (cap 26 -> piso 21 = 80%).
Private Const LBS_LACOMER_MIN_FILL As Double = 0.8
```

```
' LBS - LA COMER: altura maxima de tarima/unidad (suma AO -> AP), trial = Walmart 1.6 m.
Private Const LBS_LACOMER_MAX_HEIGHT_M As Double = 1.6
```

La palabra `trial` es literal y vale la pena tenerla presente: el límite de 1.60 m de La
Comer **no viene de una especificación del cliente**, se puso a prueba copiando el de
Walmart. Si el cliente cuestiona por qué se le limitan las unidades a 1.60 m, esa es la
respuesta honesta. La Comer sí tiene constante propia (a diferencia de Soriana o Chedraui,
que reutilizan la de Walmart), precisamente para poder ajustarla sin tocar las demás.

## 3. Reglas de negocio

**Refrigerado y seco no se mezclan.** La columna `AE` trae `R` o `C`, y
`LBS_LaComerConfRC` (`tms_fg14/modulo2.vba:12226`) solo acepta esos dos valores exactos:
cualquier otra cosa devuelve cadena vacía. Esa marca se agrega a la llave de consolidación
(`LBS_ConsolidaKey`, `tms_fg14/modulo2.vba:5469-5472`), así que un camión de La Comer queda
definido por:

```
origen (col L) | grupo | R o C | vigencia (col R)
```

Es el único caso en toda la macro donde la llave tiene cuatro componentes.

**El problema de la marca en blanco.** Cuando `AE` viene vacía, la fila se cae de la llave y
el motor la cuenta mal. El comentario de la corrección
(`tms_fg14/modulo2.vba:12235-12238`) trae el caso concreto:

```
' Cont1 bare fulls often land with blank AE while sandwich sisters on the same folio
' carry R/C. Blank AE drops them out of ConsolidaKey (...|R), so Z counts them but
' spare/pack thinks the R truck is short (P-1096 Cont1 T=13 -> false packable remains).
```

En prosa: un Full pelado llega sin marca mientras sus hermanas de sándwich en el mismo folio
sí la tienen. Sin marca no entra a la llave `...|R`, así que el total de la columna `Z` lo
cuenta pero el cálculo de espacio libre no, y el motor cree que al camión `R` le sobran 13
tarimas de lugar cuando en realidad están ocupadas. `LBS_FillLaComerBlankConfRC`
(`tms_fg14/modulo2.vba:12239`) rellena la marca en blanco tomándola de las otras filas del
mismo folio.

**División por `R`/`C`.** Si un folio quedó con las dos marcas mezcladas,
`LBS_SplitLaComerRCFolios` (`tms_fg14/modulo2.vba:12294`) lo divide en dos folios, usando
`LBS_NextLaComerFolioAD` (`tms_fg14/modulo2.vba:12288`) para generar el nuevo número.

**Equipo forzado a caja seca.** `LBS_ForceLaComerYCajaSeca` (`tms_fg14/modulo2.vba:12383`)
sobreescribe la columna `Y` del equipo. Es un ajuste de La Comer sobre lo que LBS haya
asignado.

**Sándwich propio.** La Comer está excluida a propósito del mecanismo genérico. El comentario
de `LBS_IsSandwichWChain` (`tms_fg14/modulo2.vba:6352`):

```
' LA COMER stays outside — FixLaComerSandwichAnchors owns its ANCHOR/W recovery.
```

`LBS_FixLaComerSandwichAnchors` (`tms_fg14/modulo2.vba:13301`) recupera las anclas y la
columna `W` con su propia lógica. Meterla en el mecanismo genérico rompería el suyo.

**Anclas huérfanas.** Cuando un sándwich pierde su ancla, las capas quedan sin base.
`LBS_DiscardLaComerOrphanSandwichSisters` (`tms_fg14/modulo2.vba:13177`) las descarta, en
lugar de dejarlas flotando en un camión donde físicamente no se pueden colocar.

**Tarimas craft.** `LBS_LaComerTruckHasCraftAI` (`tms_fg14/modulo2.vba:13647`) detecta si un
camión ya tiene una unidad con tarima craft. La presentación craft se reconoce por el texto
`CRAFT` en la columna `AC` (`LBS_IsCraftPalletStart`, `tms_fg14/modulo2.vba:6361`), por
ejemplo en `CORONITA ... CRAFT AF`. Es un tipo de tarima distinto que no se mezcla
libremente.

**Apertura de camiones desde `No planeado`.** Es la cadena con el mecanismo más desarrollado
en este punto. `LBS_OpenLaComerTrucksFromNoPlaneado` (`tms_fg14/modulo2.vba:13454`) abre
folios nuevos con lo descartado, `LBS_LaComerAssignNpSandwichW`
(`tms_fg14/modulo2.vba:13425`) les asigna el conteo de sándwich, y
`LBS_LaComerPullNpSistersIntoTruck` (`tms_fg14/modulo2.vba:13662`) jala las filas hermanas
del mismo sándwich al camión nuevo, para no partir una unidad física entre dos camiones.

**Recorte y finalización de cupo.** `LBS_TrimLaComerOverCap`
(`tms_fg14/modulo2.vba:13061`) y `LBS_FinalizeLaComerCap` (`tms_fg14/modulo2.vba:13407`)
cierran el cupo. Hay un comentario que documenta un caso concreto
(`tms_fg14/modulo2.vba:22330`):

```
' peel LA COMER over-cap before writing Z so P-1109 does not keep REVISION MANUAL >26.
```

El orden importa: si se escribe el total de la columna `Z` antes de recortar, el folio se
queda con una marca de `REVISION MANUAL` que ya no corresponde.

**Camas compartidas y altura, sin peso desde el catálogo.** La Comer está en
`LBS_ChainAllowsSharedCamas` (`tms_fg14/modulo2.vba:6342`) y en
`LBS_ChainEnforcesUnitHeight` (`tms_fg14/modulo2.vba:11895`), pero no en
`LBS_ChainAppliesTihiAT` (`tms_fg14/modulo2.vba:6336`).

## 4. Procedimientos

| Procedimiento | Línea | Qué hace |
|---|---|---|
| `LBS_IsLaComerChain` | `11889` | Reconoce `LACOMER` |
| `LBS_LaComerConfRC` | `12226` | Normaliza la marca `AE` a `R`, `C` o vacío |
| `LBS_FillLaComerBlankConfRC` | `12239` | Rellena la marca en blanco desde el resto del folio |
| `LBS_NextLaComerFolioAD` | `12288` | Genera el siguiente número de folio |
| `LBS_SplitLaComerRCFolios` | `12294` | Divide folios con `R` y `C` mezclados |
| `LBS_ForceLaComerYCajaSeca` | `12383` | Fuerza el equipo a caja seca |
| `LBS_SplitLaComerOverCapFolios` | `12542` | Divide folios sobre el cupo |
| `LBS_LaComerGroupEfficiencyGate` | `12584` | Gate de eficiencia por grupo |
| `LBS_PackLaComerOneCkey` | `12767` | Empaca una llave de consolidación |
| `LBS_PackLaComerToCapacity` | `12946` | Recorre todas las llaves empacando |
| `LBS_LayerLaComerRestosCharolas` | `12993` | Acomoda los restos en charolas |
| `LBS_TrimLaComerOverCap` | `13061` | Recorta lo que excede el cupo |
| `LBS_DiscardLaComerOrphanSandwichSisters` | `13177` | Descarta capas sin ancla |
| `LBS_EnforceLaComerPostConsolidation` | `13232` | Aplica el piso del 80 % |
| `LBS_FixLaComerSandwichAnchors` | `13301` | Recupera anclas y columna `W`, mecanismo propio |
| `LBS_FinalizeLaComerCap` | `13407` | Cierra el cupo antes de escribir `Z` |
| `LBS_LaComerNpTarimas` | `13419` | Tarimas de una fila `No planeado` |
| `LBS_LaComerAssignNpSandwichW` | `13425` | Conteo de sándwich para las filas `No planeado` |
| `LBS_OpenLaComerTrucksFromNoPlaneado` | `13454` | Abre camiones nuevos con lo descartado |
| `LBS_LaComerTruckHasCraftAI` | `13647` | Detecta tarima craft en el camión |
| `LBS_LaComerPullNpSistersIntoTruck` | `13662` | Jala las hermanas del sándwich al camión nuevo |

Las fases de `SummaryOptimizar` son `filtrar:split_lacomer` (`tms_fg14/modulo2.vba:4858`),
`filtrar:lacomer_gate` (`4944`) y `filtrar:pack_lacomer` (`4947`).

## 5. Cómo validarlo

Fixtures en `tms_fg14/lacomer/`:

| Archivo | Contenido |
|---|---|
| `sample.tsv` | Muestra de `Pedidos Surtidos` |
| `plan.tsv` | Plan de prueba |
| `w bug sample.tsv` | El caso del conteo `W` mal calculado |

El nombre `w bug sample.tsv` corresponde al problema de la marca `AE` en blanco descrito
arriba: la columna `W` quedaba mal porque las filas sin marca se caían de la llave.

Scripts:

| Script | Qué comprueba |
|---|---|
| `scripts/validate_lacomer_rc.py` | La separación por marca `R` y `C` |
| `scripts/validate_lacomer_restos.py` | El acomodo de restos y las anclas de sándwich |
| `scripts/validate_lacomer_weight_tihi.py` | El peso contra el catálogo TI HI |
| `scripts/validate_tihi_pkg_lacomer.py` | El empaque contra el catálogo |

## 6. Problemas conocidos y síntomas

| Síntoma | Causa probable | Dónde mirar |
|---|---|---|
| Camión con producto refrigerado y seco | La marca `AE` venía en blanco y la fila se cayó de la llave | `LBS_FillLaComerBlankConfRC`. Revisar `AE` de las filas del folio |
| Espacio libre mal calculado, camión "corto" que no lo está | El mismo problema de la marca en blanco (caso `P-1096`) | Igual |
| `REVISION MANUAL: tarimas >26` que no corresponde | El total `Z` se escribió antes del recorte | `LBS_FinalizeLaComerCap`. Es el caso `P-1109` del comentario |
| Capas de sándwich sin tarima base | El ancla se perdió en una fase anterior | `LBS_DiscardLaComerOrphanSandwichSisters` debería haberlas descartado |
| Un sándwich partido entre dos camiones | Falló el arrastre de hermanas al abrir camión nuevo | `LBS_LaComerPullNpSistersIntoTruck` |
| `REVISION MANUAL: altura >1.6` | El tope de La Comer es de prueba, copiado de Walmart | Si el cliente lo cuestiona, `LBS_LACOMER_MAX_HEIGHT_M` es ajustable sin afectar otras cadenas |
| `Descartado por baja eficiencia` en camiones de 18 a 20 tarimas | Bajo el piso de 21 | `LBS_LACOMER_MIN_FILL` |
| Folios nuevos con producto que estaba en `No planeado` | Es el comportamiento esperado | `LBS_OpenLaComerTrucksFromNoPlaneado` |
| Equipo distinto del que asignó LBS | La macro lo fuerza a caja seca | `LBS_ForceLaComerYCajaSeca` |
