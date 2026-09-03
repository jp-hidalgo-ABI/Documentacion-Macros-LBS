[Volver al índice de cadenas](README.md)

# Chedraui

Lo que distingue a Chedraui es que el gate de eficiencia opera **por destino**, no por folio
ni por grupo completo, y que el motor le aplica dos divisiones correctivas: una por orígenes
mezclados en un mismo folio y otra por folios sobre el cupo. Su piso de llenado (80 %) es
además el más alto de las cadenas de autoservicio.

## 1. Identificación

`LBS_IsChedrauiChain` (`tms_fg14/modulo2.vba:6175`), comparación exacta:

```
LBS_IsChedrauiChain = (Replace(UCase$(Trim$(CStr(m))), " ", "") = "CHEDRAUI")
```

`LBS_ChainFamily` (`tms_fg14/modulo2.vba:5205-5206`) le da familia propia, `CHEDRAUI`, con
el mismo literal.

Chedraui **no** aparece en la tabla embebida de grupos de consolidación. El comentario de
`LBS_ConsolidaGroup` lo dice explícitamente (`tms_fg14/modulo2.vba:5441`):

```
' CHEDRAUI / sin mapa: mismo criterio (grupo = dest). Sin destinatario: no consolida.
```

Cada destinatario de Chedraui es su propio grupo, salvo que el planeador lo mapee a mano en
la hoja `Consolida`. Es una decisión de negocio: Chedraui no hace multi-stop por omisión.

## 2. Parámetros

| Parámetro | Valor | Origen |
|---|---|---|
| Cupo sencillo | 26 tarimas | `LBS_METRO_TRUCK_CAP` (`modulo2.vba:19`) |
| Cupo Full | Catálogo Mode Mix, normalmente 40 | `LBS_CatalogFullShipmentCap` (`modulo2.vba:5703`) |
| Cupo caja `a`/`b` | Cupo de embarque entre 2, normalmente 20 | `LBS_FullBoxCapForRow` (`modulo2.vba:5513`) |
| Piso de llenado | 80 % del cupo = 21 de 26 | `LBS_CHEDRAUI_MIN_FILL` (`modulo2.vba:81`) |
| Rescate de pedido | 80 % del cupo | `LBS_CHEDRAUI_RESCUE_MIN_FILL` (`modulo2.vba:56`) |
| Altura máxima de unidad | 1.60 m | `LBS_ChainMaxUnitHeightM` (`modulo2.vba:11908-11912`) |
| Peso máximo | 29 t | `SK_MAX_PESO_KG` (`modulo2.vba:10`) |

Comentarios originales:

```
' LBS - CHEDRAUI: piso de llenado post-consolidacion (cap 26 -> piso 21 = 80%).
' Independent of EFICIENCIA POR CADENA (AR gate); truck floor is always 80%.
Private Const LBS_CHEDRAUI_MIN_FILL As Double = 0.8
```

```
' LBS - CHEDRAUI: rescate de pedido todo-fallido si la carga combinada >= este % del cap.
Private Const LBS_CHEDRAUI_RESCUE_MIN_FILL As Double = 0.8
```

La segunda línea del primer comentario es una precisión útil: **el piso del 80 % no se puede
ajustar desde la hoja `EFICIENCIA POR CADENA`.** Ese porcentaje solo mueve el gate de `AR`.
Para cambiar el piso del camión hay que editar la constante.

Chedraui es la única cadena donde el piso y el umbral de rescate coinciden en 80 %. En
Walmart son distintos (40 % y 50 %), porque el rescate ahí es más permisivo que el piso.

## 3. Reglas de negocio

**Gate por destino.** La llave del gate es el grupo de consolidación más la vigencia
(`LBS_ChedrauiDestPedidoKey`, `tms_fg14/modulo2.vba:2707`):

```
' CHEDRAUI: dest (sin origen ni pedido) para anclas de charolas cross-planta.
' Minuta #20: incluye vigencia para no mezclar pedidos con distinta col R en un camion.
```

La llave omite deliberadamente el origen y el pedido, para que una charola pueda anclarse en
un camión de otra planta. Pero incluye la vigencia, y el comentario cita la minuta donde se
acordó: sin ella se mezclaban pedidos con vigencias distintas.

**Las tres reglas del gate.** El comentario de `LBS_ChedrauiPedidoEfficiencyGate`
(`tms_fg14/modulo2.vba:2714-2716`):

```
' CHEDRAUI: eficiencia por destino (origen|dest) antes del descarte folio-a-folio.
' Reglas: promedio ponderado por tarimas; rescate todo-fallido si cabe en un camion;
' merge de folios fallidos sobre folios que pasan cuando hay cupo.
```

En prosa:

1. **Promedio ponderado por tarimas.** No es el promedio simple de los fill rates: un folio
   de 20 tarimas pesa más que uno de 3. Evita que un folio chico y malo hunda un destino que
   en conjunto va bien.
2. **Rescate del destino todo-fallido.** Si ningún folio del destino pasa, pero juntos llegan
   al 80 % del cupo, se rescatan. La lógica: un camión al 80 % es un camión que vale la pena
   mandar, aunque cada pedido por separado se vea mal.
3. **Unión de folios fallidos sobre los que pasan.** Los folios que reprobaron se intentan
   subir a los camiones que aprobaron y tienen espacio, antes de descartarlos.

El gate corre **antes** del descarte folio a folio, y ese orden es el punto: sin él, cada
folio se evaluaría solo y se perderían destinos completos que en conjunto eran viables.

**Un origen por folio.** Chedraui está en la lista de cadenas con división correctiva de
orígenes mezclados (`LBS_IsMixedOriginSplitChain`, `tms_fg14/modulo2.vba:13730`), junto con
OXXO, Neto y Walmart. El comentario de la división
(`tms_fg14/modulo2.vba:13735`):

```
' CHEDRAUI/OXXO/NETO/Walmart: un origen (col L) por folio; mueve origenes extra a P-#### nuevos.
```

Un folio con producto de dos plantas se divide: el origen principal se queda y los demás se
mueven a folios `P-####` nuevos. Un camión sale de una planta; no puede tener dos orígenes.

**Integridad de folio.** `LBS_EnforceChedrauiFolioIntegrity` (`tms_fg14/modulo2.vba:13717`)
revalida que los folios queden consistentes después de las divisiones.

**División de folios sobre el cupo.** `LBS_SplitChedrauiOverCapFolios`
(`tms_fg14/modulo2.vba:14137`) y su función por folio
(`LBS_SplitChedrauiOneOverCapFolio`, `tms_fg14/modulo2.vba:13959`) parten los folios que
quedaron sobre el cupo, en lugar de recortarlos y perder producto.

**Acomodo de charolas.** `LBS_LayerChedrauiRestosCharolas` (`tms_fg14/modulo2.vba:16290`)
apila los restos en charolas. La llave del gate (sin origen ni pedido) es lo que permite que
una charola se ancle en un camión de otra planta.

**Camas compartidas y altura, pero no peso desde el catálogo.** Chedraui está en
`LBS_ChainAllowsSharedCamas` (`tms_fg14/modulo2.vba:6342`) y en
`LBS_ChainEnforcesUnitHeight` (`tms_fg14/modulo2.vba:11895`), pero **no** en
`LBS_ChainAppliesTihiAT` (`tms_fg14/modulo2.vba:6336`). Se le calcula altura y puede
compartir camas, pero el peso `AT` no se refresca desde el catálogo TI HI.

**Sin piso de llenado en el sentido de la lista transversal.** Chedraui tampoco está en
`LBS_IsTruckMinFillChain` (`tms_fg14/modulo2.vba:4434`), porque su piso lo aplica un
procedimiento propio: `LBS_EnforceChedrauiPostConsolidation`
(`tms_fg14/modulo2.vba:4593`). El efecto es el mismo, la implementación es separada.

## 4. Procedimientos

| Procedimiento | Línea | Qué hace |
|---|---|---|
| `LBS_IsChedrauiChain` | `6175` | Reconoce `CHEDRAUI` |
| `LBS_ChedrauiDestPedidoKey` | `2707` | Llave del gate: grupo más vigencia |
| `LBS_ChedrauiPedidoEfficiencyGate` | `2717` | El gate por destino, con promedio ponderado, rescate y unión |
| `LBS_PackChedrauiOneCkey` | `4172` | Empaca una llave de consolidación en camiones |
| `LBS_PackChedrauiToCapacity` | `4352` | Recorre todas las llaves empacando hasta el cupo |
| `LBS_EnforceChedrauiPostConsolidation` | `4593` | Aplica el piso del 80 % después de consolidar |
| `LBS_EnforceChedrauiFolioIntegrity` | `13717` | Revalida la consistencia de los folios |
| `LBS_SplitChedrauiMixedOriginFolios` | `13736` | Divide folios con orígenes mezclados |
| `LBS_SplitChedrauiOneOverCapFolio` | `13959` | Divide un folio sobre el cupo |
| `LBS_SplitChedrauiOverCapFolios` | `14137` | Recorre todos los folios sobre el cupo |
| `LBS_LayerChedrauiRestosCharolas` | `16290` | Acomoda los restos en charolas |
| `LBS_IsMixedOriginSplitChain` | `13730` | Lista de cadenas con división por origen |

Las fases correspondientes de `SummaryOptimizar` son `filtrar:split_chedraui`
(`tms_fg14/modulo2.vba:4867`), `filtrar:chedraui_gate` (`4933`), `filtrar:pack_chedraui`
(`4937`) y `filtrar:enforce_chedraui` (`4941`). Si `SummaryOptimizar` aborta en una de ellas,
el problema es de esta cadena.

## 5. Cómo validarlo

Fixtures en `tms_fg14/chedraiu/` (el nombre de la carpeta tiene un error de dedo que ya quedó
en los scripts):

| Archivo | Contenido |
|---|---|
| `sample.tsv` | Muestra de `Pedidos Surtidos` |
| `plan.tsv` | Plan de prueba |
| `armados.tsv` | Armados por SKU |

Scripts:

| Script | Qué comprueba |
|---|---|
| `scripts/validate_chedraui_restos.py` | El acomodo de restos en charolas |
| `scripts/validate_multi_origen_consolidation.py` | La división por orígenes mezclados |
| `scripts/validate_tihi_ao_all_chains.py` | El cálculo de `AO` desde TI HI |

## 6. Problemas conocidos y síntomas

| Síntoma | Causa probable | Dónde mirar |
|---|---|---|
| `Descartado por baja eficiencia` en camiones de 18 a 20 tarimas | Quedaron bajo el piso de 21, que es el más alto de las cadenas de autoservicio | `LBS_CHEDRAUI_MIN_FILL`. Es una constante, no se ajusta desde `EFICIENCIA POR CADENA` |
| Subir el porcentaje de `EFICIENCIA POR CADENA` no cambia nada | El piso del camión no depende de esa hoja | Es el comportamiento documentado en el comentario de la constante |
| Folios `P-####` nuevos que nadie pidió | La división por orígenes mezclados los creó | `LBS_SplitChedrauiMixedOriginFolios`. Es correcto: un camión no puede tener dos orígenes |
| Un destino completo descartado | Ningún folio pasó el gate y el rescate no llegó al 80 % | `LBS_ChedrauiPedidoEfficiencyGate`; revisar el umbral de la hoja `EFICIENCIA POR CADENA` |
| Cada destinatario en su propio camión, sin consolidar | Es el comportamiento por omisión: Chedraui no está en la tabla embebida de grupos | Mapear los destinatarios a mano en la hoja `Consolida` si el cliente quiere multi-stop |
| Pedidos con vigencias distintas en el mismo camión | No debería ocurrir: la llave del gate incluye vigencia desde la minuta #20 | Si aparece, es una regresión. Verificar `LBS_ChedrauiDestPedidoKey` |
| Peso `AT` distinto del catálogo TI HI | Chedraui no refresca `AT` desde el catálogo | Es deliberado. Ver `LBS_ChainAppliesTihiAT` |
| `REVISION MANUAL: tarimas >26` en Chedraui | La división por cupo no alcanzó | `LBS_SplitChedrauiOverCapFolios` |
