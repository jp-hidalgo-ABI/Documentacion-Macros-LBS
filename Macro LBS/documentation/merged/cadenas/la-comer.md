[Volver a cadenas](README.md) · [Macro de entrada](../README.md) · [Índice general](../../README.md)

# La Comer — macro de entrada

La Comer es la cadena que motivó la existencia de la tabla de sufijos, y la única que separa
sus embarques por una marca de confirmación capturada en el Plan.

## Identificación

`LBS_IsLaComerCadena` (`merged/modulo1.vba:994-998`):

```vba
Public Function LBS_IsLaComerCadena(ByVal cadena As String) As Boolean
    Dim c As String
    c = UCase$(Trim$(Replace(Replace(Replace(CStr(cadena), Chr$(160), " "), vbTab, ""), vbCrLf, "")))
    LBS_IsLaComerCadena = (c = "LA COMER")
End Function
```

Coincidencia exacta con `LA COMER`, con el espacio en medio.

## El sufijo `LAC` y por qué existe

`merged/modulo1.vba:967`

```vba
Case "LA COMER": LBS_CadenaItemSuffix = "LAC"
```

Esta es la primera entrada de la tabla y la razón de que la tabla exista. El comentario que
encabeza la función lo dice (`merged/modulo1.vba:962`):

```
' Sufijo LBS por cadena (evita Left(cadena,3) = "LA " para LA COMER; usar solo si falta fila ArmadoChep).
```

`LEFT("LA COMER", 3)` da `"LA "` — las letras L, A y un espacio. El `Item Id` resultante sería
`3003697_LA ` con espacio final, que:

- Es visualmente indistinguible de `3003697_LA` en una hoja de Excel.
- LBS lo trata como un item distinto de `3003697_LAC`.
- Genera duplicados en el maestro de items.
- Parte los embarques de un mismo producto en dos.

### Las tres validaciones que lo vigilan

`ValidarExportMEXKA` tiene tres reglas dedicadas a este problema:

| Mensaje | Qué detecta | Cita |
|---|---|---|
| `STR/items: sin sufijos rotos tipo SKU_LA (LA COMER usa LAC)` | Mensaje OK cuando todo está bien | `merged/modulo6.vba:3350` |
| `items: N SKU con duplicado *_LA y *_LAC (max 20 detallados)` | El maestro tiene las dos variantes del mismo SKU | `merged/modulo6.vba:3355` |
| `items: SKU duplicado LA COMER <sku> ...` | El caso individual del anterior | `merged/modulo6.vba:3404` |

Y el mensaje genérico de sufijo roto cita `LAC` como ejemplo del valor correcto
(`merged/modulo6.vba:2838`):
*"sufijo Left cadena incorrecto; usar LAC/OXX en ArmadoChep"*.

Cuando aparece el hallazgo de duplicados, la corrección es consolidar todo en `*_LAC` y
borrar las filas `*_LA` del maestro. No basta con corregir la macro: los items ya creados en
el maestro de LBS siguen ahí.

## La marca de confirmación R / C

`merged/modulo1.vba:565-571`

La Comer es la quinta prioridad de la cascada de `Consolidation Class`:

```vba
ElseIf isLaComer Then
    confRC = UCase$(Trim$(CStr(Sheets("Plan").Cells(planRow, "W").Value)))
    If confRC = "R" Or confRC = "C" Then
        Sheets("stockTransportRequests").Cells(i, "AB").Value = laneId & "_" & confRC
    Else
        Sheets("stockTransportRequests").Cells(i, "AB").Value = laneId
    End If
```

La columna `W` del Plan (`No/Confirmación`) trae una letra que distingue dos tipos de
embarque. Cuando vale `R` o `C`, la clase de consolidación se vuelve
`<carril>_<letra>`, de modo que LBS no puede juntar en el mismo camión un pedido `R` con
uno `C` aunque vayan al mismo destino desde la misma planta.

Cualquier otro valor de `Plan!W`, incluido el vacío, deja la clase como el carril simple.
Es decir, la separación es **opcional por renglón**: solo aplica cuando el área comercial
capturó la letra.

Del lado de la macro de salida, esta separación se mantiene: `LBS_SplitLaComerRC` divide los
folios que mezclan `R` y `C`.

## Caja seca en sencillo

La Comer aparece nombrada en el comentario de `LBS_ForceCatalogSencilloCajaSeca`
(`merged/modulo4.vba:172-174`):

```
' Catalog Sencillo Caja Seca: fold leftover encortinado cups (Q/M) into caja (P/L).
' Prevents EquipmentByLane Z1290 when Mode Mix/Especializado has no Sencillo Encortinado
' (e.g. LA COMER City Fresko Vallejo = Caja Seca 26 only).
```

Es un caso concreto: el destino de City Fresko Vallejo tiene en el catálogo únicamente
sencillo de caja seca con 26 tarimas, sin variante encortinada. Sin esta corrección, un cupo
residual en la columna `Q` de `Handling` habría hecho que `EqByLane` emitiera un equipo
`Z1290` que el catálogo no contempla.

La función es general —aplica a cualquier destino cuyo catálogo tenga sencillo sin
encortinado— pero el caso que la motivó fue este.

## Cómo validarlo

| Qué revisar | Dónde |
|---|---|
| Que no existan items `*_LA` | Reporte de `ValidarExportMEXKA`, mensajes `items: ... duplicado *_LA y *_LAC` |
| Que la separación R/C esté en la clase de consolidación | Columna `AB` del STR: debe verse `<carril>_R` o `<carril>_C` |
| Que no aparezca `Z1290` en carriles de La Comer | Hoja `equipmentByLaneByDay` del export |

Fixtures del lado de salida: `tms_fg14/lacomer/*.tsv`, con los scripts
`scripts/validate_lacomer*.py`.

## Problemas conocidos

**Aparecen items `*_LA` en el maestro.** Son residuos de exports anteriores a la tabla de
sufijos. Hay que borrarlos del maestro de LBS, no solo del export.

**Pedidos `R` y `C` viajan juntos.** Verificar que `Plan!W` tenga exactamente la letra, en
mayúscula o minúscula (la macro la convierte con `UCase$`) y sin caracteres extra. Un valor
como `R ` con espacio sí funciona porque hay un `Trim$`, pero `Ruta` no.

**Un renglón de La Comer usa `LEFT(cadena,3)`.** Confirmar que `Plan!G` sea exactamente
`LA COMER`. Un valor como `LACOMER` sin espacio no dispara la regla y produce el sufijo
`LAC`... que casualmente es el correcto. Pero `La Comer SA` produciría `La `.

## En la macro de salida

Del lado de salida La Comer tiene cupo de 26 tarimas, piso de llenado del 80%, altura máxima
de 1.6 m, división por R/C, anclas de sándwich propias y apertura de camiones desde
No Planeado. Ver
[../../tms_fg14/cadenas/la-comer.md](../../tms_fg14/cadenas/la-comer.md).
