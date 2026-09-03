[Volver a la macro de salida](README.md) · [Índice general](../README.md)

# Las salidas: TMS y SAP

Dos archivos, dos macros independientes:

| Macro | Código | Produce |
|---|---|---|
| `TMS` → `CreaTMStemplate` | [tms_fg14/modulo4.vba](../../tms_fg14/modulo4.vba) | `ProcessShipmentOrderCreate_DS_<fecha>.xls` |
    10|| `AdicionalesSAP` → `ConsolidarItems` → `ExportarACSV` | [tms_fg14/modulo6.vba](../../tms_fg14/modulo6.vba) | `Adicionales<fecha>.csv` |

Ambas leen `Pedidos Surtidos` y ambas se pueden ejecutar en cualquier orden. Los dos archivos
se guardan **en la misma carpeta que el libro**, con marca de tiempo en el nombre.

## `TMS` — el archivo para el sistema de transporte

### La hoja `Data`

`TMS` (`tms_fg14/modulo4.vba:1`) limpia las filas 2 a 10 000 de la hoja `Data` y las vuelve a
    20|llenar desde `Pedidos Surtidos`.

| Col `Data` | Contenido | Fuente |
|---|---|---|
| `A` | Número de embarque | `"SSE1_" & DDMMYY & "_" & hh:mm & "_" & (i-1)` |
| `B` | Cantidad | `PS!T` (**tarimas completas**) |
| `C` | El mismo número de embarque que `A` | idem |
| `F` | MITCC | `PS!AD` (el folio) |
| `G` | Grupo de carga | `DDMMYY & "_" & hh:mm & "_" & (i-1)` |
| `H` | Plan | La constante `8000` |
| `I` | Fecha de inicio | `PS!AK` como `MM/DD/YYYY 00:00:00` |
    30|| `J` | Fecha de fin | `PS!R` como `MM/DD/YYYY 02:00:00` |
| `K` | Ventana de inicio | `PS!AK` como `MM/DD/YYYY 00:01:00` |
| `L` | Ventana de fin | `PS!R` como `MM/DD/YYYY 23:59:00` |
| `M` | Origen | `PS!L` (la planta) |
| `N` | Destino | `PS!O` (el destinatario) |
| `O` | Item | `PS!AA` (el material sin sufijo) |
| `Q` | Descripción | `VLOOKUP` contra la hoja `ItemMaster` |

Todo en `tms_fg14/modulo4.vba:14-45`.

Cuatro puntos que conviene entender:
    40|
**La cantidad son tarimas, no cajas.** `Data!B` toma `PS!T`, las tarimas completas. No toma
`S` (cartonaje) ni `T + W`. TMS trabaja en unidades de tarima.

**Los cuatro campos de fecha se escriben como texto.** El rango `I:L` se formatea con
`NumberFormat = "@"` antes de escribir (`tms_fg14/modulo4.vba:35`), y las fechas se generan
con `Format(..., "MM/DD/YYYY hh:mm:ss")`. Es formato **estadounidense**: mes primero. El
destino espera ese formato, y escribirlo como texto evita que Excel lo reinterprete según el
locale de la máquina.

    50|**Las horas están fijas en el código.** `00:00:00` para el inicio, `02:00:00` para el fin,
`00:01:00` y `23:59:00` para la ventana. No salen de ningún parámetro.

**El número de fila forma parte del identificador.** `(i - 1)` se concatena al número de
embarque y al grupo de carga, así que el identificador depende de la posición de la fila en
`Pedidos Surtidos`. Reordenar la hoja cambia todos los identificadores.

### La descripción del item

`tms_fg14/modulo4.vba:47-53`:

    60|```vba
Worksheets("ItemMaster").Visible = True
Sheets("Data").Range("Q2:Q" & nFilasPedidos).Formula = _
    "=IFERROR(VLOOKUP(TEXT(Data!O2,""0""),ItemMaster!$A:$H,8,0),"""")"
Sheets("Data").Range("Q2:Q" & nFilasPedidos).Copy
Sheets("Data").Range("Q2:Q" & nFilasPedidos).PasteSpecial Paste:=xlPasteValues
Worksheets("ItemMaster").Visible = False
```

La hoja `ItemMaster` se hace visible solo durante la operación —VBA no puede escribir
fórmulas que apunten a una hoja oculta en todos los contextos— y la fórmula se convierte a
    70|valores de inmediato, para que el archivo exportado no lleve referencias externas.

El `TEXT(Data!O2,"0")` fuerza la conversión del material a texto sin decimales, porque
`ItemMaster!A` los guarda como texto y un `VLOOKUP` numérico contra texto no encuentra nada.

Se busca en la **columna 8** (`H`) de `ItemMaster!A:H`.

### El `Exit For` que corta el recorrido

Es la particularidad más importante de esta macro (`tms_fg14/modulo4.vba:14-45`):

    80|```vba
For i = 2 To nFilasPedidos
    If Sheets("Pedidos Surtidos").Cells(i, "H") = "Programado" Then
        ...
    Else
        Exit For
    End If
Next i
```

**El recorrido se detiene en la primera fila que no dice `Programado`.** No usa `GoTo` para
    90|saltarla: sale del ciclo.

Eso funciona porque el pipeline garantiza que las filas `Programado` están todas arriba y las
`No planeado` todas abajo. Es lo que hacen `LBS_MoveNoPlaneadoToEnd`
(`tms_fg14/modulo2.vba:5027`) y `LBS_CompactBlankPSRows`, cuyo comentario ahora se entiende
mejor:

```
' Drop cleared ghost rows so they cannot sit between Programado y No planeado.
```

**Consecuencia práctica: una fila en blanco o `No planeado` en medio del bloque `Programado`
   100|hace que todo lo que está debajo no llegue a TMS, sin ningún mensaje de error.** El síntoma
es un archivo con menos embarques de los esperados. Correr `TMS` sin haber corrido
`SummaryOptimizar` produce exactamente eso.

Otro detalle: el número de filas se determina con la columna `L`
(`tms_fg14/modulo4.vba:11`), no con `A` ni `AD`.

Y las filas de `Data` se escriben en el **mismo índice `i`** que la fila de origen. Si la fila
50 de `Pedidos Surtidos` es la primera `No planeado`, `Data` tiene datos en las filas 2 a 49 y
el resto vacío. No hay compactación.

   110|### `CreaTMStemplate` — escribir el archivo

`tms_fg14/modulo4.vba:57-92`. La secuencia:

1. Arma el nombre: `ThisWorkbook.path & "\ProcessShipmentOrderCreate_DS_" & Format(Now,
   "yyyy_mm_dd_hh_mm") & ".xls"`.
2. Hace visibles las hojas `Data`, `Info` y `Mapping`.
3. Crea un libro nuevo y **copia las tres hojas** en ese orden.
4. Guarda con `FileFormat:=56`, que es el formato `.xls` de Excel 97-2003.
5. Cierra el libro nuevo y vuelve a ocultar las tres hojas.

   120|Las tres hojas del archivo:

| Hoja | Contenido |
|---|---|
| `Data` | Los embarques |
| `Info` | Metadatos de la plantilla, esperados por el importador de TMS |
| `Mapping` | El mapeo de campos, también esperado por el importador |

`Info` y `Mapping` no las escribe ninguna macro: son fijas del libro. Si el importador de TMS
cambia su especificación, se editan a mano en el `.xlsm`.

   130|El formato `.xls` no es una elección estética: el importador de TMS lo requiere. `xlsx` (51)
o `xlsm` (52) no se aceptan.

`DisplayAlerts` se apaga durante el guardado (`tms_fg14/modulo4.vba:84-86`) porque guardar
como `.xls` desde un Excel moderno dispara un aviso de compatibilidad.

### Limitaciones a tener presentes

- **No hay manejo de error.** Si la carpeta no tiene permisos de escritura o el archivo
  anterior está abierto, la macro falla con el error de VBA sin explicación.
- **No hay mensaje de confirmación.** `CreaTMStemplate` no avisa nada al terminar. La única
   140|  forma de saber que funcionó es buscar el archivo en la carpeta.
- **El tope es 10 000 filas**, por el `Rows("2:10000").Clear` (`tms_fg14/modulo4.vba:8`). Una
  corrida con más de 10 000 filas `Programado` dejaría residuo de la corrida anterior por
  debajo de esa fila.
- **La marca de tiempo llega al minuto.** Dos corridas en el mismo minuto sobrescriben el
  archivo sin aviso.

## `AdicionalesSAP` — el CSV para SAP

`AdicionalesSAP` (`tms_fg14/modulo6.vba:1`) genera un archivo con formato de encabezado y
   150|detalle: filas `H` con los datos del envío y filas `D` con las líneas de material.

### El preámbulo

Dos cosas antes de empezar (`tms_fg14/modulo6.vba:20-21`):

```vba
LBS_ResetConsMap   ' releer la hoja "Consolida" por si se edito
Call LBS_WriteTarimaTotalsTW(ws1, nFilasPedidos)
```

Se relee la hoja `Consolida` y se recalculan los totales de tarimas. El comentario de la
   160|segunda línea aclara la consistencia buscada:

```
' LBS - Total tarimas (col Z) por folio = SUM(T)+SUM(W), igual que SummaryOK / Fallos.
```

### Un camión = un folio

Es el cambio de diseño más importante de este módulo, y su comentario cuenta la historia
completa (`tms_fg14/modulo6.vba:24-27`):

   170|```
' Un camion = un folio (AD). Antes se abria un encabezado H por cada X=1, lo que
' partia en dos un camion consolidado: dos cadenas del mismo grupo (City Club + Soriana,
' o Walmart BA + SC) en el mismo folio llegan con dos X=1 y se generaban dos envios.
' Ahora se agrupa por folio.
```

La versión anterior abría una fila `H` cada vez que encontraba un `1` en la columna `X`. Pero
cuando dos cadenas del mismo grupo de consolidación comparten camión, la columna `X` trae dos
marcas, y se generaban dos envíos a SAP para un solo camión físico.

   180|La versión actual agrupa por `AD` y **normaliza `X` de paso**
(`tms_fg14/modulo6.vba:50-56`):

```vba
' Normalizar marca de camion (col X): solo la primera fila del folio = 1,
' el resto en blanco, para que no "reinicie" en 1 al unir City Club + Soriana.
ws1.Cells(firstRow, "X").Value = 1
For Each rr In rowsCol
    If CLng(rr) <> firstRow Then ws1.Cells(CLng(rr), "X").Value = ""
Next rr
```

   190|Es el tercer lugar donde `X` se reescribe. Ver la advertencia correspondiente en
[03-pedidos-surtidos.md](03-pedidos-surtidos.md).

### La fila `H` — el encabezado

Una por folio (`tms_fg14/modulo6.vba:59-87`):

| Col | Valor |
|---|---|
| `A` | `"H"` |
| `B` | `"2001"` |
   200|| `C` | `"ZINT"`, o `"UB"` si el origen es `PC01` |
| `D` | El centro SAP, según la planta |
| `E` | `"OC01"` |
| `F` | `"T01"` |
| `G` | Vacío |

El mapeo de planta a centro SAP:

| `PS!L` empieza con | `Data!D` | Nota |
|---|---|---|
| `PC01` | `2001` | **También cambia `C` de `ZINT` a `UB`** |
   210|| `PC03` | `2003` | |
| `PC05` | `2005` | |
| `PC07` | `2007` | |
| `PC11` | `2011` | |
| `PC13` | `2013` | |
| `PC19` | `2019` | |
| `PC29` | `2029` | |
| `PC23` | **`2021`** | La única que no sigue el patrón |
| cualquier otro | `2000` | Respaldo |

   220|Dos excepciones que hay que conocer:

**`PC01` usa tipo de pedido `UB` en lugar de `ZINT`.** `UB` es traslado entre plantas del
mismo centro; `ZINT` es interno. `PC01` es la única planta con ese tratamiento.

**`PC23` mapea a `2021`, no a `2023`.** Todas las demás plantas siguen el patrón
`PC<nn>` → `20<nn>`. Es la excepción, y está escrita a mano en el código
(`tms_fg14/modulo6.vba:79-80`).

El respaldo `2000` significa que una planta nueva sale al CSV con un centro inexistente y SAP
la rechaza. **Agregar una planta requiere editar este bloque de `ElseIf`.**
   230|
### Las filas `D` — el detalle

Una por cada fila del folio (`tms_fg14/modulo6.vba:90-99`):

| Col | Valor |
|---|---|
| `A` | `"D"` |
| `B` | `"10"` |
| `C` | `PS!AA` (el material) |
| `D` | `"BX14"` (la unidad de medida) |
| `E` | `PS!S` (**el cartonaje**) |
   240|| `F` | Vacío |
| `G` | `PS!AK` como `yyyymmdd` |

Nótese la diferencia con TMS: **SAP recibe cajas (`S`), TMS recibe tarimas (`T`).** Son dos
unidades distintas para el mismo camión, y es correcto: SAP factura cajas y TMS mueve tarimas.

La unidad de medida `BX14` está fija en el código.

El número de filas se determina con la columna `S` (`tms_fg14/modulo6.vba:17`), a diferencia
de `TMS` que usa `L`.

   250|### `ConsolidarItems` — sumar duplicados

`tms_fg14/modulo6.vba:108-150`. Recorre el archivo bloque por bloque y, dentro de cada bloque
`H`, suma las cantidades de las filas `D` que repiten material y elimina las duplicadas.

El ciclo interno se detiene al encontrar el siguiente `H`
(`tms_fg14/modulo6.vba:140-142`), así que la consolidación nunca cruza de un camión a otro.

Es un algoritmo cuadrático: para cada fila `D` recorre el resto del archivo. Con un CSV grande
se nota. Ajusta `lastRow` y el índice `k` al borrar filas
(`tms_fg14/modulo6.vba:138-139`), que es el detalle que hace correcto el recorrido hacia
   260|adelante.

Es la única macro de los dos módulos de salida sin comentarios explicativos, lo que sugiere
que es código original poco tocado.

### `ExportarACSV` — escribir el archivo

`tms_fg14/modulo6.vba:152-202`. Nombre del archivo:

```
<carpeta del libro>\Adicionales<yyyy_mm_dd_hh_mm>.csv
```

   270|A diferencia de `CreaTMStemplate`, esta sí tiene manejo de error y sí confirma:

```
El archivo CSV se ha generado en la ruta: <ruta completa>
```

Y valida que la ruta no esté vacía antes de intentar guardar
(`tms_fg14/modulo6.vba:168-171`):

```
Error: No se pudo obtener la ruta del archivo.
```

   280|Esa validación es la misma defensa que la macro de entrada aplica con
`LBS_MsgWorkbookNeedsLocalFolder`: `ThisWorkbook.path` viene vacío cuando el libro se abrió
desde un adjunto de correo o desde una vista de SharePoint sin sincronizar. Ver
[merged/01-runbook-operador.md](../merged/01-runbook-operador.md).

El error se reporta por `MsgBox` y por `Debug.Print`
(`tms_fg14/modulo6.vba:199-201`), así que también queda en la ventana Inmediato del editor de
VBA.

Se copia `ws.UsedRange` a un libro nuevo con una sola hoja llamada `Adicionales` y se guarda
   290|con `FileFormat:=xlCSV`.

**Ojo con el separador.** `xlCSV` usa el separador de lista de la configuración regional de
Windows. En español de México suele ser la coma, pero si la máquina está configurada con
punto y coma, el archivo sale con punto y coma. Vale la pena verificar el primer archivo
generado en una máquina nueva.

## Comparación de las dos salidas

| | TMS | SAP |
|---|---|---|
   300|| Macro | `TMS` (`modulo4.vba:1`) | `AdicionalesSAP` (`modulo6.vba:1`) |
| Formato | `.xls` (FileFormat 56) | `.csv` (xlCSV) |
| Nombre | `ProcessShipmentOrderCreate_DS_<fecha>.xls` | `Adicionales<fecha>.csv` |
| Granularidad | Una fila por fila `Programado` | Una `H` por folio, una `D` por fila |
| Cantidad | `PS!T` (tarimas) | `PS!S` (cajas) |
| Filas que lee | Se **detiene** en la primera no `Programado` | Recorre toda la hoja filtrando |
| Última fila por | `PS!L` | `PS!S` |
| Consolida duplicados | No | Sí, `ConsolidarItems` |
| Confirma al terminar | **No** | Sí, con la ruta |
| Maneja errores | No | Sí |
   310|| Relee `Consolida` | No | Sí, `LBS_ResetConsMap` |

La diferencia en cómo recorren la hoja es la más relevante: `AdicionalesSAP` es robusto a
filas `No planeado` intercaladas y `TMS` no.

## Qué revisar después

### Del archivo de TMS

1. Que exista, con la marca de tiempo de la corrida.
2. **Contar las filas de `Data` y compararlas con las filas `Programado` de
   320|   `Pedidos Surtidos`.** Si son menos, hay una fila que cortó el recorrido.
3. Que `Data!Q` (descripción) no venga vacía en masa. Si lo está, falta el material en
   `ItemMaster`.
4. Que las fechas de `I:L` estén en formato `MM/DD/YYYY`.
5. Que el archivo tenga las tres hojas: `Data`, `Info` y `Mapping`.

### Del CSV de SAP

1. Leer la ruta del mensaje de confirmación.
2. Que haya **una** fila `H` por camión. Dos `H` con el mismo folio indica que la
   330|   normalización de `X` falló.
3. Que la columna `D` de las filas `H` no traiga `2000`. Ese es el respaldo y significa una
   planta sin mapear.
4. Que las cantidades de las filas `D` sumen el cartonaje esperado del camión.
5. Verificar el separador de campos del archivo.
