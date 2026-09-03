[Volver al índice](README.md)

# Glosario

Terminología compartida por las dos macros. Los términos están en español porque así
aparecen en las hojas y en los comentarios del código; donde el código usa el término en
inglés se indica.

## Unidades de carga

**Cartonaje** — Cantidad de cajas (cartones) de un SKU. Es la unidad en la que el área
comercial captura el pedido. En el Plan es `Cartonaje Total` (columna `M`); en
`Pedidos Surtidos` es la columna `S`. Todo el motor de armado convierte cartonaje a tarimas
y de vuelta, y las inconsistencias entre ambos son la fuente más común de errores.

**Armado** — Cajas que caben en una tarima completa para una combinación cadena + SKU.
Se toma del catálogo homologado. Es el divisor que convierte cartonaje en tarimas:
`tarimas = cartonaje / armado`. En el Plan es `Armado de Tarimas` (columna `P`); en
`Pedidos Surtidos` es la columna `V`. Existen dos valores por SKU según el tipo de tarima:
`ArmadoChep` y `ArmadoMadera`.

**Tarima** — Pallet. Una tarima completa contiene exactamente `armado` cajas.
En `Pedidos Surtidos` la columna `T` cuenta tarimas completas.

**Resto** (o **restos**) — El sobrante de cartonaje que no llena una tarima completa:
`resto = cartonaje MOD armado`. En `Pedidos Surtidos` es la columna `U`. Los restos son el
problema central del armado: ocupan espacio físico en el camión pero no son una tarima
completa, así que hay que apilarlos o dejarlos fuera.

**Charola** — Bandeja. Una fila de puros restos de charolas se reconoce por `T=0`, `W=0`,
`U>0` (`tms_fg14/modulo2.vba:6087`). Estas filas "no consumen cupo de tarima pero deben
subir al camión del mismo pedido cuando hay ancla Programado"
(`tms_fg14/modulo2.vba:6003-6004`).

**TI / HI** — Parámetros del catálogo de acomodos. `TI` es el número de cajas por cama
(capa horizontal); `HI` es el número de camas por tarima. El producto `TI x HI` es el armado
teórico. El motor los lee de la hoja `TI HI` del propio libro, cacheados como
`Array(TI, HI, AltoCm, AltArmCm, PesoCase)` (`tms_fg14/modulo2.vba:89`).

**Cama** — Capa horizontal de cajas dentro de una tarima. Dos SKU distintos pueden
compartir cama si su altura difiere en 0.2 cm o menos, y en ese caso el clúster mixto usa
`TI - 1` cajas por cama (`tms_fg14/modulo2.vba:97-98`). Aplica solo a las cadenas de
`LBS_ChainAllowsSharedCamas`.

**Sándwich** — Dos o más unidades apiladas en la misma posición de piso del camión.
La fila **ancla** del sándwich lleva la marca en la columna `W` (`W=1` marca el slot físico)
y las capas superiores van con `T=0`, porque la tarima ya se contó en `W`
(`tms_fg14/modulo2.vba:1310`, `5963`). El motivo `LBS_SANDWICH` en la columna `AG` identifica
estas filas. Un sándwich mal anclado produce "tarimas fantasma" que descuadran el conteo del
camión.

## Camión y embarque

**Folio** (también **confirmación**) — Identificador del camión. Vive en la columna `AD` de
`Pedidos Surtidos` con formato `DDMMYYYYP-<número de shipment>`. Cuando un Full se parte en
dos cajas, se le agregan sufijos de letra: `...-1234a` y `...-1234b`.

**Shipment** — El embarque tal como lo devolvió LBS. Un shipment puede convertirse en uno o
varios folios después de `PartirTarimasFULL`.

**Cupo** — Capacidad máxima. Se usa en dos niveles distintos que conviene no confundir:
- **Cupo de camión** (truck cap): tarimas que caben en un camión físico. Por ejemplo 26 para
  el metro de Soriana/City Club, 28 para los grupos Walmart.
- **Cupo de shipment** (shipment cap): tarimas que LBS permite en un embarque antes de
  partirlo. Por ejemplo 40 para un Full de COMEXTRA, 36 para un Full de OXXO.

**Sencillo / Full** — Los dos modos de transporte. `Sencillo` es un camión sencillo;
`Full` es un tractocamión con dos cajas (`a` y `b`), cada una de aproximadamente la mitad del
cupo total. La elección por carril viene del catálogo (ver *Mode Mix*).

**Mode Mix** — El catálogo que define, por carril origen-destino, si se programa `Sencillo`
o `Full`, qué tipo de equipo se usa, si es especializado (encortinado) y cuál es el
`Pallets Max`. Vive en la hoja `Catalogo Mode Mix` y lo consultan las dos macros.

**Caja seca / Encortinado** — Los dos tipos de carrocería. El catálogo lo expresa en la
columna de especializado; algunas cadenas fuerzan caja seca sin importar el catálogo
(por ejemplo OXXO sencillo y La Comer).

**Llave de consolidación** (`ckey`) — La llave que decide qué filas pueden viajar juntas en
el mismo camión. Se construye con planta, grupo de consolidación, marca R/C y vigencia
(`LBS_ConsolidaKey`, `tms_fg14/modulo2.vba:5444-5474`). El grupo de consolidación traduce el
`Destinatario` a un grupo multi-stop mediante la hoja `Consolida`.

**Piso de llenado** (min fill) — Porcentaje mínimo del cupo que debe alcanzar un camión para
que se considere viable. Si queda por debajo, sus filas se descartan con motivo
"Descartado por baja eficiencia". Cada cadena tiene su piso: 40% Walmart, 70% Soriana/City
Club, 80% Chedraui y La Comer, 90% OXXO y COMEXTRA.

**Fill rate / eficiencia** — Porcentaje de surtido del pedido que reportó LBS. Llega en la
columna `AR` de `Pedidos Surtidos` desde la hoja `Shipments`. Se compara contra el umbral de
la hoja `EFICIENCIA POR CADENA`. Ojo: es eficiencia de **pedido**, distinta del piso de
llenado del **camión**.

## Estados y motivos

**Programado** — La fila se va a embarcar. Valor de la columna `H` de `Pedidos Surtidos`.

**No planeado** — La fila no se embarca en esta corrida. Valor de la columna `H`.
Las filas No Planeado son la materia prima de las fases de remonte: `SummaryFallo` y
`LBS_ConsolidarRestos` intentan subirlas a camiones con espacio.

**Motivo** — La columna `AG` de `Pedidos Surtidos` explica por qué una fila quedó como
No Planeado o qué marca especial lleva. Los valores más frecuentes son
"Descartado por baja eficiencia", "No reportado por LBS", "Shippable qty is 0" y el marcador
`LBS_SANDWICH`.

**Revisión manual** — La columna `AV` de `Pedidos Surtidos` acumula advertencias que el
motor no pudo resolver por sí solo: exceso de tarimas, exceso de altura, exceso de peso.
Es la primera columna que hay que revisar antes de exportar a TMS.

**Vigencia** — Fecha límite del pedido. En el Plan es `Vigencia del Pedido` (columna `L`); en
`Pedidos Surtidos` es la columna `R`. Dos pedidos con vigencia distinta no pueden compartir
folio, y hay una fase dedicada a separarlos (`LBS_SplitMixedVigenciaFolios`).

## Origen y destino

**Planta** / **PC** — El origen. Se identifica con códigos `PC01`, `PC05`, `PC13`, etc.
En el Plan es la columna `F` (`Origen`), asignada por `CambioOrigen`.

**Destinatario** — El CEDIS o punto de entrega del cliente, identificado con un número de
9 dígitos (por ejemplo `400101621`). En el Plan es la columna `I`; en `Pedidos Surtidos`
la columna `O`.

**Carril** (lane) — Par origen-destino, escrito como `<origen>_<destino>`. Es la unidad de
configuración del catálogo Mode Mix y de la hoja `connections` del export.

**FG14** — Prefijo de los destinos que no tienen un ID numérico de CEDIS y se identifican por
cadena, con la forma `FG14_<slug de cadena>`. Aplica a los clientes de comercio electrónico y
mayoristas pequeños (Amazon, Rabbit, Mercado Libre, Rapiturbo).

**NoInventario / FaltaInventario / FABRICA** — Marcas que `CambioOrigen` escribe en la
columna `F` del Plan cuando no pudo asignar una planta real. Ninguna de las tres es
exportable a LBS, y `ValidarExportMEXKA` las bloquea.

## Archivos del pipeline

| Archivo | Quién lo produce | Quién lo consume |
|---|---|---|
| `MEX KA PLANTS_Restos_v*.xlsx` | `LlenadoArchivoMDLBS` (macro de entrada) | LBS |
| `Shipments (*).xlsx` | LBS | `ProcesoMacro` (macro de salida) |
| `STP Equipment Association (*).xlsx` | LBS | `ProcesoMacro` |
| `Order Failures (*).xlsx` | LBS | `ProcesoMacro` (opcional) |
| `Pallet Container Association (*).xlsx` | LBS | `ProcesoMacro` |
| `ProcessShipmentOrderCreate_DS_*.xls` | `CreaTMStemplate` (macro de salida) | TMS |
| `AdicionalesSAP*.csv` | `ExportarACSV` (macro de salida) | SAP |
| `ItemMaster_*.xls`, `SiteMaster_*.xls`, `BodDetail_*.xls` | Módulo 5 de la macro de entrada | Carga de maestros ABPP |

## Siglas

| Sigla | Significado |
|---|---|
| **STR** | `stockTransportRequests`. La hoja del export que lleva los pedidos a LBS. |
| **EBL** | `equipmentByLaneByDay`. La hoja del export que declara qué equipo está disponible por carril y por día. |
| **PCA** | `Pallet Container Association`. La salida de LBS que detalla cómo quedó armada cada unidad. |
| **STP** | Shipment Transport Plan. Prefijo de las salidas de LBS relacionadas con equipo y fallos. |
| **ABPP** | Nombre de las plantillas de maestros que se cargan al sistema (`ItemMaster_ABPP`, `SiteMaster_ABPP`, `BodDetail_ABPP`). |
| **DRP** | Distribution Requirements Planning. El inventario proyectado que consulta `CambioOrigen`. |
| **EXC28** | Grupo de excepción de Walmart para SKU que requieren tarima completa de 28. |
| **Cap35** | Regla de cupo de 35 tarimas por shipment para SKU listados de cadenas mayoristas. |
