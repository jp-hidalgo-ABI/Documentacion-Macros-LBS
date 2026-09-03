[Volver a la macro de entrada](README.md) · [Índice general](../README.md)

# Runbook del operador — macro de entrada

Guía de ejecución diaria. Cada paso indica qué hace, qué revisar antes de avanzar y qué
hacer si algo sale mal. Los pasos marcados como **obligatorio** no se pueden saltar sin
arriesgar la corrida de LBS.

## Antes de empezar

1. **El libro debe estar en disco local.** Si está abierto desde una ruta web de SharePoint
   u OneDrive, las macros que escriben reportes fallan. Un aviso de
   "el libro necesita una carpeta local" confirma este problema
   (`merged/modulo6.vba:140-149`).
2. **Debe existir la subcarpeta `data\`** junto al libro. Ahí se escriben todos los reportes.
3. **El Plan debe estar cargado** con los pedidos del día, desde la fila 3.
4. **Verificar las banderas de la fila 1 del Plan.** En una corrida productiva, `V1` (modo
   prueba) y `W1` (alinear fechas STR/EBL) deben estar vacías o en `NO`. Ver
   [02-plan-y-parametros.md](02-plan-y-parametros.md).

## Paso 0 — Limpiar la corrida anterior (opcional)

| Botón | Qué limpia |
|---|---|
| `BorrarDatos` | Datos del `Plan` desde la fila 3 |
| `BorrarDatos2` | `Handling` |
| `BorrarDatos3` | `InventarioPaleteado` |
| `BorrarDatos4` | `InventarioDRP` |
| `BorrarDatos5` | `ProductSchedule` |

Se usan cuando se va a cargar un Plan y unos inventarios completamente nuevos. Si solo se
está reintentando la misma corrida, no hace falta limpiar: `ValidarDestinos` y `CambioOrigen`
son idempotentes sobre el mismo Plan.

No limpiar `Handling` a la ligera. `ValidarDestinos` la reconstruye desde el catálogo, pero
las banderas de planta `U:AC` que se hayan ajustado a mano se pierden.

## Paso 1 — `ValidarDestinos` (obligatorio)

`merged/modulo2.vba:976`

Es la compuerta de maestros. Hace tres cosas en orden:

1. Valida el `Catalogo Mode Mix` y lo aplica sobre `Handling`.
2. Recorre el Plan comparándolo contra `Logica Origen`, `TarimaPorDestino`, `Handling`,
   `ArmadoChep`, `ArmadoMadera` e `Inseparable`, y agrega las filas que falten donde puede.
3. Escribe los hallazgos que **no** pudo resolver en las filas 13 a 20 de la hoja
   `User Guide`.

### Qué revisar al terminar

Lo primero es la celda `Plan!X1`, que debe decir `ValidarDestinos: listo`
(`merged/modulo2.vba:1342`). `ValidarExportMEXKA` verifica exactamente ese texto más adelante
y aborta si no lo encuentra, así que si la macro se interrumpió a medias hay que volver a
correrla.

Después, las filas de hallazgos de `User Guide`:

| Fila | Qué significa que tenga contenido |
|---|---|
| 13 | Destinos o items sin `Logica Origen` |
| 15 | Destinos sin fila en `Handling` |
| 16 | Items que no están en LBS ni en `ArmadoMadera` |
| 17 | Destinos que no están en LBS ni tienen tipo de tarima |
| 18 | Pedidos con destino CHEP sin armado registrado |
| 19 | Destino con cadena inseparable |
| 20 | Carriles del Plan sin catálogo Mode Mix |

Cada una de estas filas se convierte en un hallazgo de `ValidarExportMEXKA`. Conviene
resolverlas ahora, no después de generar el export.

### Errores y qué hacer

**"Catalogo Mode Mix invalido. Revise la hoja y corrija antes de continuar."**
(`merged/modulo2.vba:1095`) — La hoja `Catalogo Mode Mix` tiene incoherencias entre el modo
(`Sencillo`/`Full`) y el tipo de equipo. La macro se detiene sin aplicar nada. Hay que
corregir la hoja; si el catálogo viene de origen, correr
`SincronizarCatalogoHomologado` para refrescarlo.

**Resumen con "Campo obligatorio sin informar (origen/destino/material/armado)"**
(`merged/modulo2.vba:1322`) — Hay renglones del Plan con celdas clave vacías. Se corrigen en
el Plan fuente.

**Resumen con "Fecha vigencia con mas de 40 dias vs hoy (formato DIA/MES/ANIO)"**
(`merged/modulo2.vba:1325`) — Casi siempre es una vigencia capturada en formato mes/día en
lugar de día/mes. Corregir el formato en el Plan.

**Resumen con "Fecha vigencia con formato incorrecto"** (`merged/modulo2.vba:1328`) — La
celda de vigencia no es una fecha. Suele venir de un copiado como texto.

## Paso 2 — `CambioOrigen` (obligatorio)

`merged/modulo3.vba:830`

Asigna la planta de origen en `Plan!F` consumiendo inventario paleteado, inventario DRP,
programa de producción, traspaleo e interplantas. Ver el detalle del algoritmo en
[04-cambio-origen.md](04-cambio-origen.md).

El progreso se ve en `Plan!Z1`.

### Qué revisar al terminar

Filtrar la columna `F` del Plan y buscar los tres valores que **no** son plantas:

| Valor en `Plan!F` | Significado | Qué hacer |
|---|---|---|
| `NoInventario` | No hay inventario en ninguna planta candidata para ese item | Revisar el inventario cargado, o aceptar que ese renglón no se embarca |
| `FaltaInventario` | Hay inventario pero no alcanza para el pedido completo | Igual que arriba; el renglón queda parcialmente cubierto |
| `FABRICA` | `CambioOrigen` no encontró ninguna planta asignable | Revisar `Logica Origen` para ese destino y las banderas de planta en `Handling` |

Ninguno de los tres es exportable. `ValidarExportMEXKA` los reporta uno por uno.

### Errores y qué hacer

**"Los valores de inventario son incorrectos. Revise InventarioPaleteado (col N) e InventarioDRP (col R)."**
(`merged/modulo3.vba:883`) — Las columnas de cantidad de los inventarios no tienen números.
Típicamente llegaron como texto desde el reporte fuente. Hay que convertirlas a número.

**Aviso "CambioOrigen - Plan fila 1"** (`merged/modulo3.vba:111`) — Hay un problema en los
parámetros de la fila 1 del Plan, normalmente el margen de días previgencia en `L1`.

**Aviso "CambioOrigen - Hojas vacias"** (`merged/modulo3.vba:200`) — Falta cargar uno de los
insumos obligatorios (inventarios o programa de producción).

**"Plan: cartonaje (col M) corrupto como fecha (ej. 0/01/1900)."**
(`merged/modulo1.vba:956-958`) — Esta es importante. La columna `M` del Plan trae fechas en
lugar de cantidades, casi siempre porque Excel reinterpretó un pegado. La macro **aborta**
y enumera hasta 12 filas afectadas con su pedido y material. Hay que corregir esas filas en
el Plan fuente; no se puede seguir.

## Paso 3 — Ajustes de `Handling` (opcional)

Solo si hay que forzar el modo de transporte de ciertos destinos por encima del catálogo.

| Botón | Efecto |
|---|---|
| `HandlingNormal` | Restaura los cupos activos `P:S` desde los base `L:O` |
| `HandlingNOFULL` | Anula el Full de los destinos Mayorista |
| `HandlingNOSENCILLO` | Anula el Sencillo de los destinos Mayorista |
| `HandlingNOFULLAUTO` | Anula el Full de los destinos Autoservicio |
| `HandlingNOSENCILLOAUTO` | Anula el Sencillo de los destinos Autoservicio |
| `HandlingOXXOFullMetro` | Deja solo Full en los destinos OXXO del metro |
| `HandlingFixConnectionOrigins` | Habilita banderas de planta para destinos con conexiones problemáticas |

El orden importa: `HandlingNormal` primero para partir de una base limpia, y después las
anulaciones. Correr una anulación dos veces no hace daño, pero correr `HandlingNormal`
después de una anulación la deshace.

Cuidado con un detalle: `ValidarDestinos` vuelve a aplicar el catálogo sobre `Handling`, así
que si se corre `ValidarDestinos` después de estos ajustes, los cupos vuelven al catálogo.
El orden correcto es `ValidarDestinos` primero, ajustes de `Handling` después.

## Paso 4 — Maestros nuevos (opcional)

Solo cuando `ValidarDestinos` reportó items o destinos que no existen en el sistema.

1. Capturar los items nuevos en `User Guide` (filas 14 y 18) y los destinos nuevos en la
   fila 15.
2. Correr `CreateItem` (`merged/modulo5.vba:59`) para los items. Genera
   `ItemMaster_*.xls` junto al libro.
3. Correr `CreateLocation` (`merged/modulo5.vba:166`) para los destinos. Genera
   `SiteMaster_*.xls` y `BodDetail_*.xls`.
4. Cargar esos tres archivos al sistema de maestros.

Estas plantillas son para la carga al sistema **externo**. No alimentan al MEX KA
directamente; para eso está la sincronización del paso 7.

## Paso 5 — `GeneraTemplates` (obligatorio)

`merged/modulo1.vba:33`

Es el paso largo. Arranca la cadena `dp` → `EqByLane` → `itemConnections` →
`LlenadoArchivoMDLBS`. El progreso se ve en `Plan!X1`.

Al final pide seleccionar el archivo MEX KA si no lo puede resolver solo. Lo busca
automáticamente con el patrón `MEX KA PLANTS*Restos*v*.xlsx` en la carpeta del libro y en
`data\` (`merged/modulo6.vba:851-909`).

### Qué revisar al terminar

El mensaje de éxito es
**"EQbyLane, StockTransReq, itemConn y connections (BodDetail/STR) completados."**
(`merged/modulo1.vba:2365`). Si aparece, el archivo MEX KA ya está escrito y guardado.

Antes de eso pueden aparecer dos avisos que **no detienen** el proceso pero sí requieren
atención:

**"Hay fechas de start mayores a la vigencia en el stockTransportRequests"**
(`merged/modulo1.vba:612`) — Hay pedidos cuya fecha de carga es posterior a su vigencia. LBS
los va a rechazar. Corregir `Plan!D` (fecha de carga) o `Plan!L` (vigencia).

**"Hay items con armado chep sin informar en stockTransportRequests, verificar el validador de datos para la orden ..."**
(`merged/modulo1.vba:624-625`) — Falta el armado CHEP de algún item. El renglón sale al export
con un armado incorrecto. Se resuelve agregando la fila en `ArmadoChep` o corriendo
`SincronizarCatalogoHomologado`.

### Errores y qué hacer

**"Llenado cancelado: el Plan no tiene datos para exportar."**
(`merged/modulo1.vba:2294`) — Después de aplicar los filtros de exportabilidad no quedó ni un
renglón. Revisar que `CambioOrigen` haya asignado plantas reales y que las vigencias sean
futuras.

**"No se pudo abrir el archivo MEX KA: ..."** (`merged/modulo1.vba:2372`) — El archivo está
abierto por otra persona, bloqueado, o la ruta cambió. Cerrarlo y reintentar.

**"No se seleccionó ningún archivo. El proceso se cancelará."**
(`merged/modulo1.vba:2250`) — Se canceló el diálogo. Las hojas internas del libro ya quedaron
construidas; se puede correr `LlenadoArchivoMDLBS` sola para completar sin repetir todo.

## Paso 6 — `ValidarExportMEXKA` (obligatorio)

`merged/modulo6.vba:156`

La compuerta antes de importar a LBS. Abre el MEX KA en solo lectura y corre unas 49
validaciones sobre el Plan y el export. El progreso se ve en `Plan!Y1`.

Al terminar muestra un resumen con los primeros cinco hallazgos y la ruta del reporte
completo: `data\mex_ka_validation_report.txt`.

**Regla de oro: cero hallazgos antes de importar a LBS.** Cada hallazgo que se ignora se
convierte en un `Order Failure` o en un embarque mal armado que hay que rehacer.

El catálogo completo de validaciones, con el mensaje exacto de cada una, su causa y su
remedio, está en [06-validaciones-mexka.md](06-validaciones-mexka.md).

Si el reporte dice
**"Ejecute ValidarDestinos (boton antes de Asignar Origen) antes de ValidarExportMEXKA"**
(`merged/modulo6.vba:197-198`), es que `Plan!X1` no tiene la marca de finalización. Volver al
paso 1.

## Paso 7 — Sincronizaciones (opcional)

Solo cuando `ValidarExportMEXKA` reportó faltantes de maestros en el export.

`SincronizarMaestrosBidireccional` (`merged/modulo7.vba:180`) completa en un paso los items,
destinos, `BodDetail`, `connections` y equipos que falten, tanto en el libro de planeación
como en el MEX KA, y al terminar vuelve a correr `ValidarExportMEXKA` automáticamente
(`merged/modulo7.vba:294-295`).

Si se quiere ir con cuidado, primero `Sync_Step0_DryRun` (`merged/modulo7.vba:145`) para ver
el reporte de lo que haría sin escribir nada, y después los pasos 1 a 6 uno por uno.
Detalle en [07-sincronizacion-maestros.md](07-sincronizacion-maestros.md).

## Paso 8 — Catálogo homologado (según calendario)

`SincronizarCatalogoHomologado` (`merged/modulo7.vba:1510`) baja el catálogo TI HI de
SharePoint y actualiza `ArmadoChep`, `ArmadoMadera`, `Plan!P` y el `itemPackage` del MEX KA.
Pide confirmación antes de escribir.

No es parte de la corrida diaria: se corre cuando el equipo de acomodos publica una versión
nueva del catálogo. Después de correrlo hay que **repetir desde el paso 1**, porque los
armados cambiaron y con ellos el conteo de tarimas de todo el Plan.

`AuditarArmadoVsCatalogo` (`merged/modulo7.vba:1528`) es la versión de solo lectura: compara
sin escribir y deja el resultado en `data\armado_vs_tihi_audit_report.txt`. Sirve para saber
si vale la pena sincronizar. Detalle en
[08-catalogo-homologado-tihi.md](08-catalogo-homologado-tihi.md).

## Resumen de celdas de estado

Durante la corrida, estas celdas de la fila 1 del `Plan` muestran el progreso. Si una se
queda congelada en un texto intermedio, la macro correspondiente se interrumpió.

| Celda | Macro que la escribe |
|---|---|
| `Plan!X1` | `ValidarDestinos` y `GeneraTemplates` |
| `Plan!Y1` | `ValidarExportMEXKA` |
| `Plan!Z1` | `CambioOrigen` |

## Reportes que quedan en `data\`

| Archivo | Lo escribe |
|---|---|
| `mex_ka_validation_report.txt` | `ValidarExportMEXKA` |
| `mex_ka_bidirectional_sync_report.txt` | `SincronizarMaestrosBidireccional` y los `Sync_Step*` |
| `catalogo_homologado_sync_report.txt` | `SincronizarCatalogoHomologado` |
| `armado_vs_tihi_audit_report.txt` | `AuditarArmadoVsCatalogo` |
