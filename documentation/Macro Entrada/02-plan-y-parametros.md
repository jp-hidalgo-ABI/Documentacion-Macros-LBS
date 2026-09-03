[Volver a la macro de entrada](README.md) · [Índice general](../README.md)

# La hoja `Plan` y los parámetros de corrida

La hoja `Plan` es la entrada única de toda la macro. Los datos empiezan en la **fila 3**
(las filas 1 y 2 son encabezado y parámetros) y ocupan las columnas **A a AF**, 32 columnas.

## Mapa de columnas

Encabezados tomados de [merged/plan.tsv](../../merged/plan.tsv), con el rol que cada columna
juega en el código. La columna "Fuente" indica quién la llena: **Comercial** si viene del
Plan cargado, **Macro** si la escribe la propia macro.

| Col | Encabezado | Fuente | Rol en el código |
|---|---|---|---|
| `A` | `LLAVE 2` | Macro | Tipo de cliente (`Mayorista` / `Autoservicio`). Lo llena `CO_FillPlanTipoCliente` cruzando el destinatario contra `TarimaPorDestino!E`; si no lo encuentra escribe `No informado` (`merged/modulo3.vba:115-135`) |
| `B` | `Semana` | Comercial | Solo informativo |
| `C` | `Fecha de Orden de Compra` | Comercial | Solo informativo |
| `D` | `Fecha de Carga` | Comercial | Fecha de inicio del STR (`Start Time`). Ver `merged/modulo1.vba:324-328` |
| `E` | `Hora de Carga` | Comercial | Solo informativo |
| `F` | **`Origen`** | Macro | La planta asignada (`PC01`, `PC29`, …) o una de las tres marcas de fallo: `NoInventario`, `FaltaInventario`, `FABRICA`. La escribe `CambioOrigen` |
| `G` | **`Cadena`** | Comercial | La cadena comercial. Determina el sufijo del item, el slug FG14 y prácticamente todas las reglas de negocio |
| `H` | `CEDIS` | Comercial | Nombre del CEDIS destino. Se usa en `SiteMaster` y en la detección de destinos OXXO del metro |
| `I` | **`Destinatario`** | Comercial | ID numérico de 9 dígitos del punto de entrega. Es la llave de `Handling`, `TarimaPorDestino` y del catálogo |
| `J` | **`Pedido`** | Comercial | Número de pedido. Es la llave que une el Plan con `Pedidos Surtidos` en la macro de salida, y la base del `Group Id` del STR |
| `K` | `Cotización SAP` | Comercial | Solo informativo |
| `L` | **`Vigencia del Pedido`** | Comercial | Fecha límite. Es el `End Time` del STR y el techo de la ventana de disponibilidad de equipo |
| `M` | **`Cartonaje Total`** | Comercial | Cajas pedidas. Es la columna que más problemas causa: si Excel la reinterpreta como fecha, la macro aborta (ver más abajo) |
| `N` | `Tarimas completas` | Comercial | Conteo de tarimas completas. Se usa en la regla de split de restos de Walmart |
| `O` | `Restos de charolas` | Comercial | Sobrante que no llena tarima. Alimenta la excepción de tarimas de Walmart |
| `P` | **`Armado de Tarimas`** | Macro / Catálogo | Cajas por tarima. Lo actualiza `SincronizarCatalogoHomologado` desde TI HI |
| `Q` `R` `S` | `Tipo de tarima1..3` | Comercial | Heredado, informativo |
| `T` | **`Material SAP`** | Comercial | El SKU. Base del `Item Id` que se construye para LBS |
| `U` | `UPC` | Comercial | Solo informativo |
| `V` | `Presentacion` | Comercial | Descripción del item. Se usa al generar `ItemMaster_ABPP` |
| `W` | `No/Confirmación` | Comercial | Para LA COMER es la marca `R` o `C` que separa los folios. Se convierte en parte del `Group Id` |
| `X` | `Fecha` | Comercial | Informativo. **Ojo: `X1` es celda de progreso, no encabezado de datos** |
| `Y` | `Llenado` | Comercial | Informativo |
| `Z` | `Priorización` | Comercial | Informativo |
| `AA` | **`Total + BACK ORDER`** | Comercial | La cantidad que se exporta a LBS como `Quantity`. `CambioOrigen` la sincroniza desde `M` |
| `AB` | `Logica Origen` | Macro | Planta candidata primaria, resuelta contra la hoja `Logica Origen` |
| `AC` | `Logica Origen2` | Macro | Planta candidata secundaria |
| `AD` | **`Tipo de tarima`** | Macro / Catálogo | `Tarima Chep`, `Tarima Plastica` o `Tarima Madera`. Decide qué armado se usa y cómo se construye el `Item Id` |
| `AE` | `Tipo de Cliente` | Macro | Clasificación que se siembra en `TarimaPorDestino` |
| `AF` | `Validador` | Macro | Resultado de la validación del renglón |

### Ejemplo de renglón real

Del archivo de muestra [merged/plan.tsv](../../merged/plan.tsv), fila 2:

| Col | Valor |
|---|---|
| `A` | `MAYORISTA` |
| `F` | `PC29` |
| `G` | `Super Ofertas` |
| `H` | `Super Ofertas Puebla` |
| `I` | `400092931` |
| `J` | `SO Puebla - Ago 04 - 26` |
| `L` | `31/08/2026` |
| `M` | `8,400` |
| `N` | `70` |
| `O` | `0` |
| `P` | `120` |
| `T` | `3003697` |
| `V` | `VICTORIA BOTE 4P H-C 24/473ML CM` |
| `AA` | `8400` |
| `AD` | `Tarima Madera` |
| `AE` | `Mayoristas` |

8 400 cajas con armado de 120 dan 70 tarimas completas y 0 de resto, que es exactamente lo
que dicen `N` y `O`. Cuando esa aritmética no cuadra es señal de que el armado del catálogo
cambió pero el Plan no se recalculó.

## La fila 1: parámetros y estado

La fila 1 no es encabezado de datos. Es un panel de configuración y de progreso.

### Parámetros de entrada

| Celda | Nombre | Valores aceptados | Valor por omisión | Qué controla |
|---|---|---|---|---|
| `L1` | Días de margen vigencia-producción | Un número | `2` | Cuántos días antes de la vigencia puede programarse la producción. Si no es numérico, se usa 2 y se avisa (`merged/modulo3.vba:99-107`) |
| `V1` | **Modo prueba** | `SI` `S` `1` `TRUE` `ON` `YES` `Y` para activar; `NO` `N` `0` `FALSE` `OFF` para desactivar | Desactivado (`LBS_TESTMODE_DEFAULT`, `merged/modulo1.vba:14`) | Ver abajo |
| `W1` | **Alinear fechas STR + EBL** | Los mismos valores que `V1` | Desactivado (`LBS_ALIGN_STR_EQUIP_DATES_DEFAULT`, `merged/modulo1.vba:7`) | Ver abajo |
| `F1` | Fórmula de control | — | — | `CambioOrigen` avisa si contiene `#REF!` (`merged/modulo3.vba:92-97`) |

Cualquier otro texto en `V1` o `W1` se interpreta como el valor por omisión, no como error
(`merged/modulo1.vba:824-831`, `852-859`).

### Celdas de progreso

| Celda | Macro que la escribe | Cita |
|---|---|---|
| `X1` | `ValidarDestinos` y `GeneraTemplates` | `merged/modulo1.vba:22`, `merged/modulo2.vba:702-706` |
| `Y1` | `ValidarExportMEXKA` | `merged/modulo6.vba:306-309` |
| `Z1` | `CambioOrigen` | `merged/modulo3.vba:74-77` |

`X1` no es solo cosmético. `ValidarExportMEXKA` exige que contenga el texto
`ValidarDestinos: listo` antes de correr, y si no lo encuentra reporta un hallazgo y aborta
(`merged/modulo6.vba:3413-3417`). Esa marca la escribe `ValidarDestinos` al terminar bien
(`merged/modulo2.vba:1342`).

## Modo prueba (`Plan!V1`)

El comentario del código lo describe así (`merged/modulo1.vba:9-13`):

```
' TEST MODE toggle: Plan!V1 = SI / 1 / TRUE habilita corridas de prueba completas:
'   - Equipo: misma union STR+Plan; si W1 off, EBL End al menos Date+120; start no despues de hoy
'   - CambioOrigen: inventario infinito (no marca NoInventario/FaltaInventario y omite la
'     validacion de inventario numerico) -> todas las filas con origen prioritario se exportan.
' Vacio o NO = comportamiento productivo normal (respeta inventario y Handling).
```

En concreto, con modo prueba activo:

- **`CambioOrigen` finge inventario infinito.** No marca `NoInventario` ni `FaltaInventario`,
  y se salta la validación de que las columnas de inventario sean numéricas
  (`merged/modulo3.vba:654`, `879`). Todos los renglones con una planta prioritaria salen al
  export.
- **El equipo siempre está disponible.** La ventana de `equipmentByLaneByDay` se extiende
  al menos hasta hoy + 120 días, y el inicio nunca queda después de hoy.
- **El límite de carga por carril sube a 999** (`LBS_TESTMODE_LOADLIMIT`,
  `merged/modulo1.vba:15`).

**Un export generado en modo prueba no se debe importar a LBS.** Contiene pedidos sin
respaldo de inventario real. La bandera es fácil de olvidar encendida, así que conviene
revisar `V1` como parte del arranque de cada corrida.

## Alinear fechas STR + EBL (`Plan!W1`)

Comentario original (`merged/modulo1.vba:3-6`):

```
' TEST toggle: Plan!W1 = SI / 1 / TRUE alinea STR + equipmentByLaneByDay:
'   - EBL expand = union STR Start/End + Plan!D..L (cubre dias STR aunque Plan!D sea posterior)
'   - EBL load limit alto + filas por dia
' Vacio o NO = export legacy (STR Start = Plan!D; EBL End = Date + rangofechas a 0:00).
```

Sirve para el caso en que las fechas de carga del Plan y las de disponibilidad de equipo no
se traslapan y LBS declara todo infactible. Con `W1` activo, la ventana de equipo se calcula
como la unión de los rangos del STR y del Plan, garantizando cobertura.

También tiene su propia función derivada: `LBS_EquipTestAvailabilityEnabled` devuelve
verdadero si está activo `W1` **o** `V1` (`merged/modulo1.vba:834-836`).

## Constantes del módulo 1

| Constante | Valor | Comentario original | Cita |
|---|---|---|---|
| `LBS_ALIGN_STR_EQUIP_DATES_DEFAULT` | `False` | Valor por omisión de `W1` | `merged/modulo1.vba:7` |
| `LBS_TESTMODE_DEFAULT` | `False` | Valor por omisión de `V1` | `merged/modulo1.vba:14` |
| `LBS_TESTMODE_LOADLIMIT` | `999` | Límite de carga en modo prueba | `merged/modulo1.vba:15` |
| `LBS_WALMART_FULL_PALLET_EQUIP` | `"Z4290_WAL"` | `' LBS Walmart 28-pallet FULL_PALLET truck (EBL only). STR Required Equipment left blank like working 1v2 exports — forcing Z4290_WAL on STR collapsed Walmart fill to ~11%.` | `merged/modulo1.vba:16-18` |
| `LBS_SkipLlenadoArchivoMDLBS` | `Boolean` público | Bandera que enciende `GeneraTemplatesBuildOnly` para no escribir el MEX KA | `merged/modulo1.vba:1` |

La nota de `LBS_WALMART_FULL_PALLET_EQUIP` vale la pena leerla completa: el equipo de 28
tarimas se declara **solo** en `equipmentByLaneByDay`, y el campo `Required Equipment` del STR
se deja en blanco a propósito. Forzarlo en el STR hizo que el llenado de Walmart cayera a
alrededor del 11%.

## La validación de cartonaje corrupto

Es la única validación que **aborta** la corrida en lugar de solo advertir, y merece
explicación aparte.

`LBS_ValidarPlanCartonajeCorrupt` (`merged/modulo1.vba:917-959`) recorre el Plan buscando
renglones donde la columna `M` (cartonaje) fue interpretada por Excel como fecha. El síntoma
clásico es un valor que se muestra como `0/01/1900`. Ocurre cuando alguien pega el Plan desde
otra fuente y Excel decide que `8,400` es una fecha.

Cuando encuentra alguno, la macro:

1. Enumera hasta 12 filas afectadas con su número de fila, origen, pedido, material y el
   texto literal de `M` (`merged/modulo1.vba:942-949`).
2. Si hay más de 12, agrega `... +N mas`.
3. Muestra el mensaje
   `"Plan: cartonaje (col M) corrupto como fecha (ej. 0/01/1900). Corrija esas filas en el Plan fuente antes de continuar (N)."`
   con el título `Plan cartonaje invalido`.
4. Devuelve falso, y quien la llamó aborta.

No hay forma de saltarse esta validación, y es correcto que sea así: un cartonaje leído como
fecha produce cantidades absurdas que se propagan hasta el conteo de camiones.
