[Volver al índice de la macro de salida](README.md)

# Diagnóstico y errores

Esta página es para cuando algo salió mal o algo no salió como se esperaba. Cubre las tres
señales que deja la macro (la columna `AV`, la hoja `STP Failures` y los mensajes en
pantalla), las herramientas de diagnóstico que se pueden correr por separado, y el modo
prueba.

Una distinción que conviene tener clara desde el principio: **una marca en `AV` no es un
error de la macro.** Es la macro avisando que un camión quedó en un estado que necesita ojo
humano. El camión se emite igual. Un error de verdad aborta con un `MsgBox` rojo y deja la
hoja a medias.

---

## La columna `AV` — `REVISION MANUAL`

`AV` acumula avisos, separados cuando hay más de uno. `SK_AppendReviewFlag`
(`tms_fg14/modulo2.vba:433`) agrega sin borrar lo anterior, así que una fila puede llevar
varias marcas.

### Marcas de cupo

| Marca | Qué pasó | Qué hacer |
|---|---|---|
| `REVISION MANUAL: tarimas >N (M tarimas)` | El camión quedó con más tarimas que su cupo | Revisar si el cupo del carril en `Catalogo Mode Mix` es el correcto. Si lo es, hay que partir el camión a mano |
| `REVISION MANUAL: tarimas <N (M tarimas)` | El camión quedó por debajo del piso de llenado pero por encima del mínimo para no descartarse | Decidir si se manda subutilizado o se espera más carga |
| `REVISION MANUAL: trim cupo N (M tarimas fuera)` | La macro **quitó** `M` tarimas del camión para respetar el cupo `N` | Verificar dónde quedaron esas tarimas: normalmente pasaron a `No planeado` |
| `REVISION MANUAL: trim caja N (M tarimas fuera)` | Igual, pero recortando una caja `a` o `b` de un Full | Igual |
| `REVISION MANUAL: overflow >N` | Al partir un Full no cupo todo y se abrió un camión de overflow | Revisar los camiones `R2`, `R3`, … del mismo folio |
| `REVISION MANUAL: tarimas no ubicadas en N+N` | Ni las dos cajas del Full alcanzaron para todas las tarimas | Ver [06-partir-tarimas-full.md](06-partir-tarimas-full.md) |

El rango de la marca `tarimas <N` está explicado en el comentario del código
(`tms_fg14/modulo2.vba:4431`): `' piso..cap-1 -> REVISION MANUAL.` Es decir, por debajo del
piso se descarta y a partir del cupo se recorta; la marca vive en la franja de en medio.

### Marcas de peso

| Marca | Qué pasó |
|---|---|
| `REVISION MANUAL: peso >29 ton (X ton)` | El camión superó el techo normal |
| `REVISION MANUAL: peso >52.5 ton (X ton)` | El camión superó el techo de carril Full |

El texto lo arma `LBS_PesoReviewFlagText` (`tms_fg14/modulo2.vba:11529`) con el tope que
aplique a esa fila. `LBS_StripPesoFlag` (`tms_fg14/modulo2.vba:22506`) las quita cuando una
fase posterior alivia el camión, y busca exactamente los dos textos:

```
' LBS - Quita marcas REVISION MANUAL de peso (29 t o 52.5 t) de la columna AV.
```

Si se ve una marca de peso en la salida final, es porque ninguna fase la resolvió. El camión
no se puede mandar así.

### Marcas de altura

| Marca | Qué pasó |
|---|---|
| `REVISION MANUAL: altura >X` | La suma de `AO` en la unidad supera el límite de la cadena (1.6 m en Walmart y La Comer) |

Se escribe en `tms_fg14/modulo2.vba:9392`. Casi siempre significa que el TI HI del catálogo
no corresponde al armado real, o que se apilaron dos SKU que no debían compartir tarima.

### Marcas de estructura

| Marca | Qué pasó | Qué hacer |
|---|---|---|
| `REVISION MANUAL: conf 33/20 en misma tarima AI (cap 28)` | Dos confirmaciones distintas quedaron sobre la misma unidad física en un camión Walmart | Revisar el sándwich: es el síntoma de un armado que mezcla pedidos donde no debe |
| `REVISION MANUAL: fila Full sin sufijo camion` | Una fila Full llegó a `PartirTarimasFULL` sin sufijo `a` / `b` | Ver si `SummaryOptimizar` corrió antes |
| `REVISION MANUAL: camion X=N tarimas (max M)` | Un sufijo de camión quedó sobre su máximo | Ver [06-partir-tarimas-full.md](06-partir-tarimas-full.md) |
| `REVISION MANUAL: cartonaje split` | El cartonaje se recalculó al partir y no cerró exacto | Comparar contra el Plan |
| `REVISION MANUAL: limite de filas en split` | Se agotó `PF_MAX_GROUP_EXTRA_ROWS` (50 filas por grupo) | El grupo es demasiado grande para partirse automáticamente |
| `REVISION MANUAL: overflow placement stuck` | El acomodo de overflow no avanzó y se cortó el ciclo | Caso raro; el folio necesita armado manual |

Las tres últimas vienen de `modulo5` (`modulo5.vba:1655`, `1693`, `1709`) y son defensas
contra ciclos infinitos, no reglas de negocio. Si aparecen, el folio se debe armar a mano.

### Cómo revisar `AV` en la práctica

Filtrar `Pedidos Surtidos` por `AV` distinta de vacío y agrupar por folio (`AD`). Un folio
con varias marcas del mismo tipo es un solo problema, no varios. Priorizar en este orden:

1. **Peso.** Es un impedimento legal y físico, no una preferencia.
2. **Altura.** Puede impedir cerrar la caja.
3. **Cupo por encima.** Ya sea `tarimas >N` o `trim`.
4. **Estructura.** `conf 33/20`, `overflow stuck`, `limite de filas`.
5. **Cupo por debajo.** `tarimas <N` es una decisión comercial, no un defecto.

---

## La hoja `STP Failures` como bitácora

Además de ser una entrada (lo que LBS no pudo colocar), `STP Failures` se usa como bitácora:
varias fases le agregan bloques al final, con su propio encabezado. Conviene leerla de abajo
hacia arriba.

### Bloque de `Armado` faltante

Lo escribe `SummaryOK` (`tms_fg14/modulo2.vba:1690-1697`). Cuatro columnas, sin encabezado
propio: pedido, material, cantidad y el mensaje `Armado faltante o invalido en Plan`.

Significa que el Plan no dice cómo se arma ese SKU para ese pedido, así que la macro no
puede calcular tarimas. **Se corrige en el Plan, no aquí.**

### Bloque de peso

Lo escribe `SummaryOK` (`tms_fg14/modulo2.vba:1704-1718`) con encabezado explícito:

| Columna | Encabezado |
|---|---|
| `A` | `Camion AD` |
| `B` | `Peso kg` |
| `C` | `Peso ton` |
| `D` | `Mensaje` |

El mensaje es siempre `REVISION MANUAL: peso supera 29 ton`. Es la misma información que la
columna `AV`, pero agregada por camión, que es como se revisa en operación.

### Bloque de overflow de `PartirTarimasFULL`

Lo escribe `PF_WriteOverflowLog` (`modulo5.vba:1375` y `1399`) con dos mensajes:
`REVISION MANUAL: N tarimas superan limite M` y `REVISION MANUAL: peso supera 29 ton`.
Detalle en [06-partir-tarimas-full.md](06-partir-tarimas-full.md).

### Filas de `CompararCartonajes`

`CompararCartonajes` (`tms_fg14/modulo3.vba:886`) no escribe en `STP Failures`: agrega filas
a `Pedidos Surtidos` con estado `No planeado` y el motivo `No reportado por LBS`. Es una
señal distinta y más grave, porque significa que hay cartonaje del Plan que LBS ni colocó ni
rechazó. Ver [07-fallos-y-remonte.md](07-fallos-y-remonte.md).

---

## Los mensajes en pantalla

### Mensajes de fin normal

| Mensaje | De dónde | Significado |
|---|---|---|
| `Proceso Finalizado` | `modulo1.vba:268` | `ProcesoMacro` importó todo |
| `SummaryOptimizar completado.` + `Filtro por cadena, consolidacion y cupo aplicados.` | `modulo2.vba:2557` | Fin normal |
| `Proceso completado correctamente.` + tiempo | `modulo3.vba:850` | Fin normal de `SummaryFallo` |
| `ConsolidaDiag generado: N folios Programado.` | `modulo2.vba:22499` | Fin del diagnóstico |
| `Hoja 'Consolida' creada/actualizada con N destinatarios.` | `modulo2.vba:5434` | Fin de la siembra |
| `Hoja 'TI HI' actualizada (sin filtrar).` | `modulo2.vba:8309` | Fin de la sincronización |
| `El archivo CSV se ha generado en la ruta: …` | `modulo6.vba:196` | Fin de la exportación a SAP |

### Advertencias (`vbExclamation`)

Terminó, pero hay que revisar algo.

| Mensaje | De dónde | Qué revisar |
|---|---|---|
| `Proceso completado con N linea(s) sin Armado en Plan. Ver STP Failures.` | `modulo2.vba:1699` y `modulo3.vba:858` | El bloque de `Armado` faltante en `STP Failures`, y el Plan |
| `Proceso completado con N camion(es) sobre tope de peso. Ver columna AV y STP Failures.` | `modulo2.vba:1719` | El bloque de peso |
| `Partir Fulles: <detalle> Ver columna AV y STP Failures.` | `modulo5.vba:1905` | Las marcas de `modulo5` en `AV` |

### Errores (`vbCritical`)

Abortó. La hoja queda a medio construir y **hay que volver a empezar desde el paso que
falló**, no continuar con el siguiente botón.

| Mensaje | De dónde |
|---|---|
| `SummaryOK failed at phase [<fase>]` + `Err N: <descripción>` | `modulo2.vba:1741` |
| `SummaryOptimizar failed at phase [<fase>]` + `Plan!X: <progreso>` + `Err N: <descripción>` | `modulo2.vba:2575` |
| `Sync TI HI fallo: <detalle>` | `modulo2.vba:8333` |
| `Error al importar Plan (<paso>): …` | `modulo1.vba:254` |
| `Error: No se pudo obtener la ruta del archivo.` | `modulo6.vba:169` |
| `Error N: <descripción>` (exportación CSV) | `modulo6.vba:200` |

El nombre de la fase entre corchetes es la pieza más útil del mensaje. Ver la sección
siguiente.

### El aviso de re-entrada

```
Ya hay una macro LBS en ejecucion. Espera a que termine (mira Plan!X / barra de estado).
```

Aparece en `modulo2.vba:602`, `modulo2.vba:2515` y `modulo3.vba:648`. Es la bandera
`SK_MacroBusy` haciendo su trabajo: **no es un error.** Significa que se hizo clic en un
botón mientras otra macro corría. Esperar y mirar `Plan!X1`.

---

## Las fases y el indicador de progreso

Los tres módulos principales publican su fase actual en `Plan!X1` con el formato
`<fase> | hh:mm:ss`, y `SummaryOK` / `SummaryOptimizar` la incluyen en el mensaje de error.
Saber en qué fase murió reduce el problema a unas pocas líneas de código.

`SK_Phase` (`tms_fg14/modulo2.vba:4`) recorre esta secuencia. Las fases de `SummaryOK`:

| Fase | Línea | Qué está haciendo |
|---|---|---|
| `init` | `625` | Preparando |
| `pca_copy` | `643` | Copiando `Pallet Container Association` |
| `pca_prep` | `865` | Construyendo la caché de alturas de PCA |
| `agrupar` | `898` | Agrupando `Shipments` |
| `load_stp` / `dict_stp` | `1084` / `1096` | Cargando `STP-Equipment Assoc` |
| `load_plan` | `1102` | Cargando el Plan |
| `dict_equip` | `1108` | Cargando `Equipments` |
| `consolidar` | `1119` | El ciclo principal de armado |
| `post_proceso` | `1407` | Ajustes posteriores |
| `ordenar` | `1419` | Ordenando `Pedidos Surtidos` |
| `split_mixed_vigencia` | `1449` | Separando vigencias mezcladas |
| `assign_dedup_ai` | `1454` | Asignando y deduplicando unidades `AI` |
| `promote_restos_fulls` | `1459` | Promoviendo restos a Full |
| `alturas_ao_ap` | `1463` | Calculando `AO` y `AP` |
| `repack_restos` | `1468` | Reacomodando restos |
| `alturas_post_repack` | `1474` | Recalculando alturas |
| `merge_thin` | `1480` | Uniendo tarimas delgadas |
| `alturas_post_merge_thin` | `1487` | Recalculando alturas |
| `coalesce_clubcity_sku` | `1492` | Uniendo restos del mismo SKU en Soriana / City Club |
| `refresh_at_tihi` | `1507` | Refrescando el peso `AT` desde TI HI |
| `promote_multi_single_pre_weight` | `1512` | Promoviendo antes de pesar |
| `discard_over_weight` | `1515` | Descartando lo que excede el peso |
| `totales_camion` | `1529` | Totales por camión |
| `totales:assign_ai` … `totales:trim_oxxo_cap` | `1555`-`1575` | Subfases de totales |

Las de `SummaryOptimizar`:

| Fase | Línea | Qué está haciendo |
|---|---|---|
| `consolidar_singles` | `2533` | Uniendo sencillos con parciales |
| `consolidar_comextra_sku` | `2537` | Consolidación por SKU de Comextra |
| `dedup_sandwich_w` | `2542` | Deduplicando la columna `W` de sándwich |
| `filtrar_eficiencia` | `2546` | Entra a `FiltrarPorEficiencia` |
| `filtrar:split_lacomer` / `split_chedraui` / `split_vigencia` | `4858`-`4873` | Divisiones por cadena y por vigencia |
| `filtrar:merge_metro` | `4880` | Uniendo por metro |
| `filtrar:clubcity_cross_vig` | `4891` | Segunda pasada cruzando vigencias |
| `filtrar:recompute_totals` | `4902` | Recalculando totales |
| `filtrar:lookup_ar` | `4908` | Buscando el umbral de `EFICIENCIA POR CADENA` |
| `filtrar:chedraui_gate` / `pack_chedraui` / `enforce_chedraui` | `4933`-`4941` | Gate, empaque y piso de Chedraui |
| `filtrar:lacomer_gate` / `pack_lacomer` | `4944` / `4947` | Gate y empaque de La Comer |
| `filtrar:walmart_gate` / `alsuper_gate` / `pack_walmart` | `4963`-`4972` | Gates y empaque de Walmart y Alsuper |
| `filtrar:min_fill` | `4982` | Aplicando los pisos de llenado |
| `filtrar:totales_finales` | `4987` | Totales finales |
| `filtrar:descartar_cadena` | `4993` | Descartes por cadena |
| `filtrar:mover_no_planeado` | `5023` | Moviendo a `No planeado` |
| `filtrar:consolidar_no_planeado` | `5043` | Intentando remontar los `No planeado` |
| `filtrar:pack_cap35` | `5056` | Empaque de mayoristas a 35 |
| `filtrar:z_final` | `5063` | Escritura final de la columna `Z` |

Interpretación rápida: si el error cae en una fase `load_*` o `dict_*`, el problema está en
los datos de entrada. Si cae en `pca_*`, es tamaño (ver `LBS_WALMART_PCA_MAX_ROWS`). Si cae
en una fase `filtrar:*` con nombre de cadena, el problema es de esa cadena y la página
correspondiente de [cadenas/](cadenas/README.md) es el siguiente lugar donde buscar.

`SummaryFallo` usa su propio indicador con textos en español y sin variable de fase
(`SF_SetProgress`, `tms_fg14/modulo3.vba:7`): `Fallos: restore multi-order SINGLE S`,
`Fallos: comparar cartonajes`, `Fallos: consolidar restos`, `Fallos: listo`.

---

## Herramientas de diagnóstico

Las tres se ejecutan a mano, desde la lista de macros o desde los botones de `User Guide`.
Ninguna modifica `Pedidos Surtidos`.

### Diagnosticar la consolidación

`LBS_DiagnosticarConsolidacion` (`tms_fg14/modulo2.vba:22403`) responde a la pregunta más
frecuente de operación: *"¿por qué estos dos folios no viajaron en el mismo camión?"*

Recorre las filas `Programado` de `Pedidos Surtidos`, agrupa por folio (`AD`), calcula la
llave de consolidación de cada folio y escribe la hoja `ConsolidaDiag`:

| Columna | Encabezado | Contenido |
|---|---|---|
| `A` | `Clave (plant|grupo)` | La llave de consolidación del folio |
| `B` | `Folio (AD)` | El folio |
| `C` | `Tarimas folio` | Tarimas del folio: Fulls (`T`) más charolas distintas (`AI`) |
| `D` | `Folios en la clave` | Cuántos folios comparten esa llave |
| `E` | `Tarimas en la clave` | Suma de tarimas de todos esos folios |
| `F` | `Deberia unirse?` | `SI (cabe en 1 camion, cap N)` o `Parcial (excede cap N)` |

El conteo de tarimas de la columna `C` merece explicación: suma las tarimas completas de `T`
y cuenta las charolas distintas por `AI`, tratando una fila sin `AI` como su propia charola
(`modulo2.vba:22438-22442`). Es la misma cuenta física que usa el motor.

Un folio sin llave se muestra como `(sin consolidacion) <cadena>`, tomando la cadena de la
columna `L` (`modulo2.vba:22426`).

El mensaje de cierre dice cómo leer el resultado (`modulo2.vba:22499-22502`):

```
ConsolidaDiag generado: N folios Programado.
Revisa la col F: 'SI' = folios que deberian estar en un solo camion.
Si dos folios que esperas juntos tienen claves distintas en la col A, el grupo no coincide (hoja Consolida).
```

Los dos diagnósticos posibles:

- **Claves distintas en la columna `A`.** Los folios no se ven entre sí. El problema está en
  la hoja `Consolida`: uno de los destinatarios no está mapeado, o están en grupos
  diferentes. Es el caso más común.
- **Misma clave y `Parcial (excede cap N)`.** Los folios sí se ven, pero juntos no caben.
  El problema es de cupo, no de mapeo, y la macro hizo lo correcto.

Un `SI` en la columna `F` con los folios en camiones distintos en la salida final es lo único
que constituye un defecto real del motor.

Ojo: llama a `LBS_ResetConsMap` al empezar (`modulo2.vba:22407`), así que relee la hoja
`Consolida`. Si se acaba de editar esa hoja, el diagnóstico ya refleja el cambio aunque el
armado anterior no.

### Sembrar la hoja `Consolida`

`LBS_SeedConsolidaSheet` (`tms_fg14/modulo2.vba:5407`) crea la hoja `Consolida` desde la
tabla embebida en el código, para que el planeador la mantenga sin tocar VBA. El comentario
es explícito sobre cuándo usarla (`modulo2.vba:5405-5406`):

```
' LBS - Crea (o reescribe) la hoja "Consolida" con la tabla por defecto, para que el
' planeador la mantenga sin tocar codigo. Ejecutar una vez; despues editar la hoja.
```

**Es destructiva.** Hace `wsC.Cells.Clear` (`modulo2.vba:5421`) antes de escribir, así que
borra cualquier mapeo que el planeador haya agregado. Se corre una vez, al montar el libro,
y después nunca más.

Formatea la columna `A` como texto (`modulo2.vba:5422`), que es necesario: los destinatarios
son números largos como `400087097` y Excel los convertiría a notación científica o les
quitaría ceros. Si se agregan destinatarios a mano, hay que respetar ese formato.

Al terminar llama a `LBS_ResetConsMap` para que la siguiente corrida lea la hoja recién
escrita.

### Sincronizar la hoja `TI HI`

`LBS_SyncTihiSheet` (`tms_fg14/modulo2.vba:8239`) trae el catálogo homologado al libro. Es un
paso previo, no parte del flujo: el comentario dice que las búsquedas de `SummaryOK` leen
**solo** la hoja y nunca abren archivos a mitad de la macro (`modulo2.vba:8237-8238`).

Busca el archivo en tres lugares, en orden (`modulo2.vba:8259-8275`):

1. La URL de SharePoint de `LBS_WALMART_TIHI_URL`.
2. Una copia local, resuelta por `LBS_WalmartTihiFindLocalPath` (normalmente en Descargas).
3. Un diálogo `Seleccione TI HI VALIDADO ACOMODOS.xlsx (catalogo homologado)`.

Copia el catálogo **sin filtrar** hasta la columna `Q` como mínimo, preservando la fila 1
oculta del archivo original para que la distribución coincida
(`LBS_WalmartTihiCopyCatalogToSheet`, `modulo2.vba:8190`). Después limpia la caché y la
reconstruye, y reporta dos números:

```
Hoja 'TI HI' actualizada (sin filtrar).

Filas de datos copiadas: N
Claves Walmart BA/SC en cache AO: M

SummaryOK usara esta hoja para alturas TI HI (no reabre archivos).
```

Si el segundo número es muy bajo (menos de `LBS_WALMART_TIHI_MIN_KEYS`, que son 50), el
catálogo no cargó bien: probablemente se abrió el archivo equivocado o la hoja no tiene
`Cadena` en la columna `F`.

Los tres errores propios que puede lanzar, con su significado:

| Mensaje | Causa |
|---|---|
| `No se pudo abrir el catalogo TI HI (SharePoint, Downloads, ni archivo elegido).` | Ninguna de las tres rutas funcionó. Descargar el archivo a mano y volver a intentar |
| `La fuente resolvio a este mismo libro. Elija TI HI VALIDADO ACOMODOS.xlsx.` | Se seleccionó el propio libro de la macro |
| `El archivo abierto no tiene hoja de catalogo (encabezado Cadena en col F).` | El archivo no es el catálogo homologado, o le cambiaron la estructura |

Sobre SharePoint hay una salvaguarda que conviene conocer
(`LBS_WalmartTihiTryOpen`, `modulo2.vba:8119`). Durante `SummaryOK` la macro **se niega** a
abrir rutas `http://`, `https://` o que contengan `sharepoint.com`; el comentario dice por
qué: `' SummaryOK / lookups: never Workbooks.Open http(s)/SharePoint (hard-closes Excel).`
Solo el botón de sincronización tiene permiso (`allowRemote = True`), y ahí se dejan los
avisos de Excel activos a propósito para que el inicio de sesión de SharePoint pueda
preguntar credenciales (`modulo2.vba:8155`).

### Limpiar las cachés

`LBS_ResetConsMap` (`tms_fg14/modulo2.vba:5256`) limpia de una vez tres cachés: el mapa
`Consolida`, la lista `Cadenas 35 Tarimas` y los umbrales de `EFICIENCIA POR CADENA`. Se
llama automáticamente al inicio del camionizado y de la exportación, pero se puede correr a
mano después de editar cualquiera de esas tres hojas.

No limpia la caché de `TI HI` ni la de `Catalogo Mode Mix`. Para la primera, `LBS_SyncTihiSheet`
la reconstruye; para la segunda, hay que cerrar y reabrir el libro.

---

## El modo prueba (`Plan!V1`)

`Plan!V1` habilita un solo comportamiento: permitir que `SummaryFallo` remonte filas cuyo
motivo de fallo en `AG` sea `Shippable qty is 0`. El comentario
(`tms_fg14/modulo2.vba:6047-6048`) delimita el alcance:

```
' TEST MODE: Plan!V1 = SI / 1 / TRUE -> allow remount of "Shippable qty is 0" AG rows.
' Vacio o NO = comportamiento productivo (nunca remonta esos motivos).
```

Valores aceptados por `LBS_IgnoreShippableQtyZeroEnabled` (`tms_fg14/modulo2.vba:6049`):

| Resultado | Valores |
|---|---|
| Activado | `SI`, `S`, `1`, `TRUE`, `ON`, `YES`, `Y`, o el booleano `VERDADERO` |
| Desactivado | `NO`, `N`, `0`, `FALSE`, `OFF` |
| Por omisión (desactivado) | Vacío o cualquier otro texto |

Cuando está activo, `SummaryFallo` lo dice en su mensaje de cierre
(`tms_fg14/modulo3.vba:853-856`):

```
TEST MODE (Plan!V1): se ignoro 'Shippable qty is 0' al remountar.
```

**Para qué sirve.** `Shippable qty is 0` significa que LBS no tenía inventario disponible
para esa línea. En producción no se debe remontar: no hay producto que cargar. En pruebas sí
conviene, porque permite ejercitar el motor de armado con datos donde el inventario no
importa, y ver cómo quedarían los camiones si el producto existiera.

**Por qué hay que apagarlo.** Un camión armado en modo prueba puede llevar tarimas que
físicamente no existen. Si ese archivo llega a TMS, se programa un viaje para carga
inexistente. La única defensa es el mensaje de cierre, así que conviene leerlo siempre.

Recomendación: dejar `Plan!V1` vacío y ponerle `SI` solo durante la prueba, borrándolo
inmediatamente después. No usar `NO`, porque una celda vacía es más visible que un `NO`
cuando alguien revisa el libro.

---

## Guía rápida por síntoma

| Síntoma | Dónde mirar primero |
|---|---|
| Un botón parece no responder | `Plan!X1`. Y **no** volver a hacer clic (ver `SK_MacroBusy` en el [README](README.md#protección-contra-doble-clic)) |
| `SummaryOK` aborta en `pca_copy` o `pca_prep` | Tamaño de `Pallet Container Association`: `LBS_WALMART_PCA_MAX_ROWS` son 250 000 filas |
| Error de importación con "reporta N filas (max 50000)" | Celdas sueltas al final de las columnas clave del archivo de entrada |
| Dos folios que deberían ir juntos van en camiones distintos | `LBS_DiagnosticarConsolidacion`, columna `A` de `ConsolidaDiag` |
| Camiones muy vacíos de una cadena | Piso de llenado en [09-parametros-y-catalogos.md](09-parametros-y-catalogos.md#pisos-de-llenado), y el umbral de `EFICIENCIA POR CADENA` |
| Alturas en cero o marcas de altura | La hoja `TI HI`: correr `LBS_SyncTihiSheet` y revisar el conteo de claves |
| Un SKU de mayorista no llega a 35 tarimas | La hoja `Cadenas 35 Tarimas` (columna `D` = `YES`) **y** la lista blanca del código |
| Filas `No reportado por LBS` | [07-fallos-y-remonte.md](07-fallos-y-remonte.md): hay cartonaje del Plan que LBS ignoró |
| Todo se ve bien pero el archivo de TMS sale corto | El `Exit For` de `TMS`: ver [08-salidas-tms-y-sap.md](08-salidas-tms-y-sap.md) |
| Cambios en una hoja de parámetros que no surten efecto | Cachés: `LBS_ResetConsMap`, o cerrar y reabrir el libro |
