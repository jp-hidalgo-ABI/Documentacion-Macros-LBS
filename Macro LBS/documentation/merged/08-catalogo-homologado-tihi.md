[Volver a la macro de entrada](README.md) · [Índice general](../README.md)

# Catálogo homologado y TI HI

El catálogo homologado es el documento maestro de acomodos: define, para cada combinación de
cadena y material, cuántas cajas caben en una tarima (`Armado`), cómo se acomodan (`TI` cajas
por cama, `HI` camas por tarima) y las dimensiones y peso resultantes.

Lo publica el equipo de acomodos en SharePoint. La macro lo consume con
`SincronizarCatalogoHomologado`, que actualiza `ArmadoChep`, `ArmadoMadera`, `Plan!P` y el
`itemPackage` del MEX KA.

## Origen de los datos

`merged/modulo7.vba:13-14`

```
https://anheuserbuschinbev.sharepoint.com/sites/MAZ3/POLOGMX/Operación Canal Moderno/
CEDIS/2026/001 _Catalogos/TI HI VALIDADO ACOMODOS.xlsx
```

Se abre por URL usando la sesión de ABI del usuario, así que funciona incluso con el libro de
planeación abierto desde SharePoint. `LBS_CatHom_ResolveCatalogSource`
(`merged/modulo7.vba:1582`) resuelve la ruta, con la URL como valor por omisión.

### Mapa de columnas del catálogo

`merged/modulo7.vba:19-27`

| Constante | Columna | Contenido |
|---|---|---|
| `CAT_HOM_COL_CADENA` | `F` | Cadena |
| `CAT_HOM_COL_MATERIAL` | `G` | Material / SKU |
| `CAT_HOM_COL_ARMADO` | `I` | Armado (cajas por tarima) |
| `CAT_HOM_COL_TI` | `J` | TI: cajas por cama |
| `CAT_HOM_COL_HI` | `K` | HI: camas por tarima |
| `CAT_HOM_COL_LARGO` | `M` | Largo, en centímetros |
| `CAT_HOM_COL_ANCHO` | `N` | Ancho, en centímetros |
| `CAT_HOM_COL_ALTO` | `O` | Alto, en centímetros |
| `CAT_HOM_COL_ALT_ARM` | `Q` | Altura del armado, en centímetros |

El mismo libro de SharePoint tiene una segunda hoja de excepciones de armado CHEP
(`merged/modulo7.vba:29-32`):

| Constante | Columna | Contenido |
|---|---|---|
| `CAT_EXC_COL_CADENA` | `E` | Cadena |
| `CAT_EXC_COL_MATERIAL` | `F` | Material |
| `CAT_EXC_COL_ARMADO` | `H` | Armado CHEP |
| `CAT_EXC_COL_ARMADO_ALT` | `K` | Armado CHEP alterno (respaldo si `H` está vacía) |

Aunque las constantes fijan estas posiciones, la macro localiza la hoja y la fila de
encabezado dinámicamente (`LBS_CatHom_FindCatalogSheet`, `merged/modulo7.vba:1607`) y guarda
los índices reales de columna en variables de módulo. Esto tolera que el catálogo agregue o
mueva columnas, siempre que los encabezados sigan reconociéndose.

## La guarda de carga parcial

`merged/modulo7.vba:10-12`

```
' Full TI HI Catalogo Homologado is ~2000+ data rows; SharePoint open often materializes only ~500.
Private Const LBS_CAT_HOM_MIN_EXPECTED_ROWS As Long = 1000
Private Const LBS_CAT_HOM_MIN_EXPECTED_KEYS As Long = 800
```

Este es un detalle importante y fácil de pasar por alto. Al abrir un archivo grande desde
SharePoint, Excel a veces materializa solo las primeras ~500 filas. Si la macro sincronizara
con ese catálogo truncado, **borraría** los armados de los items que quedaron fuera.

`LBS_CatHom_EnsureFullCatalog` (`merged/modulo7.vba:1613`) verifica que se hayan leído al
menos 1 000 filas y 800 claves distintas antes de continuar, y si no, reintenta la carga.

Si la sincronización termina con conteos de catálogo mucho más bajos de lo normal
(`Catalogo: Chep=... Madera=...` en el resumen), es señal de carga parcial y hay que volver a
correrla.

## `SincronizarCatalogoHomologado`

`merged/modulo7.vba:1510`

### El diálogo de confirmación

Antes de tocar nada, muestra esta advertencia (`merged/modulo7.vba:1511-1520`):

```
Se actualizaran armados, empaques y Catalogo Mode Mix:
- ArmadoChep: Tarima Chep + Ultrapallet
- ArmadoMadera: Tarima Plastica / Madera
- Plan col P
- MEX KA itemPackage: pallet qty, TI, alturas (m), pesos (kg), correccion cm->m
- Catalogo Mode Mix: SharePoint gana en overlaps; se conservan lanes solo del Plan

Plan en SharePoint / OneDrive web: soportado (catalogos abren por URL con sesion ABI).
Nota: el libro TMS de salida NO se modifica (Pedidos Surtidos!V = Plan!P via SummaryOK).

Recomendado: respaldo del .xlsm y del export MEX KA antes de continuar.

Continuar?
```

El botón por omisión es **No** (`vbDefaultButton2`), a propósito: esta macro modifica seis
destinos distintos y no tiene deshacer. La recomendación de respaldar el `.xlsm` y el MEX KA
es en serio.

Para ver qué haría sin escribir nada, existe `SincronizarCatalogoHomologado_DryRun`
(`merged/modulo7.vba:1505-1508`), que activa la bandera `mCatHomDryRun` y ejecuta la misma
lógica.

### Qué actualiza

| Destino | Regla |
|---|---|
| `ArmadoChep` | Items con tipo de tarima CHEP o Ultrapallet |
| `ArmadoMadera` | Items con tarima plástica o de madera |
| `Plan!P` | El armado de cada renglón, según su cadena, material y tipo de tarima |
| MEX KA `itemPackage` | Cantidad por tarima, TI, alturas en metros, pesos en kilogramos, y la corrección de centímetros a metros |
| `Catalogo Mode Mix` | Se mezcla con la versión de SharePoint: los carriles que están en ambos se sobrescriben con SharePoint, y los que solo están en el Plan se conservan |

El libro TMS de salida **no** se toca. La razón está en el propio diálogo: la columna `V` de
`Pedidos Surtidos` hereda `Plan!P` a través de `SummaryOK`, así que basta con volver a correr
la macro de salida después de sincronizar.

### Normalización de tipos de tarima

`LBS_CatHom_NormalizeTarima` (`merged/modulo7.vba:105-127`) es más tolerante de lo que
aparenta. Además de pasar a mayúsculas y quitar espacios no separables, tabuladores y saltos
de línea, **quita los acentos** de forma explícita reemplazando los caracteres Unicode
`Á É Í Ó Ú Ñ` por sus equivalentes sin acento (`merged/modulo7.vba:113-118`).

Esto resuelve el problema que `ValidarExportMEXKA` reporta como
*"Tarima Plastica: revisar acento en macro Item Id"*: en el Plan conviven
`Tarima Plástica` y `Tarima Plastica` según quién capturó el renglón.

Además, cinco variantes de Ultrapallet se colapsan a `TARIMA CHEP`
(`merged/modulo7.vba:119-123`): `TARIMA ULTRA PALLET`, `ULTRAPALLET`, `ULTRA PALLET`,
`TARIMA ULTRA` y `TARIMA ULTRAPALLET`. Es decir, Ultrapallet se trata como CHEP para efectos
de armado.

## La corrección de centímetros a metros

`LBS_CatHom_FixPkgCmRow` (`merged/modulo7.vba:5037-5087`)

El catálogo trae las dimensiones en centímetros y LBS las espera en metros. La conversión es
trivial —dividir entre 100 (`merged/modulo7.vba:4088`, `4173`, `4185`)— pero el problema real
es que el `itemPackage` del export puede tener una mezcla: unas filas ya en metros y otras
todavía en centímetros, según cuándo se cargaron.

La macro decide caso por caso con `LBS_CatHom_PkgLooksCm`, usando umbrales por nivel de
empaque (`merged/modulo7.vba:5025-5035`):

| Nivel (`cases` / `layers` / `pallets`) | Umbral que dispara la conversión |
|---|---|
| `cases` / `cajas` | Alto mayor a 1; largo o ancho mayores a 2 |
| `layers` / `capas` | Alto mayor a 1 |
| `pallets` / `plataformas` | Alto mayor a 3 |

La lógica es geométrica: una caja de más de 1 metro de alto no existe, así que un `40` en esa
columna son 40 centímetros capturados sin convertir, no 40 metros. El umbral de 3 metros para
tarimas es más laxo porque una tarima armada sí puede acercarse a los 2.5 metros.

Cada corrección se registra en el log con el valor antes y después
(`merged/modulo7.vba:5077`):

```
<etiqueta> <item> <nivel> H cm->m 180.000 -> 1.800
```

Y se cuenta en `mCatHomNPkgCmFixMexKa`, que aparece en el resumen como
`MEX KA itemPackage cm->m fix: N`.

## Alturas placeholder y tarimas altas

### Placeholders

`merged/modulo7.vba:101-102`, `4346-4347`

```
Private Const LBS_CAT_HOM_PLACEHOLDER_H1 As Double = 0.1
Private Const LBS_CAT_HOM_PLACEHOLDER_H2 As Double = 0.125
```

Los valores 0.1 m y 0.125 m son alturas de relleno que quedaron en el template del export
cuando no se conocía la altura real. `LBS_CatHom_IsPlaceholderPkgHeight`
(`merged/modulo7.vba:4344-4348`) las reconoce con una tolerancia de 0.0005 m y permite
sobrescribirlas sin preguntar, porque no representan un dato real que se pueda perder.

### Tarimas altas

`merged/modulo7.vba:100`, `4572-4592`

```
Private Const LBS_CAT_HOM_TALL_PALLET_M As Double = 3#
```

Después de sincronizar, la macro revisa si alguna tarima quedó con más de 3 metros de altura
y lo reporta como advertencia con ejemplos (`merged/modulo7.vba:4590`). No lo corrige: 3
metros es físicamente imposible, así que indica un dato malo en el catálogo fuente que hay
que reportar al equipo de acomodos.

El aviso aparece en el resumen como
`MEX KA pallets Height > 3m: N` y, si hubo casos, con el detalle
`WARN pallets Height > 3m; ej: ...`.

Nótese que el tope de la validación del export es 2.5 m
(`LBS_VALEXP_MAX_HEIGHT_M`, `merged/modulo6.vba:11`), más estricto que este de 3 m. Una
tarima de 2.7 m pasa la sincronización sin advertencia pero falla en
`ValidarExportMEXKA`.

## El resumen de la sincronización

`merged/modulo7.vba:2046-2072`

```
Sincronizacion completada          <- o "VISTA PREVIA (sin guardar)" en dry run

Catalogo: Chep=N Madera=N
ArmadoChep (Chep+Ultrapallet): N
ArmadoMadera (Plastica/Madera): N
Plan col P: N
Mode Mix overwrite/keep/append: N/N/N
Mode Mix source: <origen>
MEX KA itemPackage armado: N
MEX KA itemPackage TI/HI: N
MEX KA itemPackage cm->m fix: N
MEX KA homolog suffix: N
MEX KA armado catalogo (forzado): N
MEX KA TI/HI catalogo (forzado): N
MEX KA pallets Height > 3m: N
TMS: omitido (CatHom NO actualiza el libro TMS de salida)
  Pedidos Surtidos!V hereda Plan!P via SummaryOK — re-correr TMS tras sync.

Reporte: C:\...\data\catalogo_homologado_sync_report.txt

Siguiente:
  0) AuditarArmadoVsCatalogo — validar Plan AD/P vs TI HI
  1) ValidarDestinos — aplica Catalogo Mode Mix a Handling
  2) CambioOrigen (armado-chep-guard) si tip/armado inconsistente
  3) TMS SummaryOK — copia Plan!P a Pedidos Surtidos!V
  4) SummaryOptimizar / PartirFulles / SummaryFallo
No re-ejecutar Optimizar sobre hoja ya optimizada.
```

La lista de "Siguiente" es la parte accionable: después de sincronizar hay que rehacer el
pipeline completo, porque los armados cambiaron y con ellos el conteo de tarimas de todos los
renglones. La advertencia final —no volver a correr `SummaryOptimizar` sobre una hoja ya
optimizada— aplica a la macro de salida y se explica en
[../tms_fg14/01-runbook-operador.md](../tms_fg14/01-runbook-operador.md).

El tope del log es de 500 entradas (`LBS_CAT_HOM_MAX_LOG`, `merged/modulo7.vba:9`).

## `AuditarArmadoVsCatalogo`

`merged/modulo7.vba:1528`

La versión de solo lectura. El comentario del código la describe como
*"Client VBA audit: Plan AD/P + ArmadoChep vs TI HI (no Python)"*
(`merged/modulo7.vba:1527`): reemplaza en VBA lo que antes se hacía con scripts externos, de
modo que el cliente pueda auditar sin instalar nada.

Compara y **no escribe nada**. Produce cinco conteos
(`merged/modulo7.vba:1616`):

| Conteo | Significado |
|---|---|
| `nOk` | Renglones donde el armado del Plan coincide con TI HI |
| `nMadTipChepOnly` | Renglones de tarima de madera cuyo catálogo solo tiene armado CHEP |
| `nChepMismatch` | Renglones CHEP con armado distinto al catálogo |
| `nMadMismatch` | Renglones de madera con armado distinto al catálogo |
| `nNoTihi` | Renglones sin entrada en el catálogo TI HI |

Además de recorrer el Plan (`LBS_AuditArmado_ScanPlan`), audita la hoja `ArmadoChep` completa
contra el catálogo (`LBS_AuditArmado_ScanArmadoChepSheet`, `merged/modulo7.vba:1617`), lo que
detecta filas obsoletas que ya no coinciden.

El reporte queda en `data\armado_vs_tihi_audit_report.txt`
(`merged/modulo7.vba:1619-1621`).

### Guardas propias

`AuditarArmadoVsCatalogo` verifica dos cosas que también aplican a la sincronización:

- Que el catálogo no haya resuelto al propio libro de planeación
  (`merged/modulo7.vba:1603-1606`), con el mensaje
  `"El catalogo resolvio al libro Plan; use TI HI VALIDADO ACOMODOS.xlsx"`.
- Que la hoja tenga las columnas de cadena, material y armado
  (`merged/modulo7.vba:1608-1611`), con
  `"No se encontro hoja Cadena/Material/Armado en el catalogo."`.

Si el libro está abierto desde SharePoint, registra la advertencia
`WARN Plan abierto desde SharePoint/web; reportes en TEMP\LBS_reports`
(`merged/modulo7.vba:1597`) y escribe el reporte en la carpeta temporal.

## Cuándo correr cada cosa

| Situación | Qué correr |
|---|---|
| El equipo de acomodos publicó una versión nueva del catálogo | `AuditarArmadoVsCatalogo` primero, para ver el tamaño del cambio |
| La auditoría muestra muchos desajustes | `SincronizarCatalogoHomologado_DryRun`, revisar el reporte, y después `SincronizarCatalogoHomologado` |
| `ValidarExportMEXKA` reporta alturas mayores a 2.5 m | `SincronizarCatalogoHomologado`, que aplica la corrección de cm a m |
| Falta el armado de un item concreto (fila 18 de `User Guide`) | `SincronizarCatalogoHomologado`, o agregar la fila a mano en `ArmadoChep` |
| Nada de lo anterior | No correr nada. No es parte del ciclo diario |

Después de cualquier sincronización, **rehacer el pipeline completo** desde
`ValidarDestinos`.

## Relación con la hoja `TI HI` de la macro de salida

La macro de salida tiene su propia hoja `TI HI` dentro del libro, que alimenta el cálculo de
alturas y de camas compartidas. Esa hoja **no** la actualiza esta sincronización: se
actualiza con `LBS_SyncTihiSheet`, un botón separado del libro de salida que lee la misma URL
de SharePoint (`tms_fg14/modulo2.vba:92-94`).

Es decir, el catálogo TI HI tiene dos consumidores independientes que hay que actualizar por
separado:

| Consumidor | Cómo se actualiza | Qué alimenta |
|---|---|---|
| Libro de planeación | `SincronizarCatalogoHomologado` | `ArmadoChep`, `ArmadoMadera`, `Plan!P`, `itemPackage` del MEX KA |
| Libro de salida | `LBS_SyncTihiSheet` | La hoja `TI HI` que usa el motor de armado para alturas y camas |

Ver [../tms_fg14/10-diagnostico-y-errores.md](../tms_fg14/10-diagnostico-y-errores.md).
