[Volver a la macro de entrada](README.md) · [Índice general](../README.md)

# Sincronización de maestros

`merged/modulo7.vba`

Cuando `ValidarExportMEXKA` reporta que faltan items, destinos, carriles o equipos en el
archivo MEX KA, la sincronización los completa automáticamente. Trabaja en **las dos
direcciones**: agrega al libro de planeación lo que está en el export y le falta al libro, y
agrega al export lo que está en el Plan y le falta al export.

El encabezado del módulo lo describe así (`merged/modulo7.vba:1-2`):

```
' Modulo7 - MEX KA sync incremental rebuild (see data/MODULO7_INCREMENTAL.md)
' Step macros: dry-run -> plan-only -> export pipelines. Full sync after gates pass.
```

Ese es el patrón de uso recomendado: primero la simulación, después los pasos que solo tocan
el libro de planeación, y solo al final los que escriben en el export.

## Arquitectura: una máquina con nueve interruptores

Todas las macros públicas hacen lo mismo: llaman a `LBS_Sync_SetConfig` con una combinación
de banderas y después a `LBS_Sync_Execute` (`merged/modulo7.vba:145-183`). No hay lógica
distinta por paso, solo configuración distinta.

Las nueve banderas de `LBS_Sync_SetConfig` (`merged/modulo7.vba:185-200`):

| Bandera | Qué habilita |
|---|---|
| `dryRun` | Simulación: no guarda ni el libro ni el export |
| `writePlan` | Permite escribir en el libro de planeación |
| `writeExport` | Permite escribir en el archivo MEX KA |
| `pipeItems` | Ejecuta el pipeline de items |
| `pipeDestHandling` | Ejecuta el pipeline de destinos sobre `Handling` y `TarimaPorDestino` |
| `pipeDestSiteBod` | Ejecuta el pipeline de destinos sobre `SiteMaster_ABPP` y `BodDetail_ABPP` |
| `pipeDestConnections` | Ejecuta el pipeline de `connections` en el export |
| `pipeEquip` | Ejecuta el pipeline de equipos |
| `connectionsBodDetailOnly` | Restringe la fuente de carriles a `BodDetail_ABPP` |

## Los pasos

| Macro | Etiqueta | Dry run | Escribe Plan | Escribe export | Pipelines activos | Cita |
|---|---|---|---|---|---|---|
| `Sync_Step0_DryRun` | `STEP0_DRYRUN` | Sí | No | No | Items, Handling, SiteBod, Connections | `merged/modulo7.vba:145-148` |
| `Sync_Step1_PlanDestinosSetup` | `STEP1_PLAN_DESTINOS` | No | Sí | No | Handling y `TarimaPorDestino` | `merged/modulo7.vba:150-153` |
| `Sync_Step2_PlanItemsSetup` | `STEP2_PLAN_ITEMS` | No | Sí | No | Items | `merged/modulo7.vba:155-158` |
| `Sync_Step3_PlanBodDetailSetup` | `STEP3_PLAN_BODDETAIL` | No | Sí | No | SiteMaster y BodDetail | `merged/modulo7.vba:160-163` |
| `Sync_Step4_ExportItems` | `STEP4_EXPORT_ITEMS` | No | No | **Sí** | Items | `merged/modulo7.vba:165-168` |
| `Sync_Step5_ExportConnections` | `STEP5_EXPORT_CONNECTIONS` | No | No | **Sí** | Connections | `merged/modulo7.vba:170-173` |
| `Sync_Step6_ExportEquipments` | `STEP6_EXPORT_EQUIPMENTS` | No | No | **Sí** | Equipment | `merged/modulo7.vba:175-178` |
| `SincronizarMaestrosBidireccional` | `FULL` | No | Sí | **Sí** | Todos | `merged/modulo7.vba:180-183` |

La división en dos mitades es intencional: los pasos 1 a 3 solo tocan el libro de planeación,
que es reversible con un cierre sin guardar. Los pasos 4 a 6 escriben en el MEX KA, que es el
archivo que se va a importar.

## `LBS_Sync_Execute`: la secuencia común

`merged/modulo7.vba:202-306`

### Guardas de entrada

1. **Resuelve el archivo MEX KA.** Si no lo encuentra ni se selecciona en el diálogo, sale
   con `No se selecciono export MEX KA (.xlsx).` (`merged/modulo7.vba:225`).
2. **Verifica que existan todas las hojas requeridas** en el libro de planeación con
   `LBS_Sync_RequirePlanSheets` (`merged/modulo7.vba:229`). Si falta alguna, muestra
   `Falta la hoja requerida en el libro Plan: <nombre>` y sale (`merged/modulo7.vba:231`).

### Apertura del export

El export se abre en solo lectura salvo que el paso escriba en él y no sea dry run
(`merged/modulo7.vba:243-245`):

```
exportReadOnly = mSyncDryRun Or Not mSyncWriteExport
```

Esto significa que en dry run y en los pasos 1 a 3 el archivo MEX KA **no se puede modificar
ni por accidente**, porque Excel lo tiene abierto en solo lectura.

### Los cuatro pipelines

Se ejecutan en orden fijo, cada uno si su bandera está activa
(`merged/modulo7.vba:251-270`):

| Orden | Pipeline | Función | Qué hace |
|---|---|---|---|
| 1 | Items | `LBS_Sync_PipelineItems` | Agrega los items que faltan a `items` e `itemPackage` del export, y a las hojas de armado del libro |
| 2 | Destinos (Handling/Tarima) | `LBS_Sync_PipelineDestinosHandlingTarima` | Agrega las filas que faltan en `Handling` y `TarimaPorDestino` |
| 3 | Destinos (SiteMaster/BodDetail) | `LBS_Sync_PipelineDestinosSiteBod` | Agrega las filas que faltan en `SiteMaster_ABPP` y `BodDetail_ABPP` |
| 4 | Connections | `LBS_Sync_PipelineConnectionsExport` | Agrega los carriles que faltan en la hoja `connections` del export |
| 5 | Equipment | `LBS_Sync_PipelineEquipment` | Agrega los equipos que faltan en el export |

El orden importa: los destinos tienen que existir en `BodDetail_ABPP` antes de que el
pipeline de `connections` pueda derivar los carriles.

### Cierre

1. **Guarda el export** con `LBS_Sync_SaveExportWorkbook`, salvo en dry run
   (`merged/modulo7.vba:272-277`). En dry run registra `DRY-RUN: export no guardado`.
2. **Cierra el export sin guardar** (ya se guardó explícitamente si tocaba)
   (`merged/modulo7.vba:280-283`).
3. **Escribe el reporte** en `data\mex_ka_bidirectional_sync_report.txt`
   (`merged/modulo7.vba:286-292`).
4. **Corre `ValidarExportMEXKA` automáticamente** (`merged/modulo7.vba:294-295`). Esto es
   importante: al terminar cualquier sincronización aparecen **dos** cuadros de diálogo, el
   del validador primero y el de la sincronización después.
5. Muestra el resumen.

### El resumen

`merged/modulo7.vba:297-305`

```
Sync <ETIQUETA> completado.

Items cambios (aprox.): N
Destinos/lanes cambios (aprox.): N
Equipments agregados: N

Reporte: C:\...\data\mex_ka_bidirectional_sync_report.txt

Siguiente: registrar en data\SYNC_REBUILD_LOG.md y correr gate (Validar + Llenado + Solve).
```

En dry run agrega la línea `Modo: DRY-RUN (sin guardar)`.

La última línea es una instrucción de proceso, no del código: cada sincronización debe
quedar registrada en la bitácora `data\SYNC_REBUILD_LOG.md`, y después hay que correr el
ciclo completo de verificación (validar, llenar y resolver en LBS) para confirmar que la
sincronización no rompió nada.

## El reporte

`LBS_Sync_WriteReport` escribe `data\mex_ka_bidirectional_sync_report.txt` con el registro
línea por línea de lo que se hizo. El tope es de 200 entradas
(`LBS_SYNC_MAX_LOG`, `merged/modulo7.vba:6`); a partir de ahí el log se trunca, aunque los
conteos del resumen siguen siendo correctos.

El registro siempre abre con estas líneas (`merged/modulo7.vba:238-241`):

```
Step: <ETIQUETA>
Mode: DRY-RUN (no guarda plan ni export)     <- solo en dry run
Export: <nombre del archivo MEX KA>
Plan: <nombre del libro de planeación>
```

## Cómo usarlo

### Flujo conservador (recomendado la primera vez)

1. `Sync_Step0_DryRun` — leer el reporte completo y entender qué se va a cambiar.
2. `Sync_Step1_PlanDestinosSetup` — completar `Handling` y `TarimaPorDestino`.
3. `Sync_Step2_PlanItemsSetup` — completar items y armados.
4. `Sync_Step3_PlanBodDetailSetup` — completar `BodDetail`.
5. Revisar el libro de planeación. Si algo se ve mal, cerrar sin guardar.
6. `Sync_Step4_ExportItems`, `Sync_Step5_ExportConnections`, `Sync_Step6_ExportEquipments`.
7. Registrar en `data\SYNC_REBUILD_LOG.md`.
8. Correr `GeneraTemplates` y `ValidarExportMEXKA` para confirmar.

### Flujo rápido

`SincronizarMaestrosBidireccional` hace los siete pasos de una vez y valida al final. Es lo
apropiado cuando ya se conoce el comportamiento y los hallazgos son del tipo rutinario
("faltan tres items nuevos del mes").

## Manejo de errores

`merged/modulo7.vba:308+`

La variable `syncStep` va registrando la etapa en curso, con estos valores posibles:
`Inicio`, `Abrir export MEX KA`, `Pipeline items`,
`Pipeline destinos (Handling/Tarima)`, `Pipeline destinos (SiteMaster/BodDetail)`,
`Pipeline connections (export)`, `Pipeline equipment`, `Guardar export MEX KA`,
`Cerrar export MEX KA`, `Ruta reporte sync`, `Escribir reporte sync`,
`Validar export post-sync`.

Cuando algo falla, el mensaje de error incluye esa etapa, el número y la descripción del
error de VBA. Saber en qué pipeline se quedó indica directamente qué maestro tiene el
problema.

## La versión legacy

`merged/modulo7_legacy.vba:9` contiene una implementación anterior de
`SincronizarMaestrosBidireccional`, monolítica: hacía todo de una vez sin la separación por
pasos ni el modo dry run.

Se conserva como respaldo pero **no debe usarse**. Los dos módulos declaran un procedimiento
público con el mismo nombre, así que no pueden estar ambos importados en el mismo libro de
VBA: Excel rechazaría la compilación por nombre ambiguo. Si el libro tiene el módulo legacy
importado en lugar del nuevo, no existen las macros `Sync_Step*`, y esa es la forma más
rápida de detectarlo.

## Relación con las validaciones

Cada paso resuelve un grupo concreto de hallazgos de
[`ValidarExportMEXKA`](06-validaciones-mexka.md):

| Hallazgo típico | Paso que lo resuelve |
|---|---|
| `LBS Invalid Item (pre-import): ... no esta en hoja items del export` | `Sync_Step4_ExportItems` |
| `LBS Missing Item Package (pre-import): ... esta en items pero no en itemPackage` | `Sync_Step4_ExportItems` |
| `items: faltan N NETO en hoja items` | `Sync_Step4_ExportItems` |
| `connections: dest X en STR sin lane (Invalid Connection Id)` | `Sync_Step5_ExportConnections` |
| `HEB 400170996: itemConnections OK pero connections vacio` | `Sync_Step5_ExportConnections` |
| Fila 15 de `User Guide`: destinos sin `Handling` | `Sync_Step1_PlanDestinosSetup` |
| Fila 16 de `User Guide`: items sin `ArmadoMadera` | `Sync_Step2_PlanItemsSetup` |

Lo que la sincronización **no** resuelve son los hallazgos de datos: `NoInventario`,
`FaltaInventario`, `FABRICA`, fechas incoherentes o cartonaje corrupto. Esos requieren
corregir el Plan o los inventarios.

## Constantes del módulo

| Constante | Valor | Rol | Cita |
|---|---|---|---|
| `LBS_SYNC_MAX_LOG` | `200` | Máximo de entradas en el log de sincronización | `merged/modulo7.vba:6` |
| `LBS_CAT_MODEMIX_DEFAULT_URL` | URL de OneDrive | Origen del `Catalogo Mode Mix` | `merged/modulo7.vba:16-17` |
| `LBS_CAT_MODEMIX_SHEET_NAME` | `"Catalogo Mode Mix"` | Nombre de la hoja en ese archivo | `merged/modulo7.vba:18` |

Las constantes del catálogo homologado están documentadas en
[08-catalogo-homologado-tihi.md](08-catalogo-homologado-tihi.md).
