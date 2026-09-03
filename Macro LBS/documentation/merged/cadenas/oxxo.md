[Volver a cadenas](README.md) · [Macro de entrada](../README.md) · [Índice general](../../README.md)

# OXXO — macro de entrada

OXXO no tiene sufijo propio en `LBS_CadenaItemSuffix`, así que cae en el caso general
`LEFT(cadena, 3)` = `OXX`. Sus reglas viven en el equipo y en los cupos de `Handling`, no en
el item.

## Identificación

| Regla | Cómo se detecta | Cita |
|---|---|---|
| Cadena OXXO en `Handling` | `LCase(Handling!B) = "oxxo"` | `merged/modulo4.vba:156-157` |
| Destino OXXO de solo Full | Tres IDs fijos, o el nombre del CEDIS | `merged/modulo1.vba:1123-1144` |
| Destino OXXO Monterrey | `400101621` | `merged/modulo1.vba:562` |

## Destinos de solo Full

`LBS_OxxoSoloFullDestino` (`merged/modulo1.vba:1123-1144`) usa dos criterios, en este orden:

**Primero, tres IDs explícitos** (`merged/modulo1.vba:1126-1130`):

```
400083796  400079372  400083795
```

**Después, por nombre de CEDIS.** Si el ID no está en la lista, busca el destino en
`TarimaPorDestino` y compara el nombre del CEDIS (`merged/modulo1.vba:1132-1141`):

| Criterio sobre el nombre del CEDIS | Ejemplo |
|---|---|
| Contiene `TOLUCA` | `OXXO TOLUCA` |
| Contiene `AZCAPOTZALCO` | `OXXO AZCAPOTZALCO` |
| Es exactamente `SMO` | `SMO` |
| Termina en `SMO` | `OXXO SMO` |

El respaldo por nombre existe porque los IDs de los CEDIS de OXXO del metro cambian con más
frecuencia que los nombres. Si `TarimaPorDestino` no tiene el destino, la función devuelve
falso y el destino se trata como normal.

Nótese que la búsqueda usa `VLookup` sobre `TarimaPorDestino!C:B` con índice 2
(`merged/modulo1.vba:1133-1134`), lo que en Excel significa el rango de `C` hacia atrás hasta
`B`, es decir, el resultado sale de la columna `B`.

## `HandlingOXXOFullMetro`

`merged/modulo4.vba:106-143`

Es el botón que convierte esos destinos a solo Full. Por cada fila de `Handling` cuyo destino
pase `LBS_OxxoSoloFullDestino`, hace tres cosas:

**1. Consolida el cupo de encortinado dentro del de caja seca.** El comentario lo explica
(`merged/modulo4.vba:125`): *"OXXO sencillo never encortinado: fold leftover Q into caja and
clear Q/M/S"*.

```
cajaCap = Handling!P + Handling!Q
```

**2. Escribe ese cupo en cuatro columnas** (`merged/modulo4.vba:130-133`):
`P` y `L` (sencillo caja seca, activo y base) y `R` y `N` (Full caja seca, activo y base).

**3. Pone en cero los cuatro cupos de encortinado**: `Q`, `M`, `S`, `O`
(`merged/modulo4.vba:135-138`).

**4. Habilita dos plantas**: escribe `SI` en las columnas `W` y `Z`
(`merged/modulo4.vba:139-140`), que según el mapa de `LBS_HandlingPCFlagCol`
(`merged/modulo4.vba:342-355`) corresponden a **PC05** y **PC13**.

Esas dos plantas quedan habilitadas de forma incondicional para los destinos OXXO del metro.
No es un efecto secundario: es parte de la regla de negocio.

## Caja seca forzada en sencillo

`LBS_ForceOxxoSencilloCajaSeca` (`merged/modulo4.vba:146-170`)

Aplica a **todos** los destinos de OXXO, no solo a los del metro. El comentario es directo
(`merged/modulo4.vba:145`): *"OXXO: sencillo always caja seca — clear Handling encortinado
cups (Q/M)"*.

Suma el cupo de encortinado (`Q`) al de caja seca (`P`), y si el cupo base (`L`) está en
cero lo iguala también. Después pone `Q` y `M` en cero.

La consecuencia práctica es que `EqByLane` nunca puede emitir un equipo `Z1290` (sencillo
encortinado) para OXXO, porque el cupo correspondiente es cero.

Existe una función hermana, `LBS_ForceCatalogSencilloCajaSeca`
(`merged/modulo4.vba:175-213`), que hace lo mismo pero de forma general: para cualquier
destino cuyo catálogo tenga sencillo pero **no** tenga sencillo encortinado. El comentario
nombra el caso que la motivó (`merged/modulo4.vba:172-174`):

```
' Catalog Sencillo Caja Seca: fold leftover encortinado cups (Q/M) into caja (P/L).
' Prevents EquipmentByLane Z1290 when Mode Mix/Especializado has no Sencillo Encortinado
' (e.g. LA COMER City Fresko Vallejo = Caja Seca 26 only).
```

## Consolidación de OXXO Monterrey

`merged/modulo1.vba:562-563`

El destino `400101621` tiene su propia clase de consolidación:

```
Consolidation Class = Plan!F & "_" & Plan!H     ' origen + CEDIS
```

Es la cuarta prioridad de la cascada de `Consolidation Class` (ver
[../05-generacion-mexka.md](../05-generacion-mexka.md#consolidation-class-columna-ab)).
Consolidar por origen + CEDIS en lugar de por carril permite que LBS junte pedidos que llegan
al mismo CEDIS de Monterrey desde la misma planta, aunque el carril nominal difiera.

Este mismo destino tiene un cupo propio del lado de salida: 28 tarimas en sencillo, contra 24
del resto de OXXO (`LBS_OXXO_MTY_DEST` y `LBS_OXXO_MTY_SENCILLO_CAP`,
`tms_fg14/modulo2.vba:36-37`).

## Validaciones específicas

`ValidarExportMEXKA` tiene tres reglas dedicadas a OXXO
(`merged/modulo6.vba:1829-1844`):

| Mensaje | Significado |
|---|---|
| `equipmentByLaneByDay: sin filas OXXO dest=X origen=Y` (`merged/modulo6.vba:1834`) | El destino está en el Plan pero `EqByLane` no generó equipo |
| `OXXO dest=X Y: incluye sencillo Z5290` (`merged/modulo6.vba:1842`) | Un destino de solo Full tiene equipo sencillo. Falta correr `HandlingOXXOFullMetro` |
| `OXXO dest=X Y: sin Z3500/Z2550` (`merged/modulo6.vba:1844`) | Falta el equipo Full esperado |

Los mensajes OK son `OXXO dest=X Y: solo Full` (`merged/modulo6.vba:1840`) y
`OXXO dest=X: no en plan (N/A Full)` (`merged/modulo6.vba:1829`), este último cuando el
destino simplemente no tiene pedidos en esta corrida.

También hay un conteo informativo de `itemConnections`:
`itemConnections: N filas destinos OXXO metro` (`merged/modulo6.vba:2274`).

## Sufijo de item

OXXO no aparece en la tabla de `LBS_CadenaItemSuffix`, así que usa `LEFT("OXXO", 3)` = `OXX`.
En este caso el atajo funciona bien porque el nombre de la cadena es de cuatro letras sin
espacios.

`ValidarExportMEXKA` menciona `OXX` explícitamente en el mensaje de sufijo roto
(`merged/modulo6.vba:2838`) como uno de los sufijos que deben venir de `ArmadoChep`:
*"sufijo Left cadena incorrecto; usar LAC/OXX en ArmadoChep"*. Es decir, aunque el atajo
produzca el valor correcto, la práctica recomendada es que el `Item Id` venga de la columna
`E` de `ArmadoChep`.

Un hallazgo relacionado (`merged/modulo6.vba:2853`) usa OXXO como ejemplo:
*"SKU base <sku> - catalogo tiene variantes con sufijo; Plan Tarima Chep + ArmadoChep
(ej. <sku>_OXX)"*.

## Orden de ejecución

Como `ValidarDestinos` reaplica el catálogo sobre `Handling`, el orden correcto es:

1. `ValidarDestinos`
2. `HandlingOXXOFullMetro`
3. `CambioOrigen`
4. `GeneraTemplates`

Correr `HandlingOXXOFullMetro` antes de `ValidarDestinos` no sirve de nada: los cupos se
sobrescriben.

## Problemas conocidos

**Un destino del metro aparece con equipo sencillo.** Falta correr `HandlingOXXOFullMetro`,
o se corrió antes de `ValidarDestinos`.

**Un destino nuevo del metro no se reconoce.** Su ID no está en la lista de tres y su nombre
en `TarimaPorDestino!B` no contiene `TOLUCA` ni `AZCAPOTZALCO` ni termina en `SMO`. Ajustar
el nombre del CEDIS o agregar el ID a `LBS_OxxoSoloFullDestino`.

**`EqByLane` emite `Z1290` para OXXO.** No se aplicó `LBS_ForceOxxoSencilloCajaSeca`, o el
cupo `Q` se volvió a poblar desde el catálogo después.

## En la macro de salida

OXXO comparte página con NETO del lado de salida, con cupos de catálogo topados en 18 y 36,
sencillo de 24 (22 para `Z4290`, 28 para Monterrey), piso de llenado del 90% y la regla de
nunca partir un sencillo. Ver
[../../tms_fg14/cadenas/oxxo-neto.md](../../tms_fg14/cadenas/oxxo-neto.md).
