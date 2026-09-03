[Volver a cadenas](README.md) · [Macro de entrada](../README.md) · [Índice general](../../README.md)

# Soriana, City Club y City Fresko — macro de entrada

Las tres comparten una sola regla, pero es de las más restrictivas del sistema: **nunca
viajan en Full**.

## Identificación

`LBS_CadenaSoloSencillo` (`merged/modulo1.vba:1094-1098`):

```vba
Public Function LBS_CadenaSoloSencillo(ByVal cadena As String) As Boolean
    Dim c As String
    c = LCase$(Trim$(cadena))
    LBS_CadenaSoloSencillo = (c = "soriana" Or c = "city club" Or c = "city fresko")
End Function
```

Tres literales exactos, comparados en minúsculas y sin espacios en los extremos:

| Valor de `Plan!G` | Reconocido |
|---|---|
| `soriana` / `Soriana` / `SORIANA` | Sí |
| `city club` / `City Club` | Sí |
| `city fresko` / `City Fresko` | Sí |
| `Soriana Mercado` | **No** |
| `City Market` | **No** |

Existe un alias, `LBS_SorianaSoloSencillo` (`merged/modulo1.vba:1100-1102`), que simplemente
llama a la función anterior. Es un remanente de cuando la regla era solo de Soriana y se
extendió a City Club y City Fresko.

**City Fresko es la inclusión que más sorprende.** Es un formato pequeño que uno esperaría
que fuera flexible, pero la regla lo trata igual que Soriana.

## La regla de solo sencillo

`EqByLane` no emite filas de Full para estas tres cadenas. El comentario de
`LBS_CadenaEnforceCatalogBodyType` lo confirma (`merged/modulo1.vba:1146-1147`):
*"Walmart / solo-sencillo still skip Full emit"*.

Esto significa que, **aunque el `Catalogo Mode Mix` diga `F` para un carril de Soriana**, el
export no va a declarar equipo Full para ese carril. La regla de cadena gana sobre el
catálogo.

Es la única excepción documentada al principio general de que el catálogo manda, junto con la
de Walmart.

## Plantas válidas: solo `PC01` y `PC29`

`LBS_IsPlantMigrationPC` (`merged/modulo1.vba:1113-1121`)

```vba
' Soriana / City Club plant migration: PC01 + PC29 only (HEB Mexico reference).
Public Function LBS_IsPlantMigrationPC(ByVal origen As String) As Boolean
    Select Case UCase$(Trim$(origen))
        Case "PC01", "PC29"
            LBS_IsPlantMigrationPC = True
        Case Else
            LBS_IsPlantMigrationPC = False
    End Select
End Function
```

Estas cadenas están en una migración de plantas y, durante ella, solo pueden surtirse desde
`PC01` y `PC29`. La referencia entre paréntesis apunta al caso análogo de HEB México.

### La validación en `User Guide` fila 21

`LBS_ValidarPlantOrigenCadenas` (`merged/modulo4.vba:1216-1255`) recorre el Plan buscando
renglones de estas tres cadenas con un origen que no sea `PC01` ni `PC29`, y los reporta en
la fila 21 de `User Guide` con el formato:

```
<origen>_<destino> (origen invalido; usar PC01 o PC29)
```

La lógica de filtrado tiene una particularidad (`merged/modulo4.vba:1235-1242`). Solo se
reportan los renglones que:

1. Pasan `LBS_PlanRowExportable`.
2. Son de una cadena de solo sencillo.
3. **No** están en `NoInventario`, `FaltaInventario` ni `FABRICA`.
4. **No** están en `PC01` ni `PC29`.
5. Y además, el origen **contiene `VALDER` o empieza con `333`**
   (`merged/modulo4.vba:1242`).

Ese quinto filtro es el que hace que la validación no reporte cualquier planta distinta, sino
solo los identificadores de la nomenclatura vieja: los que contienen `VALDER` (Valle de
Derramadero) y los códigos numéricos que empiezan con `333`. Son precisamente los que la
migración debe eliminar.

Un renglón de Soriana con origen `PC13`, por ejemplo, **no** aparecería en esta validación
aunque también sea un origen no autorizado por la regla de migración.

Los hallazgos se deduplican por carril con un diccionario (`merged/modulo4.vba:1245-1249`) y
al final se compactan con `LBS_CompactUserGuideRowCatalog` (`merged/modulo4.vba:1254`), que
recorre las columnas `G` a `AZ` de la fila 21 eliminando duplicados y huecos
(`merged/modulo4.vba:1257-1281`).

## Sufijo de item

Ninguna de las tres está en `LBS_CadenaItemSuffix`, así que usan `LEFT(cadena, 3)`:

| Cadena | Sufijo resultante |
|---|---|
| `Soriana` | `Sor` |
| `City Club` | `Cit` |
| `City Fresko` | `Cit` |

**City Club y City Fresko generan el mismo sufijo `Cit`.** Si un mismo SKU se pide para
ambas cadenas, los dos renglones producen el mismo `Item Id` y LBS los va a consolidar como
si fueran el mismo producto para el mismo cliente.

En la práctica esto no ha causado problemas porque el `Item Id` preferente viene de la
columna `E` de `ArmadoChep`, donde cada cadena tiene su fila propia. Pero si falta esa fila,
la colisión es real. Es un argumento fuerte para mantener `ArmadoChep` completo para estas
dos cadenas.

## Consolidación

No hay regla especial. Estas cadenas caen en el caso por omisión de `Consolidation Class`:
el carril (`merged/modulo1.vba:574`).

## Cómo validarlo

| Qué revisar | Dónde |
|---|---|
| Orígenes inválidos de la migración | Fila 21 de `User Guide`, y el hallazgo correspondiente en el reporte de `ValidarExportMEXKA` |
| Que no haya equipo Full en estos carriles | Buscar `Z3500` o `Z2550` en `equipmentByLaneByDay` para destinos de estas cadenas |
| Colisión de sufijo entre City Club y City Fresko | Buscar items `*_Cit` duplicados en la hoja `items` del export |

Fixtures del lado de salida: `tms_fg14/soriana/*.tsv`, con el script
`scripts/validate_soriana_weight_tihi.py`.

## Problemas conocidos

**Un renglón de Soriana sale con equipo Full.** Verificar que `Plan!G` sea exactamente
`Soriana` y no una variante. `Soriana Mercado` o `SORIANA HIPER` no disparan la regla.

**La fila 21 de `User Guide` está vacía pero hay orígenes raros.** Recordar el quinto filtro:
solo se reportan los orígenes con `VALDER` o que empiezan con `333`. Otros orígenes no
autorizados pasan sin aviso.

**Un renglón de City Fresko se consolidó con uno de City Club.** Revisar que ambos tengan su
fila en `ArmadoChep` con el `LBS Item Id` correcto en la columna `E`.

## En la macro de salida

Del lado de salida las tres forman la familia `CLUBCITY`, con cupo de metro de 26 tarimas,
piso de llenado del 70%, segunda pasada de cross-vigencia y fusión de Fulls sin sufijo. Ver
[../../tms_fg14/cadenas/soriana-city-club.md](../../tms_fg14/cadenas/soriana-city-club.md).
