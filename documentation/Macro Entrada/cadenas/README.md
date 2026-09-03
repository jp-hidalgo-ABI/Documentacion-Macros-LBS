[Volver a la macro de entrada](../README.md) · [Índice general](../../README.md)

# Reglas por cadena — macro de entrada

La macro de entrada tiene menos reglas por cadena que la de salida, pero las que tiene son
decisivas porque determinan qué puede hacer LBS. Se concentran en cuatro puntos:

1. **El sufijo del `Item Id`** que se escribe en el STR.
2. **Qué equipo se declara** en `equipmentByLaneByDay` (solo sencillo, solo Full, o ambos).
3. **La clase de consolidación** que le dice a LBS qué puede juntar.
4. **Los atributos de `itemConnections`** que definen si un SKU se puede apilar.

## Índice

| Documento | Cadenas |
|---|---|
| [walmart.md](walmart.md) | Wal Mart SC, Wal Mart BA |
| [oxxo.md](oxxo.md) | OXXO |
| [soriana-city-club.md](soriana-city-club.md) | Soriana, City Club, City Fresko |
| [chedraui.md](chedraui.md) | Chedraui |
| [la-comer.md](la-comer.md) | La Comer |
| [otras-cadenas.md](otras-cadenas.md) | Go Mart, Rabbit, Amazon, Mercado Libre, Rapiturbo, Asturiano, Conasuper, Smart and Final, Super Bara, Super Ofertas, Super Kompras, Cabrito Abarrotero, ToBe, Europea |

## Tabla maestra de sufijos de item

`LBS_CadenaItemSuffix` (`merged/modulo1.vba:963-986`). La comparación se hace sobre el texto
en mayúsculas y sin espacios no separables, tabuladores ni saltos de línea
(`merged/modulo1.vba:965`).

| Cadena (`Plan!G`) | Sufijo | Ejemplo de `Item Id` |
|---|---|---|
| `LA COMER` | `LAC` | `3003697_LAC` |
| `WAL MART SC` | `WAL` | `3003697_WAL` |
| `WAL MART BA` | `WAL_BA` | `3003697_WAL_BA` |
| `SMART AND FINAL` | `SMA_FIN` | `3003697_SMA_FIN` |
| `SUPER BARA` | `SUP` | `3003697_SUP` |
| `SUPER OFERTAS` | `SUP_OFE` | `3003697_SUP_OFE` |
| `SUPER KOMPRAS` | `SUP_KOM` | `3003697_SUP_KOM` |
| `CABRITO ABARROTERO` | `CAB_ABA` | `3003697_CAB_ABA` |
| `TOBE DISTRIBUTIONS` | `TOB_DIS` | `3003697_TOB_DIS` |
| `GO MART` | `GOMC` | `3003697_GOMC` |
| `ASTURIANO` | `ASTURIAN` | `3003697_ASTURIAN` |
| `CONASUPER` | `CONASUPE` | `3003697_CONASUPE` |
| `RAPITURBO` o `RAPPI TURBO` | `Rappi Turbo` | `3003697_Rappi Turbo` |
| `AMAZON` | `AMA` | `3003697_AMA` |
| `RABBIT` | `RAB` | `3003697_RAB` |
| `EUROPEA` | `EUROPEA` | `3003697_EUROPEA` |
| **Cualquier otra** | `LEFT(cadena, 3)` | `3003697_CHE` para Chedraui |

Dos detalles que llaman la atención en la tabla:

- **`Rappi Turbo` conserva mayúsculas y minúsculas y lleva un espacio.** No es un error de
  transcripción: así está en el código (`merged/modulo1.vba:979`) y así lo espera el maestro
  de items de LBS.
- **`SUPER BARA` usa `SUP` a secas**, mientras que `SUPER OFERTAS` y `SUPER KOMPRAS` usan
  `SUP_OFE` y `SUP_KOM`. Es decir, `SUP` no es un prefijo genérico de las cadenas "Super":
  pertenece a Super Bara.

### Por qué existe esta tabla

El comentario del código lo dice en una línea (`merged/modulo1.vba:962`):

```
' Sufijo LBS por cadena (evita Left(cadena,3) = "LA " para LA COMER; usar solo si falta fila ArmadoChep).
```

La regla original era `LEFT(cadena, 3)`, y para `LA COMER` produce `"LA "` con un espacio al
final. Eso genera items como `3003697_LA `, que LBS trata como distintos de `3003697_LAC`, y
el resultado son embarques partidos por lo que en realidad es el mismo producto.

La segunda mitad del comentario es igual de importante: esta tabla es un **respaldo**. La
fuente preferida del `Item Id` es la columna `E` (`LBS Item Id`) de la hoja `ArmadoChep`. El
sufijo solo se usa cuando esa fila no existe.

### Cuándo el sufijo no aplica

Para **tarima plástica y tarima de madera**, el `Item Id` es el SKU desnudo, sin sufijo. El
sufijo pertenece exclusivamente a tarima CHEP.

`ValidarExportMEXKA` verifica ambas direcciones: reporta items con sufijo en tarima plástica
(`merged/modulo6.vba:706`) e items desnudos cuando el catálogo tiene variantes
(`merged/modulo6.vba:2853`). Ver
[../06-validaciones-mexka.md](../06-validaciones-mexka.md#grupo-13--tarima-plástica-y-madera-con-sku-base).

## Destinos FG14

`LBS_Fg14CadenaSlug` (`merged/modulo2.vba:187-204`)

Las cadenas que no tienen un ID numérico de CEDIS se identifican por un destino sintético con
la forma `FG14_<slug>`. El slug se deriva del nombre de la cadena: se pasa a mayúsculas,
**se quitan los acentos** reemplazando `Á É Í Ó Ú Ñ` por sus equivalentes
(`merged/modulo2.vba:190-195`), y los espacios se convierten en guiones bajos.

Hay cuatro slugs explícitos (`merged/modulo2.vba:196-200`):

| Cadena | Slug | Destino resultante |
|---|---|---|
| `RAPITURBO`, `RAPPITURBO` o `RAPPI TURBO` | `RAPPITURBO` | `FG14_RAPPITURBO` |
| `AMAZON` | `AMAZON` | `FG14_AMAZON` |
| `MERCADO LIBRE` | `MERCADO_LIBRE` | `FG14_MERCADO_LIBRE` |
| `RABBIT` | `RABBIT` | `FG14_RABBIT` |

Cualquier otra cadena genera su slug por transformación automática
(`merged/modulo2.vba:201-202`).

`LBS_IsFg14DestinoKey` (`merged/modulo2.vba:206-208`) reconoce un destino FG14 simplemente
por sus primeros cuatro caracteres.

El registro de estos destinos está en la hoja `FG14 Destinos`, y
`LBS_AplicarFg14DestinosAMaestros` (`merged/modulo4.vba:1085`) siembra desde ahí las filas
que falten en `Handling` y `TarimaPorDestino`.

## Tabla comparativa de reglas

| Cadena | Sufijo | Restricción de equipo | Clase de consolidación | Regla especial |
|---|---|---|---|---|
| Wal Mart SC / BA | `WAL` / `WAL_BA` | Solo sencillo. `Z4290_WAL` únicamente en EBL | ID del STR para EXC28 | 14 SKU de excepción, split de restos |
| OXXO | `LEFT` → `OXX` en `ArmadoChep` | Tres destinos del metro son solo Full | Origen + CEDIS para Monterrey | Caja seca forzada en sencillo |
| Soriana / City Club / City Fresko | `LEFT` | **Solo sencillo**, nunca Full | Carril | Solo plantas `PC01` y `PC29` |
| Chedraui | `LEFT` → `CHE` | Catálogo | **Pedido** (`Plan!J`) | Seis destinos con `itemConnections` propios |
| La Comer | `LAC` | Catálogo | Carril + marca `R`/`C` | Separación por confirmación |
| Super Ofertas | `SUP_OFE` | Catálogo | ID del STR si la parte es de 35 | Partido en bloques de 35 tarimas |
| Rabbit | `RAB` | Catálogo | Carril | Full por SKU de latón cuando hay 35 o más tarimas |
| Go Mart | `GOMC` | Catálogo | Carril | Siempre tarima plástica |

## Funciones clasificadoras

Las funciones que reconocen cada cadena o destino, para localizarlas en el código:

| Función | Qué reconoce | Cita |
|---|---|---|
| `LBS_CadenaItemSuffix` | El sufijo del item por cadena | `merged/modulo1.vba:963` |
| `LBS_IsChedrauiCadena` | Exactamente `CHEDRAUI` | `merged/modulo1.vba:988` |
| `LBS_IsLaComerCadena` | Exactamente `LA COMER` | `merged/modulo1.vba:994` |
| `LBS_IsChedrauiDestinoKey` | Seis IDs de destino de Chedraui | `merged/modulo1.vba:1001` |
| `LBS_IsWalmartExceptionSkuBase` | 14 SKU base de la excepción EXC28 | `merged/modulo1.vba:1016` |
| `LBS_IsWalmartSplitRestosItem` | Items `*_WAL` y `*_WAL_*` de la excepción | `merged/modulo1.vba:1038` |
| `LBS_CadenaSoloSencillo` | `soriana`, `city club`, `city fresko` | `merged/modulo1.vba:1094` |
| `LBS_SorianaSoloSencillo` | Alias del anterior | `merged/modulo1.vba:1100` |
| `LBS_IsValidPlantPC` | Las nueve plantas válidas | `merged/modulo1.vba:1104` |
| `LBS_IsPlantMigrationPC` | `PC01` y `PC29` | `merged/modulo1.vba:1114` |
| `LBS_OxxoSoloFullDestino` | Destinos OXXO del metro | `merged/modulo1.vba:1123` |
| `LBS_CadenaEnforceCatalogBodyType` | Siempre verdadero | `merged/modulo1.vba:1148` |
| `LBS_Fg14CadenaSlug` | El slug FG14 por cadena | `merged/modulo2.vba:187` |
| `LBS_IsFg14DestinoKey` | Destinos que empiezan con `FG14` | `merged/modulo2.vba:206` |
| `LBS_IsLatonSkuBare` | Nueve SKU de latón de madera | `merged/modulo2.vba:211` |
| `LBS_IsLatonSkuPlanRow` | Igual, por fila del Plan | `merged/modulo2.vba:222` |
| `LBS_IsGoMartCadena` | Exactamente `GO MART` | `merged/modulo2.vba:461` |

Casi todas normalizan el texto de la misma forma antes de comparar: mayúsculas, sin espacios
no separables (`Chr$(160)`), sin tabuladores y sin saltos de línea. Eso las hace tolerantes a
los pegados desde Excel y desde correo, que es de donde vienen los nombres de cadena en la
práctica.

Nótese que la mayoría compara por **igualdad exacta**, no por coincidencia parcial. Un
`Plan!G` con `"CHEDRAUI SA"` en lugar de `"CHEDRAUI"` no dispara ninguna de las reglas de
Chedraui, y el síntoma es que el pedido se consolida por carril en lugar de por pedido.

## Plantas válidas

`LBS_IsValidPlantPC` (`merged/modulo1.vba:1104-1111`) reconoce nueve plantas:

```
PC01  PC03  PC05  PC07  PC11  PC13  PC19  PC29  PC23
```

Es el mismo conjunto y el mismo orden que las columnas `U:AC` de la hoja `Handling`
(`LBS_HandlingPCFlagCol`, `merged/modulo4.vba:342-355`). Cualquier otro valor en `Plan!F` que
no sea una de las tres marcas de fallo es un dato inválido.

`LBS_IsPlantMigrationPC` (`merged/modulo1.vba:1114-1121`) reconoce solo `PC01` y `PC29`, y su
comentario indica el contexto: *"Soriana / City Club plant migration: PC01 + PC29 only (HEB
Mexico reference)"*. Ver [soriana-city-club.md](soriana-city-club.md).
