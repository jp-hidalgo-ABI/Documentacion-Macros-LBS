[Volver a cadenas](README.md) · [Macro de entrada](../README.md) · [Índice general](../../README.md)

# Chedraui — macro de entrada

La regla central de Chedraui es que **se consolida por pedido, no por carril**. Un pedido de
Chedraui trae varios SKU y todos tienen que viajar juntos.

## Identificación

`LBS_IsChedrauiCadena` (`merged/modulo1.vba:988-992`):

```vba
Public Function LBS_IsChedrauiCadena(ByVal cadena As String) As Boolean
    Dim c As String
    c = UCase$(Trim$(Replace(Replace(Replace(CStr(cadena), Chr$(160), " "), vbTab, ""), vbCrLf, "")))
    LBS_IsChedrauiCadena = (c = "CHEDRAUI")
End Function
```

Coincidencia exacta con `CHEDRAUI` en mayúsculas, tras normalizar espacios no separables,
tabuladores y saltos de línea. Un `Plan!G` con `Chedraui SA de CV` **no** dispara la regla.

## Consolidación por pedido

`merged/modulo1.vba:559-560`

Chedraui comparte la tercera prioridad de la cascada de `Consolidation Class` con los
destinos que están en la hoja `Inseparable`:

```
Consolidation Class = Plan!J    ' el pedido
```

El comentario del código da la razón (`merged/modulo1.vba:508`):

```
'   - Si cadena CHEDRAUI => Consolidation Class = Plan!J (multi-SKU por pedido; evita minEfficiency en lane)
```

Un pedido de Chedraui tiene muchos SKU distintos, cada uno con pocas tarimas. Si LBS los
consolida por carril, cada SKU individual se compara contra el umbral de eficiencia mínima
del carril, no alcanza, y el pedido completo se descarta. Consolidando por pedido, LBS ve el
volumen agregado y el pedido pasa.

Es el mismo tratamiento que reciben los destinos declarados en `Inseparable`, y la condición
en el código es literalmente `If isInsep Or isChedraui`
(`merged/modulo1.vba:551`, `559`), es decir, Chedraui está inseparable por definición sin
necesidad de estar en la hoja.

Nótese que la condición aparece **dos veces** en la cascada: una dentro del caso de Super
Ofertas partido (`merged/modulo1.vba:551`) y otra en el caso general
(`merged/modulo1.vba:559`). La primera cubre el escenario en que un renglón de Chedraui
también fuera un partido especial.

## Los seis destinos con reglas de `itemConnections`

`LBS_IsChedrauiDestinoKey` (`merged/modulo1.vba:1001-1013`)

```
400002000  400054835  400054836  400171149  400180712  101217243
```

El comentario indica qué provoca (`merged/modulo1.vba:1000`):

```
' Destinos CHEDRAUI en Plan (col I). itemConnections: Level 4 / Layer 3 / Pallet 2 (exp. restos).
```

Para estos seis destinos, las filas de `itemConnections` llevan:

| Atributo | Valor |
|---|---|
| `Level` | `4` |
| `Layer` | `3` |
| `PalletLevel` | `2` |

Esa combinación le dice a LBS que el SKU admite mezcla en capas hasta tres niveles, lo cual
es lo que permite acomodar los restos de charolas de un pedido multi-SKU. La anotación
`(exp. restos)` significa "esperando restos": estos destinos se caracterizan por generar
muchas tarimas parciales.

La función es tolerante con el formato del ID: normaliza espacios no separables y tabuladores
y convierte a número, así que funciona igual si el destino viene como texto o como valor
numérico (`merged/modulo1.vba:1003-1006`).

Que el destino `101217243` tenga un formato distinto de los demás (empieza con `101` en lugar
de `400`) es correcto: es un identificador de otra serie.

## Sufijo de item

Chedraui no está en `LBS_CadenaItemSuffix`, así que usa `LEFT("CHEDRAUI", 3)` = `CHE`. El
atajo funciona bien porque el nombre es una sola palabra.

Como siempre, la fuente preferente del `Item Id` es la columna `E` de `ArmadoChep`.

## Cómo validarlo

| Qué revisar | Dónde |
|---|---|
| Que la `Consolidation Class` sea el pedido | Columna `AB` del STR: debe traer el valor de `Plan!J`, no el carril |
| Que los seis destinos tengan `4/3/2` | Hoja `itemConnections` del export |
| Cobertura de destinos | Hallazgo `Destino Plan no esta en export STR (LBS)` (`merged/modulo6.vba:3746`) |

Fixtures del lado de salida: `tms_fg14/chedraui/*.tsv`, con los scripts
`scripts/validate_chedraui*.py`.

## Problemas conocidos

**LBS descarta pedidos completos de Chedraui por eficiencia.** El síntoma clásico de que la
`Consolidation Class` quedó como carril en lugar de pedido. Verificar que `Plan!G` sea
exactamente `CHEDRAUI`.

**Un destino nuevo de Chedraui no acomoda restos.** No está en la lista de seis de
`LBS_IsChedrauiDestinoKey`, así que sus filas de `itemConnections` no llevan `4/3/2`. Hay que
agregarlo a la función, o registrar la regla en `Condiciones Cadenas` si aplica de forma
general.

**Un pedido de Chedraui se partió entre dos camiones.** La consolidación por pedido es una
señal para LBS, no una restricción dura. Si el volumen del pedido excede el cupo del camión,
LBS lo parte de todos modos. El manejo de esa situación está del lado de salida.

## En la macro de salida

Del lado de salida Chedraui tiene cupo de 26 tarimas, piso de llenado del 80%, rescate de
pedido completo al 80%, separación por origen mixto y por sobre cupo, integridad de folio y
capas de charolas. Ver
[../../tms_fg14/cadenas/chedraui.md](../../tms_fg14/cadenas/chedraui.md).
