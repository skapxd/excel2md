# @skapxd/excel2md

[![CI](https://github.com/skapxd/excel2md/actions/workflows/ci.yml/badge.svg)](https://github.com/skapxd/excel2md/actions/workflows/ci.yml)
[![npm version](https://badge.fury.io/js/@skapxd%2Fexcel2md.svg)](https://badge.fury.io/js/@skapxd%2Fexcel2md)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

**Convierte Excel (`.xlsx`/`.xlsm`) a Markdown — conservando las fórmulas.**

A diferencia de la mayoría de conversores (que aplanan cada fórmula a su valor),
`excel2md` preserva la lógica de cálculo de la hoja y la hace **rastreable**,
ideal para documentación y **contexto de IA**.

- **Fórmulas:** cada celda calculada se muestra como `valor (=FORMULA)`.
- **Coordenadas:** rejilla con columnas `A, B, C…` y números de fila de Excel, para
  que referencias como `=PMT(D24/12,I18,…)` se puedan seguir. Se activa
  automáticamente cuando la hoja tiene fórmulas.
- **Recorte:** omite filas y columnas totalmente vacías, conservando las
  coordenadas reales de Excel.
- **Multi-hoja:** un encabezado `## NombreHoja` por cada hoja del libro.

## 🤔 ¿Por qué existe este paquete?

Necesitaba pasar hojas de Excel a Markdown para dárselas como contexto a un LLM,
**sin perder las fórmulas**. Una plantilla de cálculo (un cotizador, un modelo
financiero) *es* su lógica: `=PMT(D24/12,I18,-(D21))`, `=SUM(I19:I22)`,
`=(TODAY()-I11)/365`. Si esa lógica se descarta, el Markdown resultante es un
montón de números sueltos sin el "por qué" detrás de cada uno.

Revisé las herramientas existentes y **ninguna lo resolvía**:

### Microsoft MarkItDown

[`microsoft/markitdown`](https://github.com/microsoft/markitdown) convierte el
Excel leyéndolo con **pandas** (`pd.read_excel(..., engine="openpyxl")`), que
internamente abre el libro con **`data_only=True`**. Eso significa que pandas
pide a openpyxl **el valor cacheado, no la fórmula**. Resultado:

- La fórmula `=SUM(I19:I22)` se descarta por completo y solo queda `1500`.
- Peor aún: ese valor cacheado es el que Excel guardó la última vez que
  recalculó el archivo. Si el `.xlsx` fue **generado por código** y nunca se
  abrió en Excel, **no hay valor en caché** y la celda llega **vacía**.

Para mi caso —documentar la lógica de cálculo— esto es justo lo contrario de lo
que necesito: tira lo único que me importa.

### El resto de alternativas (npm y PyPI)

El patrón se repite en todas. Todas leen el **valor**, no la fórmula:

| Herramienta | Ecosistema | ¿Conserva fórmulas? | ¿Coordenadas? |
| --- | --- | --- | --- |
| MarkItDown | Python | ❌ | ❌ |
| [`xlsx2md`](https://pypi.org/project/xlsx2md/) | Python | ❌ | ❌ |
| [`xl2md`](https://pypi.org/project/xl2md/) | Python | ❌ | ❌ |
| [`excel-to-markdown`](https://github.com/devin-liu/excel-to-markdown) | Python | ❌ | ❌ |
| [`markdown-tables`](https://github.com/cujarrett/markdown-tables) | npm | ❌ | ❌ |
| **`@skapxd/excel2md`** | **npm** | **✅** | **✅** |

Y hay un segundo problema que **ninguna** aborda: aunque conservaras las
fórmulas, estas referencian celdas por coordenada (`D24`, `I18`…). Sin una
**rejilla de coordenadas** (las letras de columna y los números de fila de
Excel), esas referencias quedan colgando: lees `=D24/12` pero no tienes forma de
saber qué celda es `D24`. Por eso `excel2md` renderiza la rejilla
automáticamente cuando hay fórmulas — para que cada referencia sea **rastreable**.

### En resumen

`@skapxd/excel2md` existe para cubrir ese hueco: es el único conversor que, por
defecto, **preserva la fórmula junto a su valor** (`1500 (=SUM(I19:I22))`) y
mantiene las **coordenadas reales de Excel** para que la lógica de la hoja siga
siendo legible y verificable — por un humano o por un modelo.

## 🚀 Uso rápido (sin instalar)

```bash
npx @skapxd/excel2md archivo.xlsx
npx @skapxd/excel2md archivo.xlsx -o salida.md
```

## 📥 Instalación global (comando `excel2md`)

Si lo vas a usar seguido, instálalo una vez de forma global y olvídate de `npx`:

```bash
npm install -g @skapxd/excel2md
# o:  pnpm add -g @skapxd/excel2md
# o:  yarn global add @skapxd/excel2md
```

Después lo invocas directo:

```bash
excel2md archivo.xlsx
excel2md archivo.xlsx -o salida.md
```

## Opciones

| Flag | Descripción |
| --- | --- |
| `-o, --output <file>` | Escribe la salida a un archivo (por defecto: stdout). |
| `--solo-valores` | Solo el valor cacheado, sin la fórmula (comportamiento clásico). |
| `--coordenadas` | Forzar la rejilla de coordenadas en todas las hojas. |
| `--sin-coordenadas` | Nunca mostrar coordenadas (la primera fila se usa como header). |
| `--formato-excel` | Usar el texto formateado de Excel (`1.00%`, `Apr-24`) en vez del valor crudo (`0.01`). |
| `--sin-ref-celdas` | Desactiva el comentario de coordenada A1 por celda (`<!--B2-->`), que viene **activado por defecto**. |

## 📦 Uso como librería

```ts
import { convert } from '@skapxd/excel2md';

const markdown = convert('archivo.xlsx', {
  formulas: true,   // `valor (=FORMULA)` (default)
  coords: null,     // null = auto (coordenadas si hay fórmulas)
  excelFormat: false, // valor crudo (default)
  cellRefs: true,   // comentario <!--B2--> por celda (default true)
});
```

### Ejemplo de salida

```markdown
## Ventas

|  | A | B | C |
| --- | --- | --- | --- |
| 1 | Producto | Precio | Total |
| 2 | Café | 10 | 30 (=B2*3) |
```

Ahora `=B2*3` es autoexplicativo: `B2 = 10`.

## 🏷️ Coordenada por celda para agentes (activado por defecto)

**Por defecto**, cada celda lleva su coordenada A1 embebida **al inicio** en un
**comentario HTML** (`<!--B2-->`): los renderizadores lo ocultan, pero un
**agente/LLM** lo lee en el raw para ubicar y documentar celdas sin tener que
cruzar la fila y la columna de la rejilla.

```markdown
## Ventas

| <!--A2--> Café | <!--B2--> 10 | <!--C2--> 30 (=B2*3) |
```

Un agente que lee `<!--C2--> 30 (=B2*3)` sabe que la fórmula está en `C2` y que
referencia `B2` — que localiza buscando `<!--B2-->`. La coordenada va **al
inicio** a propósito: funciona como ancla para extraerla con herramientas de
texto (ver más abajo).

Es **independiente** de la rejilla visible (`coords`): por defecto sale el grid
(si hay fórmulas) **y** los comentarios. Para desactivar los comentarios:

```bash
npx @skapxd/excel2md archivo.xlsx --sin-ref-celdas
```

## 🤖 Para agentes: cómo buscar en el Markdown

La coordenada al inicio de cada celda convierte el `.md` en algo **consultable
con herramientas de texto del terminal** (`grep`, `ripgrep`, `awk`), sin cargar
el archivo entero al contexto. Usa siempre `-o`/`--only-matching` para traer solo
lo que necesitas.

**Ubicar una celda por coordenada**

```bash
grep -n '<!--I19-->' archivo.md          # en qué línea/fila está I19
```

**Leer solo el valor de una celda** (mínimo contexto)

```bash
grep -oE '<!--I19-->[^|]*' archivo.md
# <!--I19--> #NUM! (=PMT(D24/12,I18,-(D21)))
```

**Índice `coordenada → valor` de toda la hoja** (de un solo comando)

```bash
grep -oE '<!--[A-Z]+[0-9]+-->[^|]*' archivo.md
# <!--A1--> Concepto
# <!--B4--> 30 (=B2*B3)
# ...
```

**Seguir las referencias de una fórmula**

La celda `<!--I19--> ... (=PMT(D24/12,I18,...))` referencia `D24` e `I18`.
Búscalas por su marcador:

```bash
grep -oE '<!--D24-->[^|]*' archivo.md
grep -oE '<!--I18-->[^|]*' archivo.md
```

**Listar solo las celdas con fórmula**

```bash
grep -oE '<!--[A-Z]+[0-9]+-->[^|]*\(=[^|]*' archivo.md
# <!--B4--> 30 (=B2*B3)
# <!--B8--> #DIV/0! (=B2/0)
```

> `ripgrep` (`rg`) acepta los mismos patrones y es más rápido en archivos
> grandes. Cada `[^|]*` se detiene en el siguiente `|`, así que captura el valor
> de **una sola** celda.

**Ojo con las hojas:** una coordenada (`I19`) es única **dentro de una hoja**,
pero se repite entre hojas. Cada hoja empieza con un encabezado `## NombreHoja`.
Para acotar la búsqueda a una hoja, usa el rango entre encabezados:

```bash
# solo la hoja "COTIZADOR-Vivienda (Interna)"
awk '/^## /{on=$0=="## COTIZADOR-Vivienda (Interna)"} on' archivo.md \
  | grep -oE '<!--I19-->[^|]*'
```

### Referencias entre hojas

Una fórmula puede referenciar una celda de **otra hoja**, con la forma
`Hoja!Celda` (o `'Hoja con espacios'!Celda`). En el raw lo verás así:

```
<!--H21--> #VALUE! (='Seguro de vida'!I20)
```

Para resolver `'Seguro de vida'!I20`:

1. **Nombre de hoja** = lo de antes del `!`. Quítale las comillas simples si las
   tiene (`'Seguro de vida'` → `Seguro de vida`); el encabezado en el Markdown va
   **sin** comillas: `## Seguro de vida`.
2. **Celda** = lo de después del `!`. Quítale los `$` de las referencias absolutas
   (`$C$20` → `C20`); un rango (`C20:D42`) abarca varias celdas.
3. Acota a esa hoja y busca la celda:

```bash
awk '/^## /{on=$0=="## Seguro de vida"} on' archivo.md | grep -oE '<!--I20-->[^|]*'
# <!--I20--> 558.6 (=H20+G20)
```

La celda destino puede tener **su propia fórmula**, así que puedes encadenar la
búsqueda para seguir el cálculo de hoja en hoja. Para listar todas las hojas:
`grep -E '^## ' archivo.md`.

## Notas sobre los valores cacheados

El valor que acompaña a la fórmula es el **último resultado que Excel guardó en
caché**. Si el archivo fue generado por código y nunca se recalculó en
Excel/LibreOffice, ese valor puede no existir: en ese caso la celda muestra
**solo la fórmula** (`=B2*3`).

## Licencia

MIT
