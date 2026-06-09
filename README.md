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

## Opciones

| Flag | Descripción |
| --- | --- |
| `-o, --output <file>` | Escribe la salida a un archivo (por defecto: stdout). |
| `--solo-valores` | Solo el valor cacheado, sin la fórmula (comportamiento clásico). |
| `--coordenadas` | Forzar la rejilla de coordenadas en todas las hojas. |
| `--sin-coordenadas` | Nunca mostrar coordenadas (la primera fila se usa como header). |
| `--formato-excel` | Usar el texto formateado de Excel (`1.00%`, `Apr-24`) en vez del valor crudo (`0.01`). |

## 📦 Uso como librería

```ts
import { convert } from '@skapxd/excel2md';

const markdown = convert('archivo.xlsx', {
  formulas: true,   // `valor (=FORMULA)` (default)
  coords: null,     // null = auto (coordenadas si hay fórmulas)
  excelFormat: false, // valor crudo (default)
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

## Notas sobre los valores cacheados

El valor que acompaña a la fórmula es el **último resultado que Excel guardó en
caché**. Si el archivo fue generado por código y nunca se recalculó en
Excel/LibreOffice, ese valor puede no existir: en ese caso la celda muestra
**solo la fórmula** (`=B2*3`).

## Licencia

MIT
