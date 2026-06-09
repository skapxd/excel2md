import * as XLSXStar from 'xlsx';

// Interop CJS/ESM: el paquete `xlsx` es CommonJS. Al consumir este paquete como
// ESM (dist/index.mjs) sus exports quedan bajo `.default`; en CJS están directos.
// Resolvemos el objeto real para que XLSX.readFile/utils/SSF funcionen en ambos.
const XLSX = (
  'readFile' in XLSXStar
    ? XLSXStar
    : (XLSXStar as unknown as { default: typeof XLSXStar }).default
) as typeof XLSXStar;

/** Subconjunto tipado de XLSX.SSF (no siempre expuesto en los tipos). */
const SSF = (XLSX as unknown as {
  SSF: {
    format: (fmt: string, value: number) => string;
    is_date: (fmt: string) => boolean;
  };
}).SSF;

export interface ConvertOptions {
  /**
   * Incluir la fórmula junto al valor: `valor (=FORMULA)`.
   * @default true
   */
  formulas?: boolean;
  /**
   * Rejilla de coordenadas (columnas A,B,C… y números de fila de Excel).
   * `null`/`undefined` = automático: se activa si la hoja tiene fórmulas.
   * @default null
   */
  coords?: boolean | null;
  /**
   * Usar el texto formateado de Excel (`1.00%`, `Apr-24`) en vez del valor
   * crudo (`0.01`, `2024-04-01 00:00:00`).
   * @default false
   */
  excelFormat?: boolean;
  /**
   * Añade en cada celda no vacía un comentario HTML con su coordenada A1
   * (p. ej. ` <!--B2-->`). Los renderizadores lo ocultan; un agente puede leerlo
   * para ubicar y documentar celdas. Activado por defecto; pásalo en `false`
   * para desactivarlo. Es independiente de `coords` (la rejilla visible).
   * @default true
   */
  cellRefs?: boolean;
}

interface CellInfo {
  letter: string;
  text: string;
}

interface RowInfo {
  rownum: number;
  cells: Map<number, CellInfo>;
}

/** Valor crudo de la celda; las fechas se formatean ISO sin depender de la zona horaria. */
function rawCellValue(cell: XLSXStar.CellObject): string {
  if (cell.t === 'e') return cell.w != null ? String(cell.w) : ''; // error (#DIV/0!, …)
  const v = cell.v;
  if (v == null) return '';
  if (cell.t === 'n' && typeof cell.z === 'string' && SSF.is_date(cell.z)) {
    return SSF.format('yyyy-mm-dd hh:mm:ss', v as number);
  }
  return String(v);
}

/** Texto tal como se ve en Excel (respeta el formato de la celda). */
function formattedCellValue(cell: XLSXStar.CellObject): string {
  if (cell.w != null) return String(cell.w);
  if (cell.v != null) return String(cell.v);
  return '';
}

/** Texto Markdown de una celda, escapando pipes y saltos de línea. */
function cellToMarkdown(
  cell: XLSXStar.CellObject | undefined,
  formulas: boolean,
  excelFormat: boolean,
): string {
  if (!cell) return '';
  const value = excelFormat ? formattedCellValue(cell) : rawCellValue(cell);
  const hasFormula = typeof cell.f === 'string' && cell.f.length > 0;

  let text: string;
  if (formulas && hasFormula) {
    const formula = '=' + cell.f;
    text = value === '' ? formula : `${value} (${formula})`;
  } else {
    text = value;
  }
  return text.replace(/\|/g, '\\|').replace(/\n/g, '<br>');
}

function sheetData(
  ws: XLSXStar.WorkSheet,
  formulas: boolean,
  excelFormat: boolean,
): { rows: RowInfo[]; hasFormula: boolean } {
  const rows: RowInfo[] = [];
  let hasFormula = false;

  const ref = ws['!ref'];
  if (!ref) return { rows, hasFormula };

  const range = XLSX.utils.decode_range(ref);
  for (let r = 0; r <= range.e.r; r++) {
    const cells = new Map<number, CellInfo>();
    for (let c = 0; c <= range.e.c; c++) {
      const cell = ws[XLSX.utils.encode_cell({ r, c })] as XLSXStar.CellObject | undefined;
      if (cell && typeof cell.f === 'string' && cell.f.length > 0) hasFormula = true;
      cells.set(c, {
        letter: XLSX.utils.encode_col(c),
        text: cellToMarkdown(cell, formulas, excelFormat),
      });
    }
    rows.push({ rownum: r + 1, cells });
  }
  return { rows, hasFormula };
}

/** Renderiza las filas a una tabla Markdown, recortando filas/columnas vacías. */
function renderTable(rows: RowInfo[], coords: boolean, cellRefs: boolean): string {
  const colLetter = new Map<number, string>();
  const nonemptyCols = new Set<number>();
  for (const { cells } of rows) {
    for (const [c, info] of cells) {
      colLetter.set(c, info.letter);
      if (info.text !== '') nonemptyCols.add(c);
    }
  }

  const cols = [...nonemptyCols].sort((a, b) => a - b);
  if (cols.length === 0) return '';

  const textAt = (cells: Map<number, CellInfo>, c: number): string => cells.get(c)?.text ?? '';

  // Texto de salida: con cellRefs, antepone `<!--B2--> ` (coordenada A1) a cada
  // celda no vacía. Va al INICIO para que actúe como ancla/clave en búsquedas de
  // agente (`grep -oE '<!--B2-->[^|]*'`). El comentario HTML lo ocultan los
  // renderizadores pero un agente lo lee en el raw.
  const out = (cells: Map<number, CellInfo>, c: number, rownum: number): string => {
    const text = textAt(cells, c);
    if (cellRefs && text !== '') {
      const letter = cells.get(c)?.letter ?? colLetter.get(c) ?? '';
      return `<!--${letter}${rownum}--> ${text}`;
    }
    return text;
  };

  const kept = rows.filter(({ cells }) => cols.some((c) => textAt(cells, c) !== ''));
  if (kept.length === 0) return '';

  const lines: string[] = [];
  if (coords) {
    const header = ['', ...cols.map((c) => colLetter.get(c) ?? '')];
    lines.push('| ' + header.join(' | ') + ' |');
    lines.push('| ' + Array(cols.length + 1).fill('---').join(' | ') + ' |');
    for (const { rownum, cells } of kept) {
      const row = [String(rownum), ...cols.map((c) => out(cells, c, rownum))];
      lines.push('| ' + row.join(' | ') + ' |');
    }
  } else {
    const [headerRow, ...rest] = kept;
    if (!headerRow) return '';
    const header = cols.map((c) => out(headerRow.cells, c, headerRow.rownum));
    lines.push('| ' + header.join(' | ') + ' |');
    lines.push('| ' + Array(cols.length).fill('---').join(' | ') + ' |');
    for (const { cells, rownum } of rest) {
      lines.push('| ' + cols.map((c) => out(cells, c, rownum)).join(' | ') + ' |');
    }
  }
  return lines.join('\n');
}

/** Convierte un workbook ya cargado (SheetJS) a Markdown. */
export function convertWorkbook(wb: XLSXStar.WorkBook, options: ConvertOptions = {}): string {
  const formulas = options.formulas !== false;
  const excelFormat = options.excelFormat === true;
  const cellRefs = options.cellRefs !== false; // activado por defecto
  const coordMode = options.coords ?? null;

  const parts: string[] = [];
  for (const name of wb.SheetNames) {
    const ws = wb.Sheets[name];
    if (!ws) continue;
    const { rows, hasFormula } = sheetData(ws, formulas, excelFormat);
    // La rejilla visible es independiente de los comentarios: por defecto se
    // activa cuando la hoja tiene fórmulas (a menos que el usuario fuerce coords).
    const coords = coordMode === null ? hasFormula : coordMode;
    parts.push(`## ${name}\n`);
    const table = renderTable(rows, coords, cellRefs);
    parts.push(table ? table + '\n' : '_(hoja vacía)_\n');
  }
  return parts.join('\n').replace(/\s+$/, '') + '\n';
}

/** Lee un archivo Excel desde disco y lo convierte a Markdown. */
export function convert(filePath: string, options: ConvertOptions = {}): string {
  const wb = XLSX.readFile(filePath, { cellFormula: true, cellNF: true });
  return convertWorkbook(wb, options);
}
