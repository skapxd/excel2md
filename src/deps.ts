import * as XLSXStar from 'xlsx';

// Interop CJS/ESM (igual que en excel2md.ts).
const XLSX = (
  'utils' in XLSXStar
    ? XLSXStar
    : (XLSXStar as unknown as { default: typeof XLSXStar }).default
) as typeof XLSXStar;

// Una referencia A1 con hoja opcional: 'Hoja'!A1, Hoja!A1, A1, A1:B3, $A$1.
// - lookbehind: no parte de un identificador mayor.
// - lookahead `(?!...\()`: descarta nombres de función (van seguidos de `(`).
const REF_RE =
  /(?<![A-Za-z0-9_'])(?:('[^']+'|[A-Za-z_][A-Za-z0-9_.]*)!)?(\$?[A-Z]{1,3}\$?[0-9]+(?::\$?[A-Z]{1,3}\$?[0-9]+)?)(?![A-Za-z0-9_(])/g;

function quoteSheet(name: string): string {
  return /^[A-Za-z0-9_]+$/.test(name) ? name : `'${name}'`;
}

/** Referencias cualificadas `Hoja!Celda` (o `Hoja!Rango`) extraídas de una fórmula. */
function extractRefs(formula: string, currentSheet: string): string[] {
  // Descarta literales de texto (no contienen referencias).
  const f = formula.replace(/"[^"]*"/g, '""');
  const out = new Set<string>();
  for (const m of f.matchAll(REF_RE)) {
    const cell = (m[2] ?? '').replace(/\$/g, ''); // quita absolutos ($A$1 -> A1)
    if (!cell) continue;
    const rawSheet = m[1];
    const sheet = rawSheet ? rawSheet.replace(/^'|'$/g, '') : currentSheet;
    out.add(`${sheet}!${cell}`);
  }
  return [...out];
}

interface Edge {
  src: string; // cualificada: Hoja!Celda
  refs: string[]; // cualificadas
}

function buildEdges(wb: XLSXStar.WorkBook): Edge[] {
  const edges: Edge[] = [];
  for (const sheet of wb.SheetNames) {
    const ws = wb.Sheets[sheet];
    const ref = ws?.['!ref'];
    if (!ws || !ref) continue;
    const range = XLSX.utils.decode_range(ref);
    for (let r = 0; r <= range.e.r; r++) {
      for (let c = 0; c <= range.e.c; c++) {
        const addr = XLSX.utils.encode_cell({ r, c });
        const cell = ws[addr] as XLSXStar.CellObject | undefined;
        if (cell && typeof cell.f === 'string' && cell.f.length > 0) {
          const src = `${sheet}!${addr}`;
          const refs = extractRefs(cell.f, sheet).filter((x) => x !== src);
          edges.push({ src, refs });
        }
      }
    }
  }
  return edges;
}

/** Celdas que forman parte de algún ciclo (DFS sobre el grafo cualificado). */
function findCyclic(edges: Edge[]): Set<string> {
  const adj = new Map<string, string[]>();
  for (const e of edges) adj.set(e.src, e.refs);

  const color = new Map<string, number>(); // 0 sin visitar, 1 en pila, 2 cerrado
  const inCycle = new Set<string>();
  const stack: string[] = [];

  const dfs = (u: string): void => {
    color.set(u, 1);
    stack.push(u);
    for (const v of adj.get(u) ?? []) {
      if (!adj.has(v)) continue; // v es una entrada (no calcula), no hay ciclo por ahí
      const cv = color.get(v) ?? 0;
      if (cv === 1) {
        const i = stack.lastIndexOf(v); // back-edge: marca v..u
        for (let k = i; k < stack.length; k++) inCycle.add(stack[k] as string);
      } else if (cv === 0) {
        dfs(v);
      }
    }
    stack.pop();
    color.set(u, 2);
  };

  for (const e of edges) if ((color.get(e.src) ?? 0) === 0) dfs(e.src);
  return inCycle;
}

/**
 * Resumen del grafo de dependencias del libro, agrupado por hoja, listo para
 * anteponer al Markdown. Devuelve '' si no hay fórmulas.
 */
export function dependencySummary(wb: XLSXStar.WorkBook): string {
  const edges = buildEdges(wb);
  if (edges.length === 0) return '';
  const cyclic = findCyclic(edges);

  const sheetOf = (q: string): string => q.slice(0, q.lastIndexOf('!'));
  const cellOf = (q: string): string => q.slice(q.lastIndexOf('!') + 1);

  const bySheet = new Map<string, Edge[]>();
  for (const e of edges) {
    const s = sheetOf(e.src);
    (bySheet.get(s) ?? bySheet.set(s, []).get(s)!).push(e);
  }

  const lines: string[] = [
    '## 🔗 Dependencias (resumen)',
    '',
    '> Mapa de qué celda depende de cuál. Cárgalo primero para navegar el resto del',
    "> archivo sin leerlo entero. Formato `Celda <- dependencias`; las de otra hoja",
    "> van como `'Hoja'!Celda`, y `⟲` marca una dependencia circular real.",
    '',
  ];

  for (const [sheet, es] of bySheet) {
    lines.push('```text');
    lines.push(`# ${sheet}`);
    for (const e of es) {
      const deps = e.refs.map((ref) => {
        const rs = sheetOf(ref);
        const rc = cellOf(ref);
        return rs === sheet ? rc : `${quoteSheet(rs)}!${rc}`;
      });
      const mark = cyclic.has(e.src) ? '  ⟲' : '';
      lines.push(`${cellOf(e.src)} <- ${deps.join(' ')}${mark}`);
    }
    lines.push('```');
    lines.push('');
  }

  return lines.join('\n').trimEnd() + '\n';
}
