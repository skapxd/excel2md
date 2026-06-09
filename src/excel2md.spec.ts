import { describe, it, expect } from 'vitest';
import * as XLSX from 'xlsx';
import { convertWorkbook } from './index';

/** Construye un workbook de una hoja con una celda de fórmula explícita. */
function buildWorkbook(): XLSX.WorkBook {
  const ws: XLSX.WorkSheet = {
    '!ref': 'A1:C2',
    A1: { t: 's', v: 'Producto' },
    B1: { t: 's', v: 'Precio' },
    C1: { t: 's', v: 'Total' },
    A2: { t: 's', v: 'Café' },
    B2: { t: 'n', v: 10 },
    C2: { t: 'n', v: 30, f: 'B2*3' }, // valor cacheado + fórmula
  };
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Ventas');
  return wb;
}

describe('convertWorkbook', () => {
  it('conserva la fórmula junto al valor por defecto', () => {
    const md = convertWorkbook(buildWorkbook());
    expect(md).toContain('30 (=B2*3)');
  });

  it('activa coordenadas automáticamente cuando hay fórmulas', () => {
    const md = convertWorkbook(buildWorkbook());
    // Header de coordenadas + filas numeradas
    expect(md).toContain('| A | B | C |');
    expect(md).toMatch(/\|\s*2\s*\|.*Café/);
  });

  it('--solo-valores omite la fórmula', () => {
    const md = convertWorkbook(buildWorkbook(), { formulas: false });
    expect(md).toContain('30');
    expect(md).not.toContain('=B2*3');
  });

  it('--sin-coordenadas usa la primera fila como header', () => {
    const md = convertWorkbook(buildWorkbook(), { coords: false, cellRefs: false });
    expect(md).toContain('| Producto | Precio | Total |');
    expect(md).not.toContain('| A | B | C |');
  });

  it('por defecto embebe el comentario de coordenada al inicio de cada celda', () => {
    const md = convertWorkbook(buildWorkbook());
    expect(md).toContain('<!--A2-->');
    expect(md).toContain('<!--C2--> 30 (=B2*3)');
  });

  it('cellRefs:false desactiva los comentarios de coordenada', () => {
    const md = convertWorkbook(buildWorkbook(), { cellRefs: false });
    expect(md).not.toContain('<!--');
  });

  it('incluye el nombre de la hoja como encabezado', () => {
    const md = convertWorkbook(buildWorkbook());
    expect(md).toContain('## Ventas');
  });

  it('--formato-excel usa el texto formateado de la celda', () => {
    const ws: XLSX.WorkSheet = {
      '!ref': 'A1:A1',
      A1: { t: 'n', v: 0.01, w: '1.00%' },
    };
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'H');
    expect(convertWorkbook(wb, { excelFormat: true })).toContain('1.00%');
    expect(convertWorkbook(wb, { excelFormat: false })).toContain('0.01');
  });

  it('renderiza el error cacheado de una celda (#DIV/0!)', () => {
    const ws: XLSX.WorkSheet = {
      '!ref': 'A1:A1',
      A1: { t: 'e', v: 0x07, w: '#DIV/0!', f: 'A2/A3' },
    };
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'H');
    expect(convertWorkbook(wb)).toContain('#DIV/0! (=A2/A3)');
  });

  it('formatea las fechas como YYYY-MM-DD HH:mm:ss (valor crudo)', () => {
    const ws: XLSX.WorkSheet = {
      '!ref': 'A1:A1',
      A1: { t: 'n', v: 45383, z: 'm/d/yy' }, // celda con formato de fecha
    };
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'H');
    expect(convertWorkbook(wb)).toMatch(/\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}/);
  });

  it('una fórmula sin valor cacheado muestra solo la fórmula', () => {
    const ws: XLSX.WorkSheet = {
      '!ref': 'A1:A1',
      A1: { t: 'n', f: 'B1+1' }, // sin v
    };
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'H');
    expect(convertWorkbook(wb)).toContain('=B1+1');
  });

  it('marca las hojas vacías', () => {
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, {} as XLSX.WorkSheet, 'Vacia');
    expect(convertWorkbook(wb)).toContain('_(hoja vacía)_');
  });

  it('ignora nombres de hoja sin worksheet asociado', () => {
    const wb = XLSX.utils.book_new();
    wb.SheetNames.push('Ghost'); // nombre sin entrada en wb.Sheets
    expect(() => convertWorkbook(wb)).not.toThrow();
  });

  it('comentarios + --sin-coordenadas: render limpio con la coordenada al inicio', () => {
    const md = convertWorkbook(buildWorkbook(), { coords: false });
    // Comentario (al inicio) con la coordenada en una celda con contenido
    expect(md).toContain('<!--A2--> Café');
    expect(md).toContain('<!--C2--> 30 (=B2*3)');
    // Primera fila como header (sin rejilla visible)
    expect(md).toContain('| <!--A1--> Producto | <!--B1--> Precio | <!--C1--> Total |');
    expect(md).not.toMatch(/\|\s*\|\s*A\s*\|\s*B\s*\|\s*C\s*\|/);
  });

  it('--ref-celdas no anota celdas vacías (conserva el recorte)', () => {
    const ws: XLSX.WorkSheet = {
      '!ref': 'A1:B1',
      A1: { t: 's', v: 'X' },
      // B1 vacía
    };
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'H');
    const md = convertWorkbook(wb, { cellRefs: true });
    expect(md).toContain('<!--A1--> X');
    expect(md).not.toContain('<!--B1-->'); // la columna vacía se recorta, sin comentario
  });

  it('--ref-celdas + --coordenadas combina rejilla y comentarios', () => {
    const md = convertWorkbook(buildWorkbook(), { cellRefs: true, coords: true });
    expect(md).toContain('| A | B | C |'); // rejilla visible
    expect(md).toContain('<!--A2--> Café'); // y comentario al inicio
  });

  it('--formato-excel cae al valor crudo si la celda no tiene texto formateado', () => {
    const ws: XLSX.WorkSheet = {
      '!ref': 'A1:A1',
      A1: { t: 'n', v: 42 }, // sin .w
    };
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'H');
    expect(convertWorkbook(wb, { excelFormat: true })).toContain('42');
  });
});
