import { describe, it, expect } from 'vitest';
import * as XLSX from 'xlsx';
import { dependencySummary } from './deps';
import { convertWorkbook } from './index';

type Sheet = [string, XLSX.WorkSheet];

function wb(sheets: Sheet[]): XLSX.WorkBook {
  const b = XLSX.utils.book_new();
  for (const [name, ws] of sheets) XLSX.utils.book_append_sheet(b, ws, name);
  return b;
}

describe('dependencySummary', () => {
  it('lista dependencias por celda y preserva los rangos', () => {
    const ws: XLSX.WorkSheet = {
      '!ref': 'A1:D1',
      A1: { t: 'n', v: 1 },
      B1: { t: 'n', v: 2 },
      C1: { t: 'n', v: 3 },
      D1: { t: 'n', v: 6, f: 'SUM(A1:C1)' },
    };
    const s = dependencySummary(wb([['H', ws]]));
    expect(s).toContain('# H');
    expect(s).toContain('D1 <- A1:C1'); // rango sin expandir
  });

  it('descarta literales de texto y números (no son referencias)', () => {
    const ws: XLSX.WorkSheet = {
      '!ref': 'A1:D1',
      D1: { t: 'n', v: 0, f: 'IF(A1="",B1,C1)' },
    };
    const s = dependencySummary(wb([['H', ws]]));
    expect(s).toContain('D1 <- A1 B1 C1');
  });

  it('califica las referencias entre hojas', () => {
    const main: XLSX.WorkSheet = { '!ref': 'A1:A1', A1: { t: 'n', v: 5, f: 'Datos!B2' } };
    const datos: XLSX.WorkSheet = { '!ref': 'B2:B2', B2: { t: 'n', v: 5 } };
    const s = dependencySummary(wb([['Main', main], ['Datos', datos]]));
    expect(s).toContain('A1 <- Datos!B2');
  });

  it('cita las hojas con espacios en el nombre', () => {
    const main: XLSX.WorkSheet = { '!ref': 'A1:A1', A1: { t: 'n', v: 1, f: "'Seg de vida'!I20" } };
    const seg: XLSX.WorkSheet = { '!ref': 'I20:I20', I20: { t: 'n', v: 1 } };
    const s = dependencySummary(wb([['Main', main], ['Seg de vida', seg]]));
    expect(s).toContain("A1 <- 'Seg de vida'!I20");
  });

  it('marca las dependencias circulares reales con ⟲', () => {
    const ws: XLSX.WorkSheet = {
      '!ref': 'A1:A2',
      A1: { t: 'n', f: 'A2' },
      A2: { t: 'n', f: 'A1' },
    };
    const s = dependencySummary(wb([['C', ws]]));
    expect(s).toContain('⟲');
  });

  it('devuelve cadena vacía si no hay fórmulas', () => {
    const ws: XLSX.WorkSheet = { '!ref': 'A1:A1', A1: { t: 's', v: 'hola' } };
    expect(dependencySummary(wb([['H', ws]]))).toBe('');
  });
});

describe('convertWorkbook + resumen de dependencias', () => {
  const sample = (): XLSX.WorkBook =>
    wb([['H', { '!ref': 'A1:B1', A1: { t: 'n', v: 1 }, B1: { t: 'n', v: 2, f: 'A1+1' } }]]);

  it('antepone el resumen por defecto, antes de los datos', () => {
    const md = convertWorkbook(sample());
    expect(md).toContain('## 🔗 Dependencias (resumen)');
    expect(md.indexOf('## 🔗')).toBeLessThan(md.indexOf('## H'));
  });

  it('deps:false omite el resumen', () => {
    const md = convertWorkbook(sample(), { deps: false });
    expect(md).not.toContain('Dependencias (resumen)');
  });
});
