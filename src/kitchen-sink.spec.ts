import { describe, it, expect } from 'vitest';
import path from 'node:path';
import { convert } from './index';

// Fixture sintético (sin datos reales) generado por scripts/make-kitchen-sink.mjs.
// Cubre: fórmulas (simple/anidada/entre-hojas), fecha, porcentaje, booleano,
// error (#DIV/0!), celda combinada, recorte de fila/columna vacía, escapado de
// pipes y saltos de línea.
const FIXTURE = path.join(__dirname, '__fixtures__', 'kitchen-sink.xlsx');

describe('kitchen-sink (golden snapshot)', () => {
  it('default — fórmulas + coordenadas auto + valor crudo', () => {
    expect(convert(FIXTURE)).toMatchSnapshot();
  });

  it('--sin-coordenadas — primera fila como header', () => {
    expect(convert(FIXTURE, { coords: false })).toMatchSnapshot();
  });

  it('--sin-ref-celdas — sin comentarios de coordenada', () => {
    expect(convert(FIXTURE, { cellRefs: false })).toMatchSnapshot();
  });

  it('--solo-valores — sin fórmulas', () => {
    expect(convert(FIXTURE, { formulas: false })).toMatchSnapshot();
  });

  it('--formato-excel — texto formateado de Excel', () => {
    expect(convert(FIXTURE, { excelFormat: true })).toMatchSnapshot();
  });
});
