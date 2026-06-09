// Genera el fixture sintético `src/__fixtures__/kitchen-sink.xlsx`.
// Sin datos reales: solo estructuras que ejercitan cada capacidad de excel2md.
//   node scripts/make-kitchen-sink.mjs
import XLSX from 'xlsx';
import { mkdirSync } from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const outDir = path.join(__dirname, '..', 'src', '__fixtures__');
mkdirSync(outDir, { recursive: true });
const outPath = path.join(outDir, 'kitchen-sink.xlsx');

// --- Hoja "Cotiz": una fila por caso ---------------------------------------
// Columnas C vacía (recorte), fila 13 vacía (recorte), fila 14 merged.
const cotiz = {
  '!ref': 'A1:D14',
  '!merges': [{ s: { r: 13, c: 0 }, e: { r: 13, c: 3 } }], // A14:D14

  A1: { t: 's', v: 'Concepto' },
  B1: { t: 's', v: 'Monto' },
  D1: { t: 's', v: 'Nota' },

  A2: { t: 's', v: 'Precio' },
  B2: { t: 'n', v: 10 },

  A3: { t: 's', v: 'Cantidad' },
  B3: { t: 'n', v: 3 },

  A4: { t: 's', v: 'Total (form+cache)' },
  B4: { t: 'n', v: 30, f: 'B2*B3' }, // fórmula con valor cacheado

  A5: { t: 's', v: 'Formula anidada' },
  B5: { t: 'n', v: 13, f: 'IF(B2>5,B2+B3,0)' }, // fórmula anidada con valor
  // (el caso "fórmula sin valor cacheado" lo cubre el unit test en memoria:
  //  SheetJS no escribe celdas de fórmula sin `v`, así que no puede ir aquí)

  A6: { t: 's', v: 'Porcentaje' },
  B6: { t: 'n', v: 0.01, z: '0.00%' },

  A7: { t: 's', v: 'Fecha' },
  B7: { t: 'n', v: 45383, z: 'yyyy-mm-dd' }, // celda con formato de fecha

  A8: { t: 's', v: 'Error' },
  B8: { t: 'e', v: 0x07, w: '#DIV/0!', f: 'B2/0' }, // error cacheado + fórmula

  A9: { t: 's', v: 'Booleano' },
  B9: { t: 'b', v: true },

  A10: { t: 's', v: 'Entre hojas' },
  B10: { t: 'n', v: 99, f: 'Aux!A1' }, // referencia a otra hoja (con valor)

  A11: { t: 's', v: 'Pipe|en|texto' },
  B11: { t: 's', v: 'a|b|c' }, // escapado de pipes

  A12: { t: 's', v: 'Multilinea' },
  B12: { t: 's', v: 'linea1\nlinea2' }, // escapado de salto de línea

  // fila 13 totalmente vacía -> debe recortarse
  A14: { t: 's', v: 'PIE DE TABLA (merged)' }, // celda combinada
};

// --- Hoja "Aux": destino de la referencia entre hojas ----------------------
const aux = { '!ref': 'A1:A1', A1: { t: 'n', v: 99 } };

// --- Hoja "Vacia": sin contenido -------------------------------------------
const vacia = {};

const wb = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(wb, cotiz, 'Cotiz');
XLSX.utils.book_append_sheet(wb, aux, 'Aux');
XLSX.utils.book_append_sheet(wb, vacia, 'Vacia');

XLSX.writeFile(wb, outPath, { bookType: 'xlsx' });
console.log('Fixture escrito:', path.relative(path.join(__dirname, '..'), outPath));
