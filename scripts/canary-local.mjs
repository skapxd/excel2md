// Capa 2 — canary LOCAL contra archivos reales (p. ej. cotizadores internos).
// Convierte cada test-local/*.xlsx y lo compara con su baseline test-local/*.md.
// Si no existe baseline, la crea. Reporta PASS/DIFF.
//
// IMPORTANTE: test-local/ está en .gitignore. Los archivos reales y sus baselines
// NUNCA se commitean (pueden contener datos sensibles). Úsalo antes de publicar.
//
//   pnpm build && pnpm canary:local
import { readdirSync, readFileSync, writeFileSync, existsSync } from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const here = path.dirname(fileURLToPath(import.meta.url));
const root = path.join(here, '..');

const distEntry = path.join(root, 'dist', 'index.mjs');
if (!existsSync(distEntry)) {
  console.error('Falta dist/. Corre primero:  pnpm build');
  process.exit(1);
}
const { convert } = await import(distEntry);

const dir = path.join(root, 'test-local');
if (!existsSync(dir)) {
  console.error('No existe test-local/. Crea la carpeta y copia ahí tus .xlsx reales.');
  process.exit(1);
}

const files = readdirSync(dir).filter((f) => /\.(xlsx|xlsm|xls)$/i.test(f));
if (files.length === 0) {
  console.log('test-local/ no tiene archivos .xlsx. Nada que validar.');
  process.exit(0);
}

let pass = 0;
let created = 0;
let fail = 0;

for (const file of files.sort()) {
  const md = convert(path.join(dir, file));
  const baseline = path.join(dir, file.replace(/\.[^.]+$/, '.md'));

  if (!existsSync(baseline)) {
    writeFileSync(baseline, md, 'utf-8');
    console.log(`📝 baseline creada  ${file}`);
    created++;
  } else if (readFileSync(baseline, 'utf-8') === md) {
    console.log(`✅ ${file}`);
    pass++;
  } else {
    console.log(`❌ DIFF  ${file}  (la salida cambió respecto a la baseline)`);
    fail++;
  }
}

console.log(`\n${pass} ok · ${created} baseline(s) nueva(s) · ${fail} diff(s)`);
if (fail > 0) {
  console.log('Revisa el cambio. Si es intencional, borra el .md correspondiente y vuelve a correr para regenerar la baseline.');
}
process.exit(fail > 0 ? 1 : 0);
