#!/usr/bin/env node

import { program } from 'commander';
import fs from 'fs';
import path from 'path';
import { convert, type ConvertOptions } from './excel2md';

// Leer la versión desde package.json (ubicado un nivel arriba de dist/).
const pkgPath = path.join(__dirname, '../package.json');
let version = '0.0.0';
try {
  const pkg = JSON.parse(fs.readFileSync(pkgPath, 'utf-8'));
  version = pkg.version;
} catch {
  // ignore
}

program
  .name('excel2md')
  .description('Convierte un Excel (.xlsx/.xlsm) a Markdown conservando las fórmulas.')
  .version(version)
  .argument('<file>', 'Ruta al archivo .xlsx/.xlsm')
  .option('-o, --output <file>', 'Archivo de salida (por defecto: stdout)')
  .option('--solo-valores', 'Solo el valor cacheado, sin la fórmula')
  .option('--coordenadas', 'Forzar rejilla de coordenadas (A,B,C / 1,2,3) en todas las hojas')
  .option('--sin-coordenadas', 'Nunca mostrar coordenadas (primera fila como header)')
  .option('--formato-excel', 'Usar el texto formateado de Excel en vez del valor crudo')
  .action((file: string, options) => {
    let coords: boolean | null = null;
    if (options.coordenadas) coords = true;
    if (options.sinCoordenadas) coords = false;

    const opts: ConvertOptions = {
      formulas: !options.soloValores,
      coords,
      excelFormat: Boolean(options.formatoExcel),
    };

    let md: string;
    try {
      md = convert(file, opts);
    } catch (e) {
      console.error(`Error al convertir "${file}": ${(e as Error).message}`);
      process.exit(1);
    }

    if (options.output) {
      fs.writeFileSync(options.output, md, 'utf-8');
      console.error(`Escrito: ${options.output}`);
    } else {
      process.stdout.write(md);
    }
  });

program.parse(process.argv);
