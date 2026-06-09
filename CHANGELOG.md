# Changelog

Todas las novedades relevantes de este proyecto se documentan en este archivo.

El formato se basa en [Keep a Changelog](https://keepachangelog.com/es-ES/1.1.0/)
y el proyecto sigue [Versionado Semántico](https://semver.org/lang/es/).

## [1.1.0] - 2026-06-09

### Added

- **Comentario de coordenada por celda** (activado por defecto): cada celda no
  vacía lleva su coordenada A1 **al inicio**, en un comentario HTML oculto
  (`<!--B2--> valor`). Los renderizadores lo ocultan; un agente lo lee en el raw
  para ubicar y documentar celdas. Se desactiva con `--sin-ref-celdas`
  (opción de librería `cellRefs`).
- **Guía "Para agentes"** en el README: recetas de `grep`/`ripgrep`/`awk` para
  ubicar una celda, construir un índice `coordenada → valor`, seguir las
  referencias de una fórmula, acotar la búsqueda a una hoja y resolver
  **referencias entre hojas** (`'Hoja'!Celda`).
- Documentación de **instalación global** (`npm i -g @skapxd/excel2md` →
  comando `excel2md`).

### Fixed

- **Entrada ESM rota**: `import { convert } from '@skapxd/excel2md'` fallaba con
  `XLSX.readFile is not a function`. El paquete `xlsx` es CommonJS y sus exports
  quedaban bajo `.default` al consumirse como ESM; ahora se resuelve la interop.
  El CLI y el `require()` (CJS) no estaban afectados.

### Internal

- Fixture sintético `kitchen-sink` + 5 snapshot tests (22 tests en total) que
  cubren fórmulas, fechas, errores, celdas combinadas, recorte de vacíos y
  escapado. Script `canary:local` para validar contra archivos reales sin
  versionarlos.

## [1.0.0] - 2026-06-09

### Added

- Conversión inicial de Excel (`.xlsx`/`.xlsm`) a Markdown **conservando las
  fórmulas**: cada celda calculada se muestra como `valor (=FORMULA)`.
- **Rejilla de coordenadas** (columnas `A,B,C…` y números de fila de Excel),
  activada automáticamente cuando la hoja tiene fórmulas.
- **Recorte** de filas y columnas totalmente vacías, conservando las coordenadas
  reales de Excel.
- Soporte **multi-hoja** (un encabezado `## NombreHoja` por hoja).
- Opciones: `--solo-valores`, `--coordenadas` / `--sin-coordenadas`,
  `--formato-excel`.
- CLI `excel2md` y API de librería `convert()` / `convertWorkbook()`.

[1.1.0]: https://github.com/skapxd/excel2md/releases/tag/v1.1.0
[1.0.0]: https://github.com/skapxd/excel2md/releases/tag/v1.0.0
