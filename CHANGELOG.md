# Changelog

Todas las novedades relevantes de este proyecto se documentan en este archivo.

El formato se basa en [Keep a Changelog](https://keepachangelog.com/es-ES/1.1.0/)
y el proyecto sigue [Versionado Semántico](https://semver.org/lang/es/).

## [1.4.0] - 2026-06-09

### Added

- La celda que es **destino de un named range** lleva ahora su nombre en un
  segundo comentario HTML (`<!--C4--> <!--Rate-->`), así que el nombre se busca
  igual que la coordenada (`grep '<!--Rate-->'`). El `<!--C4-->` se mantiene
  intacto, no rompe búsquedas existentes. Para nombres que apuntan a un rango se
  usa la celda ancla.
- API: `buildCellNameMap(workbook)`.

## [1.3.0] - 2026-06-09

### Added

- El resumen de dependencias ahora **resuelve los named ranges** (rangos con
  nombre) a su celda real usando la tabla de nombres del libro: una fórmula como
  `=PMT(Rate/Pmts_per_year,Npmts,-Loan_amt)` pasa de no mostrar dependencias a
  `H3 <- C4 C5 C6 C3`. Hojas enteras basadas en nombres dejan de salir vacías.

### Changed

- Se **omiten** del resumen las celdas cuya fórmula no referencia ninguna celda
  (p. ej. `=TODAY()`, constantes): no aportan navegación y eran ruido.

## [1.2.1] - 2026-06-09

### Documentation

- Subsección **"Cómo leer el grafo"**: explica la flecha `<-` ("depende de") y el
  concepto de **entradas** (celdas que nadie calcula).
- Destacado de la intro actualizado: el **mapa de dependencias ya viene en el
  output**, no es solo una guía manual.
- Auditoría del README: la opción `deps` en el ejemplo de la API, columna
  **"¿Grafo de deps?"** en la tabla comparativa, y varios ejemplos sincronizados
  con el resumen de dependencias (incl. `--sin-deps`).

## [1.2.0] - 2026-06-09

### Added

- **Resumen del grafo de dependencias** al inicio del Markdown, **activado por
  defecto** (`--sin-deps` / opción `deps` para omitirlo). Agrupa por hoja un mapa
  `Celda <- dependencias` que:
  - **califica** las referencias entre hojas (`'Hoja'!Celda`),
  - **preserva los rangos** (`A1:B3`) sin expandirlos,
  - **detecta dependencias circulares reales** (`⟲`) con DFS — y al calificar por
    hoja elimina los falsos ciclos que produce el análisis con shell.

  Pensado para que un agente cargue solo ese mapa y navegue el resto del archivo
  sin leerlo entero.
- API pública: `dependencySummary(workbook)` y opción `ConvertOptions.deps`.
- Guía para agentes ampliada: cómo extraer el resumen (`sed`) y profundizar.

## [1.1.1] - 2026-06-09

### Documentation

- Sección **"Un vistazo: de Excel a Markdown"** con un ejemplo simple (sumas por
  fila y por columna) y la salida real de la herramienta, para entenderla sin
  descargar nada.
- Instrucciones de **instalación global** (`npm i -g @skapxd/excel2md` → comando
  `excel2md`).
- Primer `CHANGELOG`.

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

[1.4.0]: https://github.com/skapxd/excel2md/releases/tag/v1.4.0
[1.3.0]: https://github.com/skapxd/excel2md/releases/tag/v1.3.0
[1.2.1]: https://github.com/skapxd/excel2md/releases/tag/v1.2.1
[1.2.0]: https://github.com/skapxd/excel2md/releases/tag/v1.2.0
[1.1.1]: https://github.com/skapxd/excel2md/releases/tag/v1.1.1
[1.1.0]: https://github.com/skapxd/excel2md/releases/tag/v1.1.0
[1.0.0]: https://github.com/skapxd/excel2md/releases/tag/v1.0.0
