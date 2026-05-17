# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [2.2.0] - YYYY-MM-DD

### Added

- Cycle date customization: optional override of the district's default start/end dates via `dateRangeInput`
- Recess options expanded: toggle to skip recess entirely, and optional custom recess date range
- `FECHA_MIN_UI` / `FECHA_MAX_UI` computed from data to bound all date range selectors
- Observers to keep cycle and recess inputs in sync when the district changes or customization is toggled off
- `sin_receso` toggle disables recess-related controls reactively

### Changed

- App renamed to "Generar Libro de Temas"
- Recess label in the output changed from "RECESO INVERNAL" to "RECESO ACADÉMICO"
- Weekday checkboxes now render in a two-column layout
- Sidebar section separators replaced with styled `div` borders for visual consistency
- Hour input label text simplified

### Fixed

- Date range selectors now respect bounds derived from the loaded data, preventing out-of-range date selection

## [2.1.0] - YYYY-MM-DD

### Added

- Contract validation layer that runs before Shiny initializes: verifies required sheets and columns in `datos.xlsx`, fails with descriptive errors instead of silent runtime crashes
- Encoding normalization for the `Descripción`/`Descripcion` column name inconsistency across holiday sheets
- Holiday row deduplication by date before joins, preventing row multiplication in the output
- `unir_feriados_temas()` helper to consolidate the national + community calendar join logic
- Filename sanitization: characters invalid on Windows are replaced before building the `.xlsx` filename
- `tryCatch` wrapper around the full generation logic — errors surface as notifications instead of crashing the session
- Guard for hora cátedra inputs not yet rendered: detects the `renderUI` race condition and prompts the user to retry
- `validate(need(...))` guards in `renderTable` and the download handler
- Warning notification when the selected combination produces zero rows
- District choices are now trimmed, deduplicated, and NA-filtered before populating the selector; `DISTRITO_DEFAULT` falls back to the first available district if CABA is not present

### Fixed

- Recess row insertion now handles three previously unaddressed edge cases: empty schedule, all rows falling before the recess, and all rows falling after it

## [2.0.0] - YYYY-MM-DD

### Added

- External reference data (`datos.xlsx`) with district schedules and holiday calendars
- Multi-district support via district selector; cycle and recess dates load automatically per district
- Armenian and Islamic community calendars, in addition to the existing Jewish calendar
- Community calendar selector replacing the previous binary checkbox
- `shinyjs`: download button disabled until data is generated and a title is entered
- Styled Excel export via `openxlsx`: bold headers, thin borders on all cells, sheet named "Planificación"
- Borders in the preview table
- Feedback link and donation button in the sidebar

### Changed

- Cycle and recess dates now derived from the selected district, replacing manual date pickers
- Export library switched from `writexl` to `openxlsx`
- Download filename updated to include " - Planificación diaria anual" suffix
- Title reminder now only appears after data has been generated

### Removed

- `shinyWidgets` dependency (`airDatepickerInput` replaced by district selector)
- Manual date pickers for cycle start/end and recess period
- CABA-hardcoded date defaults
- `incluir_judios` binary checkbox

## [1.0.0] - YYYY-MM-DD

Initial release.

### Added

- Schedule generated from a manually entered academic period and recess range via `airDatepickerInput`
- Weekday selector with configurable hours per day (1–4), generating one row per class hour
- Optional Jewish community calendar applied via checkbox
- National holidays joined and labeled automatically in the "Tema" column
- Sequential class numbering, skipping holiday and recess rows
- Recess row inserted at the correct position in the output
- Excel export via `writexl`
- Download button disabled until a title is entered

[Unreleased]: https://github.com/ikitay/libro-de-temas-shiny/compare/v2.2.0...HEAD
[2.2.0]: https://github.com/ikitay/libro-de-temas-shiny/compare/v2.1.0...v2.2.0
[2.1.0]: https://github.com/ikitay/libro-de-temas-shiny/compare/v2.0.0...v2.1.0
[2.0.0]: https://github.com/ikitay/libro-de-temas-shiny/compare/v1.0.0...v2.0.0
[1.0.0]: https://github.com/ikitay/libro-de-temas-shiny/releases/tag/v1.0.0