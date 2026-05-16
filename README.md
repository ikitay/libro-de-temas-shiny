# 📅 Generador de Libros de Temas

> Aplicación web desarrollada en R/Shiny que automatiza la generación de planillas de planificación anual para docentes de nivel medio en Argentina — construida desde la experiencia directa en el aula.

<!-- BADGES -->
![R](https://img.shields.io/badge/R-4.x-276DC3?logo=r&logoColor=white)
![Shiny](https://img.shields.io/badge/Shiny-deployed-blue?logo=rstudio&logoColor=white)
![Status](https://img.shields.io/badge/status-active-brightgreen)
![License](https://img.shields.io/badge/license-MIT-lightgrey)

**[▶ Demo en vivo →](https://ikitay.shinyapps.io/libro-de-temas-shiny/)**

---

## El problema

Los docentes de nivel medio en Argentina están obligados a presentar un *libro de temas* — un registro diario que asocia cada clase con una fecha, un día de la semana y un tema- de acuerdo a un formato oficial definido por el Ministerio de Educación correspondiente. Construir la plantilla de cronograma es enteramente manual, realizado habitualmente por personal administrativo junto a docentes: calcular qué días hay clase, cruzar contra el calendario escolar oficial, excluir feriados nacionales y asuetos de colectividades, numerar las clases correlativamente, contemplar el receso invernal, y repetirlo para cada materia y división.

Esta app elimina ese proceso por completo.

---

## Qué hace

A partir del distrito escolar, los días de clase y la carga horaria semanal, la app genera en segundos una planilla de planificación completa y lista para completar, siguiendo el formato oficial.

- **Cronograma por distrito** — las fechas de inicio y fin del ciclo lectivo y del receso invernal se cargan automáticamente desde una tabla de referencia que cubre múltiples distritos de Buenos Aires
- **Soporte de calendarios múltiples** — aplica feriados nacionales argentinos más un calendario de colectividad opcional (judía, armenia o islámica), cubriendo la diversidad institucional de las escuelas porteñas
- **Filas configurables por hora** — cada día puede generar entre 1 y 4 filas, una por hora cátedra, ajustándose a la carga semanal de la materia
- **Exportación Excel formateada** — descarga un archivo `.xlsx` con encabezados en negrita, bordes en todas las celdas y hoja con nombre, listo para completar
- **Numeración correlativa de clases** — los feriados y el receso quedan etiquetados y excluidos del conteo automáticamente

---

## Capturas

<!-- Agregá una captura o grabación de pantalla de la app -->
<!-- Recomendado: un GIF que muestre el flujo generar → previsualizar → descargar -->

![Captura de la app](screenshot.png)

---

## Aspectos técnicos destacados

El proyecto atravesó tres ciclos de release documentados. Algunas decisiones relevantes:

**Flujo de datos reactivo** — el cálculo de fechas es completamente reactivo: seleccionar un distrito actualiza el calendario académico, que alimenta el filtro de fechas, que genera la tabla de salida. Sin necesidad de reenvíos manuales.

**Validación de contrato al inicio** — v2.1 introdujo una capa de validación que corre antes de que se inicialice la sesión de Shiny. Verifica hojas y columnas requeridas en el Excel, normaliza inconsistencias de encoding (`Descripción` vs `Descripcion`), deduplica filas de feriados para prevenir multiplicación en los joins, y falla con mensajes descriptivos en lugar de errores silenciosos en runtime.

**Exportación Excel estilizada** — migración de `writexl` (exportación plana) a `openxlsx` (workbook formateado) para producir archivos inmediatamente utilizables sin post-procesamiento.

**Lógica de generación defensiva** — el paso de generación del cronograma maneja edge cases: race conditions de `renderUI` en inputs dinámicos, resultados con cero filas, y tres posiciones de inserción del receso invernal que anteriormente no estaban contempladas.

Ver [Releases](../../releases) para el changelog completo.

---

## Stack

| Capa | Herramientas |
|---|---|
| Framework de la app | R · Shiny · shinyjs |
| Transformación de datos | dplyr · tidyr · lubridate |
| Lectura de datos | readxl |
| Exportación Excel | openxlsx |
| Deploy | shinyapps.io |

---

## Arquitectura de datos
 
La app consume un archivo Excel externo (`datos.xlsx`, no incluido en el repo) que actúa como fuente de verdad para todos los datos de referencia del ciclo lectivo. Está organizado en cinco hojas con contratos de columnas validados al inicio de la sesión:
 
| Hoja | Descripción |
|---|---|
| `Feriados nacionales` | Feriados nacionales argentinos con fecha y descripción |
| `No laborables judíos` | Asuetos del calendario de la comunidad judía |
| `No laborables armenios` | Asuetos del calendario de la comunidad armenia |
| `No laborables musulmanes` | Asuetos del calendario de la comunidad islámica |
| `Cronograma por distrito` | Límites del ciclo lectivo y receso invernal por distrito |
 
La separación entre lógica de la app y datos de referencia permite actualizar el calendario anualmente sin modificar el código.
 
---

## Historial de versiones

| Versión | Descripción |
|---|---|
| [v1.0](../../releases/tag/v1.0) | MVP — CABA específico, ingreso manual de fechas, exportación con `writexl` |
| [v2.0](../../releases/tag/v2.0) | Soporte multi-distrito, calendarios de 3 colectividades, exportación Excel estilizada |
| [v2.1](../../releases/tag/v2.1) | Validación de datos al inicio, manejo de errores con `tryCatch`, hardening de edge cases |
| [v2.2](../../releases/tag/v2.2) | Mejoras de interfaz y corrección del rango de fechas en los selectores de calendario |

---

## Sobre el proyecto

Trabajé varios años como docente de nivel medio en Buenos Aires. El *libro de temas* es una constante burocrática en la educación argentina — cada docente lo completa, cada año, para cada materia que dicta.

En algún momento me cansé de reconstruir la misma planilla de Excel a mano cada marzo y escribí un script para generarla. Ese script se convirtió en esta app.

Fue utilizada por colegas durante el ciclo lectivo 2025. Las iteraciones de v1.0 a v2.1 reflejan feedback real de usuarios reales y problemas de datos reales que aparecieron en producción.

Hoy estoy construyendo sobre esta base para hacer la transición al análisis de datos y el desarrollo de forma profesional. Este proyecto refleja cómo me acerco a los problemas: identificar el proceso manual, automatizarlo, documentar las decisiones y entregar algo que la gente pueda usar.

---

## Contacto

**[LinkedIn →](https://www.linkedin.com/in/ivankitay/)**  
**[GitHub →](https://github.com/ikitay/)**

---

---

## 🇬🇧 Project summary (English)

**Libro de Temas Generator** is a parameterized R/Shiny web app that automates the creation of academic planning spreadsheets for Argentine secondary school teachers.

Given a school district, a set of class days, and hours per day, the app outputs a date-indexed `.xlsx` schedule with national holidays, optional community calendars (Jewish, Armenian, Islamic), winter recess, and sequential session numbering — formatted and ready to submit.

Built from real classroom experience and used in production during the 2025 school year. Three documented releases covering MVP → multi-district expansion → data validation and error handling.

**[▶ Live demo →](https://ikitay.shinyapps.io/libro-de-temas-shiny/)**

**Stack:** R · Shiny · dplyr · tidyr · lubridate · readxl · openxlsx · shinyjs
