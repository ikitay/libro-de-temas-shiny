# Load required libraries
library(shiny)
library(readxl)
library(dplyr)
library(openxlsx)
library(lubridate)
library(tidyr)
library(shinyjs)

PATH_DATOS <- "datos.xlsx"

REQUIRED_SHEETS <- c(
  "Feriados nacionales",
  "No laborables judíos",
  "No laborables armenios",
  "No laborables musulmanes",
  "Cronograma por distrito"
)

CRONOGRAMA_REQUIRED_COLS <- c(
  "distrito",
  "inicio ciclo lectivo",
  "fin ciclo lectivo",
  "inicio receso",
  "fin receso"
)

FERIADO_REQUIRED_COLS <- c("Fecha")

# --- Helpers: validation and robustness ---

validate_excel_sheets <- function(path) {
  have <- readxl::excel_sheets(path)
  missing <- setdiff(REQUIRED_SHEETS, have)
  if (length(missing) > 0) {
    stop(
      "Faltan hojas en datos.xlsx: ", paste(missing, collapse = ", "),
      ". Hojas encontradas: ", paste(have, collapse = ", ")
    )
  }
}

require_columns <- function(df, cols, context) {
  bad <- setdiff(cols, names(df))
  if (length(bad) > 0) {
    stop(
      "Faltan columnas en ", context, ": ", paste(bad, collapse = ", "),
      ". Columnas encontradas: ", paste(names(df), collapse = ", ")
    )
  }
}

ensure_descripcion_col <- function(df, context) {
  if ("Descripción" %in% names(df)) {
    return(df)
  }
  if ("Descripcion" %in% names(df)) {
    names(df)[names(df) == "Descripcion"] <- "Descripción"
    return(df)
  }
  stop(
    "Falta la columna de descripción (Descripción) en ", context, ". ",
    "Columnas encontradas: ", paste(names(df), collapse = ", ")
  )
}

check_dates_no_na <- function(dates, context) {
  n_na <- sum(is.na(dates))
  if (n_na > 0) {
    stop(
      "Hay ", n_na, " fecha(s) inválida(s) o vacía(s) en ", context,
      " tras el parseo. Revisá el formato (dd-mm-yyyy) o celdas vacías."
    )
  }
}

dedupe_feriados_por_fecha <- function(df) {
  if (!"Fecha" %in% names(df)) {
    return(df)
  }
  df %>% distinct(Fecha, .keep_all = TRUE)
}

sanitize_filename_windows <- function(x) {
  gsub("[\\\\/:*?\"<>|]+", "_", x)
}

# Join nacionales + opcional colectividad: misma lógica que antes, con menos duplicación (P2)
unir_feriados_temas <- function(filas, nacionales, feriados_especiales) {
  nat <- nacionales %>%
    select(Fecha, Descripcion = Descripción)

  if (is.null(feriados_especiales)) {
    return(
      filas %>%
        left_join(nat, by = "Fecha") %>%
        mutate(Tema = ifelse(!is.na(Descripcion), Descripcion, "")) %>%
        select(-Descripcion)
    )
  }

  esp <- feriados_especiales %>%
    select(Fecha, DescripcionEspecial = Descripción)

  filas %>%
    left_join(nat, by = "Fecha") %>%
    left_join(esp, by = "Fecha") %>%
    mutate(
      Tema = ifelse(
        !is.na(Descripcion), Descripcion,
        ifelse(!is.na(DescripcionEspecial), DescripcionEspecial, "")
      )
    ) %>%
    select(-Descripcion, -DescripcionEspecial)
}

# P0: validate contract then load
validate_excel_sheets(PATH_DATOS)

feriados_nacionales <- read_excel(PATH_DATOS, sheet = "Feriados nacionales")
feriados_nacionales <- ensure_descripcion_col(feriados_nacionales, "Feriados nacionales")
require_columns(feriados_nacionales, c(FERIADO_REQUIRED_COLS, "Descripción"), "Feriados nacionales")

feriados_judios     <- read_excel(PATH_DATOS, sheet = "No laborables judíos")
feriados_judios     <- ensure_descripcion_col(feriados_judios, "No laborables judíos")
require_columns(feriados_judios, c(FERIADO_REQUIRED_COLS, "Descripción"), "No laborables judíos")

feriados_armenios   <- read_excel(PATH_DATOS, sheet = "No laborables armenios")
feriados_armenios   <- ensure_descripcion_col(feriados_armenios, "No laborables armenios")
require_columns(feriados_armenios, c(FERIADO_REQUIRED_COLS, "Descripción"), "No laborables armenios")

feriados_musulmanes <- read_excel(PATH_DATOS, sheet = "No laborables musulmanes")
feriados_musulmanes <- ensure_descripcion_col(feriados_musulmanes, "No laborables musulmanes")
require_columns(feriados_musulmanes, c(FERIADO_REQUIRED_COLS, "Descripción"), "No laborables musulmanes")

cronogramas <- read_excel(PATH_DATOS, sheet = "Cronograma por distrito")
require_columns(cronogramas, CRONOGRAMA_REQUIRED_COLS, "Cronograma por distrito")

# Ensure dates are in the correct format
# For holiday tables: explicit format since dates are often stored as strings (dd-mm-yyyy)
# For cronogramas: across() converts all columns except "distrito" to Date
feriados_nacionales$Fecha <- as.Date(feriados_nacionales$Fecha, format = "%d-%m-%Y")
feriados_judios$Fecha     <- as.Date(feriados_judios$Fecha,     format = "%d-%m-%Y")
feriados_armenios$Fecha  <- as.Date(feriados_armenios$Fecha,   format = "%d-%m-%Y")
feriados_musulmanes$Fecha <- as.Date(feriados_musulmanes$Fecha, format = "%d-%m-%Y")

check_dates_no_na(feriados_nacionales$Fecha, "Feriados nacionales$Fecha")
check_dates_no_na(feriados_judios$Fecha, "No laborables judíos$Fecha")
check_dates_no_na(feriados_armenios$Fecha, "No laborables armenios$Fecha")
check_dates_no_na(feriados_musulmanes$Fecha, "No laborables musulmanes$Fecha")

cronogramas <- cronogramas %>% mutate(across(-distrito, as.Date))
for (cn in c(
  "inicio ciclo lectivo",
  "fin ciclo lectivo",
  "inicio receso",
  "fin receso"
)) {
  check_dates_no_na(cronogramas[[cn]], paste0("Cronograma por distrito$", cn))
}

# P1: deduplicate holiday rows by date before joins (avoids row multiplication)
feriados_nacionales <- dedupe_feriados_por_fecha(feriados_nacionales)
feriados_judios     <- dedupe_feriados_por_fecha(feriados_judios)
feriados_armenios   <- dedupe_feriados_por_fecha(feriados_armenios)
feriados_musulmanes <- dedupe_feriados_por_fecha(feriados_musulmanes)

# Distrito choices: trim, unique, no NA (P2-style cleanup used for UI)
DISTRITO_CHOICES <- unique(trimws(as.character(cronogramas$distrito)))
DISTRITO_CHOICES <- DISTRITO_CHOICES[!is.na(DISTRITO_CHOICES) & nzchar(DISTRITO_CHOICES)]
if (length(DISTRITO_CHOICES) == 0) {
  stop("No hay distritos disponibles en la hoja 'Cronograma por distrito'.")
}
DISTRITO_DEFAULT <- if ("CABA" %in% DISTRITO_CHOICES) {
  "CABA"
} else {
  DISTRITO_CHOICES[[1]]
}

default_cron_fila <- cronogramas[cronogramas$distrito == DISTRITO_DEFAULT, , drop = FALSE][1L, , drop = FALSE]
DEFAULT_CICLO_START <- default_cron_fila$`inicio ciclo lectivo`
DEFAULT_CICLO_END <- default_cron_fila$`fin ciclo lectivo`
DEFAULT_RECESO_START <- default_cron_fila$`inicio receso`
DEFAULT_RECESO_END <- default_cron_fila$`fin receso`

FECHA_MIN_UI <- as.Date(format(min(c(
  cronogramas[["inicio ciclo lectivo"]],
  cronogramas[["fin ciclo lectivo"]],
  cronogramas[["inicio receso"]],
  cronogramas[["fin receso"]]
), na.rm = TRUE), "%Y-01-01"))

FECHA_MAX_UI <- as.Date(format(max(c(
  cronogramas[["inicio ciclo lectivo"]],
  cronogramas[["fin ciclo lectivo"]],
  cronogramas[["inicio receso"]],
  cronogramas[["fin receso"]]
), na.rm = TRUE), "%Y-12-31"))

MSG_SIN_FILAS <- "No se pudo generar ninguna fila con la combinación de opciones elegidas."

# Define weekday mapping (ordered vector for UI display)
WEEKDAY_MAPPING <- c("Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado")

# Define the UI
ui <- fluidPage(
  useShinyjs(),
  titlePanel("Generar Libro de Temas"),

  sidebarLayout(
    sidebarPanel(
      # File title input
      textInput("titulo_archivo", "Ingrese un nombre para la planilla:",
            placeholder = "Ej. Porteña Modelo - Matemática - 4to A"),
      tags$div(style = "border-top: 1px solid #ddd; margin: 10px 0;"),

      # District selector: automatically defines the academic period and winter break dates
      selectInput("distrito", "Seleccione el distrito:",
                  choices = DISTRITO_CHOICES,
                  selected = DISTRITO_DEFAULT),
      tags$p("Las fechas generadas corresponderán al calendario oficial para el nivel medio del distrito seleccionado.", style = "font-size: 90%;"),

      fluidRow(
        column(7, checkboxInput("personalizar_ciclo", "Personalizar fechas del ciclo", value = FALSE)),
        column(
          5,
          tags$p(
            "Seleccione para elegir fechas libremente.",
            style = "font-size: 90%; margin-top: 10px;"
          )
        )
      ),
      conditionalPanel(
        condition = "input.personalizar_ciclo == true",
        dateRangeInput(
          "rango_ciclo",
          label = "Rango del ciclo lectivo:",
          start = DEFAULT_CICLO_START,
          end = DEFAULT_CICLO_END,
          min = FECHA_MIN_UI,
          max = FECHA_MAX_UI,
          format = "dd/mm/yyyy",
          language = "es",
          separator = " a "
        )
      ),
      conditionalPanel(
        condition = "input.personalizar_ciclo == true",
        tags$p(
          style = "font-size: 90%; color: #555;",
          "Las fechas personalizadas pueden no coincidir con el calendario oficial del distrito."
        )
      ),

      tags$div(style = "border-top: 1px solid #ddd; margin: 10px 0;"),

      selectInput("calendario_especial", "Agregar asuetos de calendarios especiales (colectividades):",
                  choices = c("Ninguno", "Comunidad Judía", "Comunidad Armenia", "Comunidad Islámica"),
                  selected = "Ninguno"),
      tags$p("El calendario de cada institución puede ser ligeramente diferente.", style = "font-size: 90%;"),

      tags$div(style = "border-top: 1px solid #ddd; margin: 10px 0;"),

      # Class day selector and hours per day
      tags$label("Seleccione los días de clase:"),
      div(
        style = "column-count: 2; column-gap: 1rem;",
        checkboxGroupInput("dias_seleccionados",
                           label = NULL,
                           choices = WEEKDAY_MAPPING,
                           selected = NULL)
      ),

      # Inputs for class hours per day
      uiOutput("horasCatedraInputs"),

      tags$div(style = "border-top: 1px solid #ddd; margin: 10px 0;"),
      tags$p(tags$strong("Receso académico"), style = "font-size: 110%;"),
      fluidRow(
        column(7, checkboxInput("sin_receso", "Sin receso académico", value = FALSE)),
        column(
          5,
          tags$p(
            "Seleccione para omitir el receso académico oficial.",
            style = "font-size: 90%; margin-top: 10px;"
          )
        )
      ),
      fluidRow(
        column(7, checkboxInput("personalizar_receso", "Personalizar fechas del receso", value = FALSE)),
        column(
          5,
          tags$p(
            "Seleccione para elegir fechas libremente.",
            style = "font-size: 90%; margin-top: 10px;"
          )
        )
      ),
      conditionalPanel(
        condition = "input.personalizar_receso == true && input.sin_receso == false",
        dateRangeInput(
          "rango_receso",
          label = "Rango del receso académico:",
          start = DEFAULT_RECESO_START,
          end = DEFAULT_RECESO_END,
          min = FECHA_MIN_UI,
          max = FECHA_MAX_UI,
          format = "dd/mm/yyyy",
          language = "es",
          separator = " a "
        )
      ),
      checkboxInput("excluir_receso", "Ocultar fechas del período de receso académico.", value = TRUE),

      tags$div(style = "border-top: 1px solid #ddd; margin: 10px 0;"),
      # Button to generate the list
      actionButton("generate", strong("Generar planilla"), class = "btn-primary btn-block"),

      # Button to download the file (initially disabled)
      downloadButton("downloadExcel", strong("Descargar planilla (.xlsx)"), class = "btn-success btn-block"),

      # Conditionally show the title reminder text based on whether title is empty
      uiOutput("titleReminder"),

      tags$div(style = "border-top: 1px solid #ddd; margin: 10px 0;"),
      tags$p(
        tags$strong(
          "¿Tenés comentarios o sugerencias? ",
          tags$a(
            href = "https://tally.so/r/q4AzG8",
            "Dejalas acá.",
            style = "font-size: 120%;",
            target = "_blank"
          )
        ),
        style = "font-size: 110%;"
      ),
      div(
        style = "display: flex; align-items: center; justify-content: space-between;",
        tags$p(tags$strong("¿Te sirvió la app y querés apoyar su desarrollo?"), style = "font-size: 100%; margin: 0;"),
        tags$a(
          href = "https://cafecito.app/3runtalon",
          rel = "noopener",
          target = "_blank",
          tags$img(
            srcset = "https://cdn.cafecito.app/imgs/buttons/button_6.png 1x, https://cdn.cafecito.app/imgs/buttons/button_6_2x.png 2x, https://cdn.cafecito.app/imgs/buttons/button_6_3.75x.png 3.75x",
            src = "https://cdn.cafecito.app/imgs/buttons/button_6.png",
            alt = "Invitame un café en cafecito.app",
            style = "height: 40px;"
          )
        )
      )
    ),

    mainPanel(
      div(
        style = "overflow-x: auto",
        tableOutput("resultTable")
      )
    )
  )
)

# Define the server
server <- function(input, output, session) {
  # Reactive value to store results
  datos_resultantes <- reactiveVal()

  # Initially disable the download button
  shinyjs::disable("downloadExcel")

  observe({
    if (isTRUE(input$sin_receso)) {
      shinyjs::disable("personalizar_receso")
      shinyjs::disable("rango_receso")
      shinyjs::disable("excluir_receso")
    } else {
      shinyjs::enable("personalizar_receso")
      shinyjs::enable("rango_receso")
      shinyjs::enable("excluir_receso")
    }
  })

  observeEvent(input$sin_receso, {
    if (isTRUE(input$sin_receso)) {
      updateCheckboxInput(session, "personalizar_receso", value = FALSE)
    }
  }, ignoreNULL = TRUE)

  observeEvent(input$distrito, {
    fila <- cronogramas[cronogramas$distrito == input$distrito, , drop = FALSE]
    if (nrow(fila) < 1) return()
    fila <- fila[1, , drop = FALSE]
    if (!isTRUE(input$personalizar_ciclo)) {
      updateDateRangeInput(
        session, "rango_ciclo",
        start = fila$`inicio ciclo lectivo`,
        end = fila$`fin ciclo lectivo`
      )
    }
    if (!isTRUE(input$personalizar_receso) && !isTRUE(input$sin_receso)) {
      updateDateRangeInput(
        session, "rango_receso",
        start = fila$`inicio receso`,
        end = fila$`fin receso`
      )
    }
  }, ignoreInit = TRUE)

  observeEvent(input$personalizar_ciclo, {
    if (!isTRUE(input$personalizar_ciclo)) {
      req(input$distrito)
      fila <- cronogramas[cronogramas$distrito == input$distrito, , drop = FALSE]
      if (nrow(fila) < 1) return()
      fila <- fila[1, , drop = FALSE]
      updateDateRangeInput(
        session, "rango_ciclo",
        start = fila$`inicio ciclo lectivo`,
        end = fila$`fin ciclo lectivo`
      )
    }
  }, ignoreNULL = TRUE)

  observeEvent(input$personalizar_receso, {
    if (!isTRUE(input$personalizar_receso)) {
      req(input$distrito)
      if (isTRUE(input$sin_receso)) return()
      fila <- cronogramas[cronogramas$distrito == input$distrito, , drop = FALSE]
      if (nrow(fila) < 1) return()
      fila <- fila[1, , drop = FALSE]
      updateDateRangeInput(
        session, "rango_receso",
        start = fila$`inicio receso`,
        end = fila$`fin receso`
      )
    }
  }, ignoreNULL = TRUE)

  # Enable/disable the download button based on titulo_archivo and if data is generated
  observe({
    d <- datos_resultantes()
    if (nzchar(input$titulo_archivo) && !is.null(d) && nrow(d) > 0) {
      shinyjs::enable("downloadExcel")
    } else {
      shinyjs::disable("downloadExcel")
    }
  })

  # Dynamically generate inputs for class hours based on selected days
  output$horasCatedraInputs <- renderUI({
    req(input$dias_seleccionados)

    tagList(
      lapply(input$dias_seleccionados, function(dia) {
        selectInput(
          inputId = paste0("horas_", dia),
          label = paste0("Cantidad de horas del día ", dia, ":"),
          choices = 1:4,
          selected = 1
        )
      }),
      tags$p("Repite la fila por cada hora.", style = "font-size: 90%;"),
      tags$p("Para generar una sola fila por día, elegir el valor 1.", style = "font-size: 90%;")
    )
  })

  # Show warning if data is generated but no title is provided
  output$titleReminder <- renderUI({
    if (!is.null(datos_resultantes()) && input$titulo_archivo == "") {
      tags$p(
        tags$strong("Recordá ingresar un nombre para poder descargar la planilla."),
        style = "color: red; font-size: 90%; margin-top: 10px;"
      )
    }
  })

  # Reactives: academic cycle and effective recess (PLAN 2.2)
  fechas_completas <- reactive({
    if (isTRUE(input$personalizar_ciclo)) {
      req(input$rango_ciclo)
      r <- input$rango_ciclo
      if (length(r) < 2 || any(is.na(r)) || r[1] > r[2]) {
        return(as.Date(character(0)))
      }
      return(seq(as.Date(r[1]), as.Date(r[2]), by = "day"))
    }
    req(input$distrito)
    fila <- cronogramas[cronogramas$distrito == input$distrito, , drop = FALSE]
    if (nrow(fila) < 1) {
      return(as.Date(character(0)))
    }
    fila <- fila[1, , drop = FALSE]
    seq(fila$`inicio ciclo lectivo`, fila$`fin ciclo lectivo`, by = "day")
  })

  receso_fechas <- reactive({
    if (isTRUE(input$sin_receso)) {
      return(NULL)
    }
    req(input$distrito)
    if (isTRUE(input$personalizar_receso)) {
      req(input$rango_receso)
      r <- input$rango_receso
      if (length(r) < 2 || any(is.na(r)) || r[1] > r[2]) {
        return(c(as.Date(NA), as.Date(NA)))
      }
      return(c(as.Date(r[1]), as.Date(r[2])))
    }
    fila <- cronogramas[cronogramas$distrito == input$distrito, , drop = FALSE]
    if (nrow(fila) < 1) {
      return(c(as.Date(NA), as.Date(NA)))
    }
    fila <- fila[1, , drop = FALSE]
    c(fila$`inicio receso`, fila$`fin receso`)
  })

  output$resultTable <- renderTable(
    {
      validate(need(!is.null(datos_resultantes()), "Generá la planilla con el botón «Generar planilla»."))
      validate(need(nrow(datos_resultantes()) > 0, MSG_SIN_FILAS))
      datos_resultantes()
    },
    bordered = TRUE
  )

  observeEvent(input$generate, {
    tryCatch(
      {
        if (is.null(input$dias_seleccionados) || length(input$dias_seleccionados) == 0) {
          showNotification("Por favor, seleccione los días de clase.", type = "error")
          return()
        }

        # P0: validate dynamic hour inputs before computing
        horas_ids <- paste0("horas_", input$dias_seleccionados)
        faltan_horas <- vapply(horas_ids, function(id) is.null(input[[id]]), logical(1))
        if (any(faltan_horas)) {
          showNotification(
            "Esperá un instante y volvé a generar: aún se están cargando los campos de horas cátedra.",
            type = "error"
          )
          return()
        }

        if (isTRUE(input$personalizar_ciclo)) {
          r <- input$rango_ciclo
          if (is.null(r) || length(r) < 2 || any(is.na(r))) {
            showNotification("Indicá un rango válido para el ciclo lectivo.", type = "error")
            return()
          }
          if (r[1] > r[2]) {
            showNotification(
              "La fecha de inicio del ciclo debe ser anterior o igual a la de fin.",
              type = "error"
            )
            return()
          }
        }

        fechas_vector <- fechas_completas()
        if (length(fechas_vector) == 0) {
          showNotification(
            "No hay un rango de ciclo lectivo válido. Revisá el distrito o las fechas personalizadas.",
            type = "error"
          )
          return()
        }

        sin_rec <- isTRUE(input$sin_receso)

        if (!sin_rec && isTRUE(input$personalizar_receso)) {
          rr <- input$rango_receso
          if (is.null(rr) || length(rr) < 2 || any(is.na(rr))) {
            showNotification("Indicá un rango válido para el receso académico.", type = "error")
            return()
          }
          if (rr[1] > rr[2]) {
            showNotification(
              "La fecha de inicio del receso debe ser anterior o igual a la de fin.",
              type = "error"
            )
            return()
          }
        }

        rango_receso <- if (sin_rec) NULL else receso_fechas()

        if (!sin_rec) {
          if (is.null(rango_receso) || length(rango_receso) < 2 || any(is.na(rango_receso))) {
            showNotification(
              "Faltan fechas de receso académico para el distrito seleccionado en datos.xlsx.",
              type = "error"
            )
            return()
          }

          ciclo_start <- min(fechas_vector)
          ciclo_end <- max(fechas_vector)
          rec_start <- rango_receso[1]
          rec_end <- rango_receso[2]
          if (max(rec_start, ciclo_start) > min(rec_end, ciclo_end)) {
            showNotification(
              "El período de receso elegido está fuera del ciclo lectivo elegido.",
              type = "error"
            )
            return()
          }
        }

        excluir_efectivo <- isTRUE(input$excluir_receso) && !sin_rec

        fechas_excluidas <- if (excluir_efectivo && !is.null(rango_receso)) {
          seq(rango_receso[1], rango_receso[2], by = "day")
        } else {
          NULL
        }

        fechas_finales <- setdiff(fechas_vector, fechas_excluidas)
        fechas_finales <- as.Date(fechas_finales, origin = "1970-01-01")

        selected_weekday_numbers <- match(input$dias_seleccionados, WEEKDAY_MAPPING)
        if (any(is.na(selected_weekday_numbers))) {
          showNotification("Días de clase inválidos. Volvé a seleccionar los días.", type = "error")
          return()
        }

        fechas_dias <- data.frame(
          Fecha = fechas_finales,
          weekday_num = wday(fechas_finales, week_start = 1),
          stringsAsFactors = FALSE
        )

        fechas_dias <- fechas_dias %>%
          filter(weekday_num %in% selected_weekday_numbers)

        if (nrow(fechas_dias) == 0) {
          showNotification(MSG_SIN_FILAS, type = "warning")
          empty_df <- setNames(
            data.frame(character(0), character(0), character(0), character(0)),
            c("Clase Nº", "Fecha", "Día de la Semana", "Tema")
          )
          datos_resultantes(empty_df)
          return()
        }

        filas_finales <- fechas_dias %>%
          group_by(Fecha, weekday_num) %>%
          mutate(
            HsCatedra = list(1:input[[paste0("horas_", WEEKDAY_MAPPING[weekday_num])]])
          ) %>%
          unnest(cols = c(HsCatedra)) %>%
          ungroup()

        feriados_especiales <- switch(
          input$calendario_especial,
          "Comunidad Judía" = feriados_judios,
          "Comunidad Armenia" = feriados_armenios,
          "Comunidad Islámica" = feriados_musulmanes,
          NULL
        )

        filas_finales <- unir_feriados_temas(filas_finales, feriados_nacionales, feriados_especiales)

        filas_finales <- filas_finales %>%
          arrange(Fecha, HsCatedra) %>%
          group_by(Fecha) %>%
          mutate(ClaseNumero = ifelse(Tema != "", NA, row_number())) %>%
          ungroup()

        filas_finales <- filas_finales %>%
          mutate(ClaseNumero = ifelse(
            is.na(ClaseNumero), NA, cumsum(!is.na(ClaseNumero) & !duplicated(Fecha))
          ))

        filas_finales <- filas_finales %>%
          mutate(ClaseNumero = ifelse(Tema != "", "-", as.character(ClaseNumero)))

        if (excluir_efectivo) {
          fila_receso <- data.frame(
            ClaseNumero = "-",
            Fecha = rango_receso[1],
            weekday_num = NA_real_,
            HsCatedra = NA_real_,
            Tema = "RECESO ACADÉMICO",
            stringsAsFactors = FALSE
          )

          fin_receso <- rango_receso[2]
          n <- nrow(filas_finales)
          if (n == 0L) {
            filas_finales <- fila_receso
          } else {
            idx_posteriores <- which(filas_finales$Fecha > fin_receso)
            if (length(idx_posteriores) == 0L) {
              filas_finales <- bind_rows(filas_finales, fila_receso)
            } else {
              pos <- min(idx_posteriores)
              part1 <- if (pos > 1L) {
                filas_finales[seq_len(pos - 1L), , drop = FALSE]
              } else {
                filas_finales[0L, , drop = FALSE]
              }
              part2 <- filas_finales[seq.int(pos, n, 1L), , drop = FALSE]
              filas_finales <- bind_rows(part1, fila_receso, part2)
            }
          }
        }

        filas_finales <- filas_finales %>%
          mutate(
            DiaSemana = if_else(
              !is.na(weekday_num),
              WEEKDAY_MAPPING[as.integer(weekday_num)],
              NA_character_
            )
          ) %>%
          select(ClaseNumero, Fecha, DiaSemana, Tema)

        names(filas_finales) <- c("Clase Nº", "Fecha", "Día de la Semana", "Tema")
        filas_finales$Fecha <- format(filas_finales$Fecha, "%d-%m-%Y")

        datos_resultantes(filas_finales)
      },
      error = function(e) {
        showNotification(
          conditionMessage(e),
          type = "error",
          duration = NULL
        )
      }
    )
  })

  output$downloadExcel <- downloadHandler(
    filename = function() {
      base <- sanitize_filename_windows(input$titulo_archivo)
      paste0(base, " - Planificación diaria anual.xlsx")
    },
    content = function(file) {
      validate(need(!is.null(datos_resultantes()), "No hay datos generados aún."))
      validate(need(nrow(datos_resultantes()) > 0, "No hay filas para exportar."))
      d <- datos_resultantes()

      wb <- createWorkbook()
      addWorksheet(wb, "Planificación")

      estilo_bordes <- createStyle(
        border = c("top", "bottom", "left", "right"),
        borderStyle = "thin"
      )

      estilo_encabezado <- createStyle(
        textDecoration = "bold",
        border = c("top", "bottom", "left", "right"),
        borderStyle = "thin"
      )

      writeData(wb, "Planificación", d, startRow = 1, startCol = 1)

      addStyle(
        wb, "Planificación", estilo_encabezado,
        rows = 1, cols = 1:ncol(d), gridExpand = TRUE, stack = TRUE
      )

      addStyle(
        wb, "Planificación", estilo_bordes,
        rows = 1:(nrow(d) + 1), cols = 1:ncol(d), gridExpand = TRUE, stack = TRUE
      )

      saveWorkbook(wb, file, overwrite = TRUE)
    }
  )
}

# Run the application
shinyApp(ui = ui, server = server)
