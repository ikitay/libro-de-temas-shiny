# Load required libraries
library(shiny)
library(readxl)
library(dplyr)
library(writexl)
library(lubridate)
library(tidyr)
library(shinyWidgets)
library(shinyjs)

# Load data files
feriados_nacionales <- read_excel("Feriados 2025.xlsx", sheet = "Feriados Nacionales")
feriados_judios <- read_excel("Feriados 2025.xlsx", sheet = "Feriados judíos")

# Ensure dates are in the correct format
feriados_nacionales$Fecha <- as.Date(feriados_nacionales$Fecha, format = "%d-%m-%Y")
feriados_judios$Fecha <- as.Date(feriados_judios$Fecha, format = "%d-%m-%Y")

# Define weekday mapping (ordered vector for UI display)
WEEKDAY_MAPPING <- c("Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado")

# Define the UI
ui <- fluidPage(
  useShinyjs(),
  titlePanel("Generar cronograma para planificación anual"),
  
  sidebarLayout(
    sidebarPanel(
      # File title input
      textInput("titulo_archivo", "Ingrese un nombre para la planilla:", 
                placeholder = "Ej. Porteña Modelo - Matemática - 4to A"),
      
      tags$hr(),
      
      # Date range input for selecting start and end dates
      airDatepickerInput("date_range", "Seleccione el período académico:", 
                         range = TRUE, 
                         value = c("2025-05-02", "2025-12-19"),
                         separator = " a ",
                         language = "es",
                         dateFormat = "dd-MM-yyyy"),
      
      # Agregar información del período académico y link al calendario
      tags$p(
        tags$span(
          "Los valores por defecto corresponden al nivel medio de CABA. ",
          style = "font-size: 80%;"
        ),
        tags$a(
          href = "https://buenosaires.gob.ar/educacion/agenda-educativa",
          "Ver calendario escolar completo.",
          style = "font-size: 80%;",
          target = "_blank"
        )
      ),
      tags$hr(),
      
      # Checkboxes for class days
      tags$h4("Seleccione los días de clase:"),
      checkboxGroupInput("dias_seleccionados", 
                         label = NULL,
                         choices = WEEKDAY_MAPPING,
                         selected = NULL),
      
      # Inputs for class hours per day
      uiOutput("horasCatedraInputs"),
      
      tags$hr(),
      # Checkboxes for additional options
      checkboxInput("excluir_receso", "Excluir período de receso invernal.", value = TRUE),
      uiOutput("recesoInvernalUI"),
      checkboxInput("incluir_judios", "Incluir días no laborables del calendario escolar judío.", value = TRUE),
      
      tags$hr(),
      # Button to generate the list
      actionButton("generate", strong("Generar planilla"), class = "btn-primary btn-block"),
      
      # Button to download the file (initially disabled)
      downloadButton("downloadExcel", strong("Descargar planilla (.xlsx)"), class = "btn-success btn-block"),
      
      # Conditionally show the title reminder text based on whether title is empty
      uiOutput("titleReminder")
    ),
    
    mainPanel(
      div(style = 'overflow-x: auto',
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
  
  # Enable/disable the download button based on titulo_archivo and if data is generated
  observe({
    if (input$titulo_archivo != "" && !is.null(datos_resultantes())) {
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
        selectInput(inputId = paste0("horas_", dia), 
                    label = paste0("Cantidad de horas del día ", dia, ": (repite la fila por cada hora)"),
                    choices = 1:4,
                    selected = 1)
      }),
      tags$p("Para generar una sola fila por día, elegir el valor 1.", style = "font-size: 80%;")
    )
  })
  
  # Show or hide the winter break date range input based on the checkbox
  output$recesoInvernalUI <- renderUI({
    if (input$excluir_receso) {
      tagList(
        airDatepickerInput("receso_invernal", "Seleccione el período de receso invernal:", 
                           range = TRUE, 
                           value = c("2025-07-21", "2025-08-01"),
                           separator = " a ",
                           language = "es",
                           dateFormat = "dd-MM-yyyy"),
        tags$p("Los valores por defecto corresponden al nivel medio de CABA.", style = "font-size: 80%;")
      )
    }
  })
  
  # Conditional UI for title reminder - only show if title is empty
  output$titleReminder <- renderUI({
    if (input$titulo_archivo == "") {
      tags$p(tags$strong("Para descargar, ingrese un nombre para la planilla."), style = "font-size: 90%;")
    }
  })
  
  # Reactive dates based on user input
  fechas_completas <- reactive({
    seq(as.Date(input$date_range[1]), as.Date(input$date_range[2]), by = "day")
  })
  
  observeEvent(input$generate, {
    # Consolidated input validation
    if (is.null(input$dias_seleccionados) || length(input$dias_seleccionados) == 0) {
      showNotification("Por favor, seleccione los días de clase.", type = "error")
      return()
    }
    
    # Generate dates and exclude the specific period
    fechas_completas <- fechas_completas()
    
    # Modification for winter break based on checkbox
    fechas_excluidas <- if(input$excluir_receso) {
      seq(as.Date(input$receso_invernal[1]), as.Date(input$receso_invernal[2]), by = "day")
    } else {
      NULL
    }
    
    fechas_2025 <- setdiff(fechas_completas, fechas_excluidas)
    fechas_2025 <- as.Date(fechas_2025, origin = "1970-01-01")
    
    # Convert selected Spanish weekdays to lubridate numbers
    selected_weekday_numbers <- match(input$dias_seleccionados, WEEKDAY_MAPPING)
    
    # Create initial DataFrame with weekday numbers
    fechas_dias <- data.frame(
      Fecha = fechas_2025,
      weekday_num = wday(fechas_2025, week_start = 1),
      stringsAsFactors = FALSE
    )
    
    # Filter for selected weekdays using lubridate numbers
    fechas_dias <- fechas_dias %>%
      filter(weekday_num %in% selected_weekday_numbers)
    
    # Create final rows replicating according to class hours
    filas_finales <- fechas_dias %>%
      group_by(Fecha, weekday_num) %>%
      mutate(HsCatedra = list(1:input[[paste0("horas_", WEEKDAY_MAPPING[weekday_num])]])) %>%
      unnest(cols = c(HsCatedra)) %>%
      ungroup()
    
    # Add holiday descriptions
    filas_finales <- filas_finales %>%
      left_join(feriados_nacionales %>% select(Fecha, Descripcion = Descripción), by = "Fecha")
    
    if (input$incluir_judios) {
      filas_finales <- filas_finales %>%
        left_join(feriados_judios %>% select(Fecha, DescripcionJudia = Descripción), by = "Fecha") %>%
        mutate(Tema = ifelse(!is.na(Descripcion), Descripcion, 
                             ifelse(!is.na(DescripcionJudia), DescripcionJudia, ""))) %>%
        select(-Descripcion, -DescripcionJudia)
    } else {
      filas_finales <- filas_finales %>%
        mutate(Tema = ifelse(!is.na(Descripcion), Descripcion, "")) %>%
        select(-Descripcion)
    }
    
    # Add class number, temporarily excluding non-empty "Tema" rows
    filas_finales <- filas_finales %>%
      arrange(Fecha, HsCatedra) %>%
      group_by(Fecha) %>%
      mutate(ClaseNumero = ifelse(Tema != "", NA, row_number())) %>%
      ungroup()
    
    # Sequentially number classes, skipping rows with NA in "Clase Nº"
    filas_finales <- filas_finales %>%
      mutate(ClaseNumero = ifelse(is.na(ClaseNumero), NA, cumsum(!is.na(ClaseNumero) & !duplicated(Fecha))))
    
    # Set "-" for rows with non-empty "Tema"
    filas_finales <- filas_finales %>%
      mutate(ClaseNumero = ifelse(Tema != "", "-", as.character(ClaseNumero)))
    
    # Insert winter break if excluded
    if (input$excluir_receso) {
      fila_receso <- data.frame(
        ClaseNumero = "-",
        Fecha = as.Date(input$receso_invernal[1]),
        weekday_num = NA,
        HsCatedra = NA,
        Tema = "RECESO INVERNAL",
        stringsAsFactors = FALSE
      )
      
      posicion_receso <- min(which(filas_finales$Fecha > as.Date(input$receso_invernal[2])))
      
      if (length(posicion_receso) > 0) {
        filas_finales <- bind_rows(
          filas_finales[1:(posicion_receso - 1), ],
          fila_receso,
          filas_finales[posicion_receso:nrow(filas_finales), ]
        )
      }
    }
    
    # Convert weekday numbers back to Spanish names for display
    filas_finales <- filas_finales %>%
      mutate(DiaSemana = if_else(!is.na(weekday_num), 
                                 WEEKDAY_MAPPING[weekday_num], 
                                 NA_character_)) %>%
      select(ClaseNumero, Fecha, DiaSemana, Tema)
    
    # Rename columns for final presentation
    names(filas_finales) <- c("Clase Nº", "Fecha", "Día de la Semana", "Tema")
    
    # Format date as text
    filas_finales$Fecha <- format(filas_finales$Fecha, "%d-%m-%Y")
    
    # Save and show results
    datos_resultantes(filas_finales)
    
    output$resultTable <- renderTable({
      filas_finales
    })
  })
  
  output$downloadExcel <- downloadHandler(
    filename = function() {
      paste0(input$titulo_archivo, "- Planificación.xlsx")
    },
    content = function(file) {
      write_xlsx(datos_resultantes(), file)
    }
  )
}

# Run the application 
shinyApp(ui = ui, server = server)