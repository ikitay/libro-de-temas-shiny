# Load required libraries
library(shiny)
library(readxl)
library(dplyr)
library(openxlsx)
library(lubridate)
library(tidyr)
library(shinyjs)

# Load data files
feriados_nacionales <- read_excel("datos.xlsx", sheet = "Feriados nacionales")
feriados_judios     <- read_excel("datos.xlsx", sheet = "No laborables judíos")
feriados_armenios   <- read_excel("datos.xlsx", sheet = "No laborables armenios")
feriados_musulmanes <- read_excel("datos.xlsx", sheet = "No laborables musulmanes")
cronogramas <- read_excel("datos.xlsx", sheet = "Cronograma por distrito")

# Ensure dates are in the correct format
# For holiday tables: explicit format since dates are stored as strings (dd-mm-yyyy)
# For cronogramas: across() converts all columns except "distrito" to Date
feriados_nacionales$Fecha <- as.Date(feriados_nacionales$Fecha, format = "%d-%m-%Y")
feriados_judios$Fecha <- as.Date(feriados_judios$Fecha, format = "%d-%m-%Y")
feriados_armenios$Fecha    <- as.Date(feriados_armenios$Fecha,    format = "%d-%m-%Y")
feriados_musulmanes$Fecha  <- as.Date(feriados_musulmanes$Fecha,  format = "%d-%m-%Y")
cronogramas <- cronogramas %>% mutate(across(-distrito, as.Date))

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
      
      # District selector: automatically defines the academic period and winter break dates
      
      selectInput("distrito", "Seleccione el distrito:",
                  choices = cronogramas$distrito,
                  selected = cronogramas %>% filter(distrito == "CABA")),
      tags$p("Las fechas generadas corresponderán al calendario oficial para el nivel medio del distrito seleccionado.", style = "font-size: 90%;"),
      
      tags$hr(),
      
      selectInput("calendario_especial", "Agregar asuetos de calendarios especiales (colectividades):",
                  choices = c("Ninguno", "Comunidad Judía", "Comunidad Armenia", "Comunidad Islámica"),
                  selected = "Ninguno"),
      tags$p("El calendario de cada institución puede ser ligeramente diferente.", style = "font-size: 90%;"),
      
      tags$hr(),
      
      # Class day selector and hours per day
      tags$label("Seleccione los días de clase:"),
      checkboxGroupInput("dias_seleccionados", 
                         label = NULL,
                         choices = WEEKDAY_MAPPING,
                         selected = NULL),
      
      # Inputs for class hours per day
      uiOutput("horasCatedraInputs"),
      
      tags$hr(),
      # Checkbox for excluding the winter break
      checkboxInput("excluir_receso", "Ocultar fechas del período de receso invernal.", value = TRUE),
      
      tags$hr(),
      # Button to generate the list
      actionButton("generate", strong("Generar planilla"), class = "btn-primary btn-block"),
      
      # Button to download the file (initially disabled)
      downloadButton("downloadExcel", strong("Descargar planilla (.xlsx)"), class = "btn-success btn-block"),
      
      # Conditionally show the title reminder text based on whether title is empty
      uiOutput("titleReminder"),
      
      tags$hr(),
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
  
  # Reactives: academic cycle and winter break dates derived from the selected district
    fechas_completas <- reactive({
      req(input$distrito)
      fila <- cronogramas[cronogramas$distrito == input$distrito, ]
      seq(fila$`inicio ciclo lectivo`, fila$`fin ciclo lectivo`, by = "day")
    })
    
    receso_fechas <- reactive({
      req(input$distrito)
      fila <- cronogramas[cronogramas$distrito == input$distrito, ]
      c(fila$`inicio receso`, fila$`fin receso`)
    })
  
  observeEvent(input$generate, {
    # Consolidated input validation
    if (is.null(input$dias_seleccionados) || length(input$dias_seleccionados) == 0) {
      showNotification("Por favor, seleccione los días de clase.", type = "error")
      return()
    }
    
    # Generate dates and exclude the specific period
    fechas_vector <- fechas_completas()
    
    # Exclude winter break dates if checkbox is checked
    fechas_excluidas <- if(input$excluir_receso) {
      seq(receso_fechas()[1], receso_fechas()[2], by = "day")
    } else {
      NULL
    }
    
    fechas_finales <- setdiff(fechas_vector, fechas_excluidas)
    fechas_finales <- as.Date(fechas_finales, origin = "1970-01-01")
    
    # Convert selected Spanish weekdays to lubridate numbers
    selected_weekday_numbers <- match(input$dias_seleccionados, WEEKDAY_MAPPING)
    
    # Create initial DataFrame with weekday numbers
    fechas_dias <- data.frame(
      Fecha = fechas_finales,
      weekday_num = wday(fechas_finales, week_start = 1),
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
    
    # Select special calendar data frame based on chosen community, and add holiday descriptions
    filas_finales <- filas_finales %>%
      left_join(feriados_nacionales %>% select(Fecha, Descripcion = Descripción), by = "Fecha")
    
    feriados_especiales <- switch(input$calendario_especial,
                                  "Comunidad Judía"    = feriados_judios,
                                  "Comunidad Armenia"  = feriados_armenios,
                                  "Comunidad Islámica" = feriados_musulmanes,
                                  NULL
    )
    
    if (!is.null(feriados_especiales)) {
      filas_finales <- filas_finales %>%
        left_join(feriados_especiales %>% select(Fecha, DescripcionEspecial = Descripción), by = "Fecha") %>%
        mutate(Tema = ifelse(!is.na(Descripcion), Descripcion,
                             ifelse(!is.na(DescripcionEspecial), DescripcionEspecial, ""))) %>%
        select(-Descripcion, -DescripcionEspecial)
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
        Fecha = receso_fechas()[1],
        weekday_num = NA,
        HsCatedra = NA,
        Tema = "RECESO INVERNAL",
        stringsAsFactors = FALSE
      )
      
      posicion_receso <- min(which(filas_finales$Fecha > receso_fechas()[2]))
      
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
    }, bordered = TRUE)
  })
  
  output$downloadExcel <- downloadHandler(
    filename = function() {
      paste0(input$titulo_archivo, " - Planificación diaria anual.xlsx")
    },
    content = function(file) {
      # Create workbook
      wb <- createWorkbook()
      addWorksheet(wb, "Planificación")
      
      # Define border style (thin borders on all sides)
      estilo_bordes <- createStyle(
        border = c("top", "bottom", "left", "right"),
        borderStyle = "thin"
      )
      
      # Define header style (bold + borders)
      estilo_encabezado <- createStyle(
        textDecoration = "bold",
        border = c("top", "bottom", "left", "right"),
        borderStyle = "thin"
      )
      
      # Write data to worksheet
      writeData(wb, "Planificación", datos_resultantes(), startRow = 1, startCol = 1)
      
      # Apply header style to first row
      addStyle(wb, "Planificación", estilo_encabezado, rows = 1, cols = 1:ncol(datos_resultantes()), gridExpand = TRUE, stack = TRUE)
      
      # Apply border style to all data cells (including header)
      addStyle(wb, "Planificación", estilo_bordes, rows = 1:(nrow(datos_resultantes()) + 1), cols = 1:ncol(datos_resultantes()), gridExpand = TRUE, stack = TRUE)
      
      # Save workbook
      saveWorkbook(wb, file, overwrite = TRUE)
    }
  )
}

# Run the application 
shinyApp(ui = ui, server = server)