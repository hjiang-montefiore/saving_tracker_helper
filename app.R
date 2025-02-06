options(shiny.maxRequestSize = 3000 * 1024^2)

library(shiny)
library(shinyjs)      # For enabling/disabling UI elements
library(readxl)
library(dplyr)
library(lubridate)
library(openxlsx)

ui <- fluidPage(
  useShinyjs(),  # Initialize shinyjs
  titlePanel("Saving Tracker Helper"),
  sidebarLayout(
    sidebarPanel(
      fileInput("file", "Upload .xlsm File Only", accept = ".xlsm"),
      dateInput("fia_start_date", "FIA Start Date", value = Sys.Date(), format = "yyyy-mm-dd"),
      dateInput("fia_end_date", "FIA End Date", value = NULL, format = "yyyy-mm-dd"),
      dateInput("start_date", "Contract Start Date", format = "yyyy-mm-dd"),
      dateInput("end_date", "Contract End Date", value = NULL, format = "yyyy-mm-dd"),
      actionButton("run_analysis", "Run Analysis", class = "btn-primary"),
      width = 3
    ),
    mainPanel(
      tableOutput("savings_table"),
      # Initially disable the download button
      shinyjs::disabled(downloadButton("download_data", "Download Helper File"))
    )
  )
)

server <- function(input, output, session) {
  # Reactive values to trigger analysis and store the last input state
  analysis_trigger <- reactiveVal(0)
  last_input_state <- reactiveVal(NULL)
  
  # Update FIA End Date so that it cannot be earlier than FIA Start Date
  observeEvent(input$fia_start_date, {
    updateDateInput(session, "fia_end_date",
                    min = input$fia_start_date,
                    value = input$fia_start_date)
  })
  
  # Update Contract End Date so that it cannot be earlier than Contract Start Date
  observeEvent(input$start_date, {
    updateDateInput(session, "end_date",
                    min = input$start_date,
                    value = input$start_date)
  })
  
  # Observe the Run Analysis button
  observeEvent(input$run_analysis, {
    # Create a list of current input values
    current_state <- list(
      file = input$file,
      fia_start_date = input$fia_start_date,
      fia_end_date = input$fia_end_date,
      start_date = input$start_date,
      end_date = input$end_date
    )
    
    # If the inputs have not changed since the last run, disable the button and notify
    if (!is.null(last_input_state()) && identical(current_state, last_input_state())) {
      shinyjs::disable("run_analysis")
      showNotification("Please don't repeat clicking the button without changing inputs.", type = "error")
      shinyjs::delay(3000, shinyjs::enable("run_analysis"))  # Re-enable after 3 seconds
      return()
    }
    
    # Update the stored input state and trigger the analysis
    last_input_state(current_state)
    analysis_trigger(analysis_trigger() + 1)
  })
  
  # Reactive expression for file upload and validation
  data <- reactive({
    req(input$file)
    inFile <- input$file
    
    # Validate file extension using the original file name (case-insensitive)
    validate(
      need(tolower(tools::file_ext(inFile$name)) == "xlsm", "Please upload a .xlsm file.")
    )
    
    # Validate that the required sheet exists
    validate(
      need("POH Data" %in% excel_sheets(inFile$datapath), "POH Data tab not found in the uploaded file.")
    )
    
    df <- read_excel(inFile$datapath, sheet = "POH Data", skip = 1)
    
    # Validate that the required column exists
    validate(
      need("Savings:" %in% colnames(df), "Column 'Savings:' not found in POH Data.")
    )
    
    df
  })
  
  # Calculate contract days (inclusive)
  contract_days <- reactive({
    req(input$start_date, input$end_date)
    as.numeric(difftime(input$end_date, input$start_date, units = "days")) + 1
  })
  
  # Calculate FIA investigation period days (inclusive)
  fia_days <- reactive({
    req(input$fia_start_date, input$fia_end_date)
    as.numeric(difftime(input$fia_end_date, input$fia_start_date, units = "days")) + 1
  })
  
  # Calculate savings data when analysis is triggered
  savings_data <- eventReactive(analysis_trigger(), {
    req(data(), contract_days())
    
    df <- data() %>% 
      group_by(FACILITY_NAME) %>% 
      summarise(fia_saving = sum(`Savings:`)) %>% 
      ungroup() %>% 
      mutate(
        contract_lifetime_saving = (fia_saving / fia_days()) * contract_days(),
        fia_saving = round(fia_saving, 0),
        contract_lifetime_saving = round(contract_lifetime_saving, 0)
      )
    
    # Add a total row
    total_row <- df %>% 
      summarise(
        FACILITY_NAME = "Total",
        fia_saving = sum(fia_saving),
        contract_lifetime_saving = sum(contract_lifetime_saving)
      )
    
    bind_rows(df, total_row) %>% 
      select(FACILITY_NAME, fia_saving, contract_lifetime_saving)
  })
  
  output$savings_table <- renderTable({
    savings_data()
  }, digits = 0)
  
  # Enable the download button only when savings_data is ready
  observe({
    if (!is.null(savings_data())) {
      shinyjs::enable("download_data")
    } else {
      shinyjs::disable("download_data")
    }
  })
  
  output$download_data <- downloadHandler(
    filename = function() {
      paste0("contract_savings-", Sys.Date(), ".xlsx")
    },
    content = function(file) {
      req(input$file, savings_data())
      wb <- createWorkbook()
      addWorksheet(wb, "Saving Tracker")
      writeData(wb, "Saving Tracker", savings_data())
      saveWorkbook(wb, file, overwrite = TRUE)
    }
  )
}

shinyApp(ui = ui, server = server)
