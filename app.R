library(shiny)
library(shinyjs)
library(readxl)
library(dplyr)
library(lubridate)
library(openxlsx)

options(shiny.maxRequestSize = 3000 * 1024^2)

ui <- fluidPage(
  useShinyjs(),
  titlePanel("Saving Tracker Helper"),
  sidebarLayout(
    sidebarPanel(
      fileInput("file", "Upload .xlsm File Only", accept = ".xlsm"),
      shinyjs::hidden(div(id = "file_error", style = "color: red;", "Only .xlsm files are accepted.")),
      shinyjs::hidden(div(id = "sheet_error", style = "color: red;", "Uploading File does not contain required Tab POH Data")),
      shinyjs::hidden(div(id = "colname_error", style = "color: red;", "Uploading File does not contain required column, Savings: in POH Data")),
      dateInput("start_date", "Contract Start Date", format = "yyyy-mm-dd"),
      dateInput("end_date", "Contract End Date", value = NULL, format = "yyyy-mm-dd"),
      shinyjs::hidden(div(id = "contract_date_error", style = "color: red;", "Contract end date cannot be earlier than start date.")),
      actionButton("run_analysis", "Run Analysis", class = "btn-primary"),
      width = 3
    ),
    mainPanel(
      tableOutput("savings_table"),
      shinyjs::disabled(downloadButton("download_data", "Download Helper File"))
    )
  )
)

server <- function(input, output, session) {
  analysis_trigger <- reactiveVal(0)
  last_input_state <- reactiveVal(NULL)
  
  # Track validation states
  invalid_file <- reactiveVal(FALSE)
  invalid_sheet <- reactiveVal(FALSE)
  invalid_column <- reactiveVal(FALSE)
 #invalid_fia_dates <- reactiveVal(FALSE)
  invalid_contract_dates <- reactiveVal(FALSE)
  
  # Sequential validation
  observeEvent(input$file, {
    # Reset validation states
    invalid_file(FALSE)
    invalid_sheet(FALSE)
    invalid_column(FALSE)
    
    # File validation
    if (!is.null(input$file)) {
      ext <- tolower(tools::file_ext(input$file$name))
      if (ext != "xlsm") {
        invalid_file(TRUE)
        return()  # Stop further checks if file validation fails
      }
    } else {
      invalid_file(FALSE)
      return()  # Stop if no file is uploaded
    }
    
    # Tab validation (only if file validation passes)
    sheets <- excel_sheets(input$file$datapath)
    if (!("POH Data" %in% sheets)) {
      invalid_sheet(TRUE)
      return()  # Stop further checks if sheet validation fails
    }
    
    # Column validation (only if file and sheet validation pass)
    df <- read_excel(input$file$datapath, sheet = "POH Data", skip = 1)
    if (!("Savings:" %in% colnames(df))) {
      invalid_column(TRUE)
      return()  # Stop further checks if column validation fails
    }
  })
  

  
  # Date validations
  observe({
    req(input$fia_start_date, input$fia_end_date)
    invalid_fia_dates(input$fia_end_date < input$fia_start_date)
  })
  
  observe({
    req(input$start_date, input$end_date)
    invalid_contract_dates(input$end_date < input$start_date)
  })
  
  # Show/hide error messages
  observe({
    shinyjs::toggle("file_error", condition = invalid_file())
    shinyjs::toggle("sheet_error", condition = invalid_sheet())
    shinyjs::toggle("colname_error", condition = invalid_column())
    #shinyjs::toggle("fia_date_error", condition = invalid_fia_dates())
    shinyjs::toggle("contract_date_error", condition = invalid_contract_dates())
  })
  
  # Disable run button and set tooltip
  observe({
    any_invalid <- invalid_file() || invalid_sheet() || invalid_column()|| invalid_contract_dates()
    shinyjs::toggleState("run_analysis", condition = !any_invalid)
    
    if (any_invalid) {
      errors <- c()
      if (invalid_file()) errors <- c(errors, "Invalid file type")
      if (invalid_sheet()) errors <- c(errors, "POH Data tab not found")
      if (invalid_column()) errors <- c(errors, "Column 'Savings:' not found")
      #if (invalid_fia_dates()) errors <- c(errors, "FIA end date error")
      if (invalid_contract_dates()) errors <- c(errors, "Contract date error")
      shinyjs::runjs(paste0(
        '$("#run_analysis").attr("title", "', paste(errors, collapse = "\\n"), '")'
      ))
    } else {
      shinyjs::runjs('$("#run_analysis").removeAttr("title")')
    }
  })
  
  # observeEvent(input$fia_start_date, {
  #   updateDateInput(session, "fia_end_date",
  #                   min = input$fia_start_date,
  #                   value = input$fia_start_date)
  # })
  # 
  # observeEvent(input$start_date, {
  #   updateDateInput(session, "end_date",
  #                   min = input$start_date,
  #                   value = input$start_date)
  # })
  
  observeEvent(input$run_analysis, {
    current_state <- list(
      file = input$file,
      #fia_start_date = input$fia_start_date,
      #fia_end_date = input$fia_end_date,
      start_date = input$start_date,
      end_date = input$end_date
    )
    
    if (!is.null(last_input_state()) && identical(current_state, last_input_state())) {
      shinyjs::disable("run_analysis")
      showNotification("Please don't repeat clicking without changing inputs.", type = "error")
      shinyjs::delay(3000, shinyjs::enable("run_analysis"))
      return()
    }
    
    last_input_state(current_state)
    analysis_trigger(analysis_trigger() + 1)
  })
  
  data <- reactive({
    req(input$file)
    inFile <- input$file
    
    validate(
      need(tools::file_ext(input$file$name) == "xlsm", "Only .xlsm files are accepted."),
      need("POH Data" %in% excel_sheets(inFile$datapath), "POH Data tab not found."),
      need("Savings:" %in% colnames(read_excel(inFile$datapath, sheet = "POH Data", skip = 1)), 
           "Column 'Savings:' not found.")
    )
    
    df = read_excel(inFile$datapath, sheet = "POH Data", skip = 1)
    df
  })
  
  contract_days <- reactive({
    req(input$start_date, input$end_date)
    as.numeric(difftime(input$end_date, input$start_date, units = "days")) + 1
  })
  
  fia_days <- reactive(365)
  
  savings_data <- eventReactive(analysis_trigger(), {
    req(data(), contract_days(), fia_days())
    
    df <- data() %>% 
      group_by(FACILITY_NAME) %>% 
      summarise(fia_saving = sum(`Savings:`)) %>% 
      ungroup() %>% 
      mutate(
        contract_lifetime_saving = (fia_saving / fia_days()) * contract_days(),
        across(c(fia_saving, contract_lifetime_saving), round, 0)
      )
    
    bind_rows(df, summarise(df, FACILITY_NAME = "Total", across(where(is.numeric), sum)))
  })
  
  output$savings_table <- renderTable(savings_data(), digits = 0)
  
  observe({
    toggleState("download_data", !is.null(savings_data()))
  })
  
  output$download_data <- downloadHandler(
    filename = function() paste0("contract_savings-", Sys.Date(), ".xlsx"),
    content = function(file) {
      wb <- createWorkbook()
      addWorksheet(wb, "Saving Tracker")
      writeData(wb, "Saving Tracker", savings_data())
      saveWorkbook(wb, file, overwrite = TRUE)
    }
  )
}
shinyApp(ui = ui, server = server)
