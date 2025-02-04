options(shiny.maxRequestSize = 300 * 1024^2)


library(shiny)
library(readxl)
library(dplyr)
library(lubridate)
library(openxlsx)

ui <- fluidPage(
  titlePanel("Saving Tracker Helper"),
  sidebarLayout(
    sidebarPanel(
      fileInput("file", "Upload XLSM File", accept = c(".xlsm")),
      dateInput('fia_start_date', 'FIA Start Date', format = 'yyyy-mm-dd'),
      dateInput('fia_end_date', 'FIA End Date', format = 'yyyy-mm-dd'),
      dateInput("start_date", "Contract Start Date", format = "yyyy-mm-dd"),
      dateInput("end_date", "Contract End Date", format = "yyyy-mm-dd"),
      actionButton("run_analysis", "Run Analysis", class = "btn-primary"),
      width = 3
    ),
    mainPanel(
      tableOutput("savings_table"),
      downloadButton("download_data", "Download Helper File")
    )
  )
)

server <- function(input, output) {
  # Reactive value to store calculations
  analysis_trigger <- reactiveVal(0)
  
  # Observe run analysis button
  observeEvent(input$run_analysis, {
    analysis_trigger(analysis_trigger() + 1)
  })
  
  data <- reactive({
    req(input$file)
    inFile <- input$file
    df <- read_excel(inFile$datapath, sheet = "POH Data", skip = 1)
    
    validate(
      need("Savings:" %in% colnames(df), "Column 'Savings:' not found in POH Data")
    )
    
    df
  })
  
  contract_days <- reactive({
    req(input$start_date, input$end_date)
    as.numeric(difftime(input$end_date, input$start_date, units = "days")) + 1
  })
  
  fia_days <- reactive({
    req(input$fia_start_date, input$fia_end_date)
    as.numeric(difftime(input$fia_end_date, input$fia_start_date, units = "days")) + 1
  })
  
  savings_data <- eventReactive(analysis_trigger(), {
    req(data(), contract_days())
    
    df <- data() %>% 
      group_by(FACILITY_NAME) %>% 
      summarise(fia_saving = sum(`Savings:`)) %>% 
      ungroup() %>% 
      mutate(
        contract_lifetime_saving = (fia_saving / fia_days()) * contract_days()
      ) %>%
      mutate(
        fia_saving = round(fia_saving, 0),
        contract_lifetime_saving = round(contract_lifetime_saving, 0)
      )
    
    # Add total row
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