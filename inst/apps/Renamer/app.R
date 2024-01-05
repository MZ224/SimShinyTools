library(shiny)
library(readxl)
library(openxlsx)
library(dplyr)
library(shinythemes)
library(DT)

ui <- fluidPage(
  theme = shinytheme("sandstone"), # Bootstrap theme for a nicer aesthetic

  tags$head(
    tags$style(HTML("
      body {padding-top: 50px;}
      .navbar {background-color: #f0ad4e; border-color: #eea236;}
      .navbar-brand {color: #ffffff;}
      .navbar-brand:hover {color: #ffffff;}
      .btn {margin-left: 5px; margin-right: 5px; vertical-align: middle;}
      .shiny-input-container {display: inline-block; vertical-align: middle; width: auto; margin-right: 5px;}
      .instructions {background-color: #f7f7f7; padding: 15px; border-radius: 5px; margin-bottom: 20px;}
    "))
  ),

  navbarPage("Excel File Checker", id="nav",
             tabPanel("File Operations",
                      fluidRow(
                        column(12,
                               div(class = "instructions",
                                   tags$h4("Instructions:"),
                                   tags$p("Enter the directory path where Excel files are located. Use 'Show Excel Files' to list files, 'Save File List' to export the list, and 'Rename Files' to rename files based on Simcyp outputs.")
                               )
                        )
                      ),
                      fluidRow(
                        column(12,
                               div(
                                 style = "display: flex; justify-content: space-between; align-items: center;",
                                 textInput("dirInput", "Enter Directory:", value = getwd()),
                                 actionButton("showBtn", "Show Excel Files", icon = icon("eye"), class = "btn-primary"),
                                 actionButton("saveBtn", "Save File List", icon = icon("save"), class = "btn-primary"),
                                 actionButton("renameBtn", "Rename Files", icon = icon("edit"), class = "btn-primary")
                               )
                        )
                      ),
                      fluidRow(
                        column(12,
                               DTOutput("table")
                        )
                      )
             )
  )
)

server <- function(input, output) {
  # Variable to store the results
  results <- reactiveVal(data.frame())

  observeEvent(input$showBtn, {
    req(input$dirInput)
    files <- list.files(path = input$dirInput, pattern = "\\.xlsx$", full.names = TRUE)

    resultData <- lapply(files, function(file) {
      sheets <- excel_sheets(file)
      list(
        FileName = basename(file),
        IsSimcypOutput = "Summary" %in% sheets
      )
    })

    results(do.call(rbind, resultData))

    output$table <- renderDT({
      datatable(results(), options = list(pageLength = 10))
    })
  })

  observeEvent(input$saveBtn, {
    req(input$dirInput)
    req(nrow(results()) > 0)  # Ensure there is data to save

    # Debugging: Print the data to be saved
    file_list<-as.data.frame(results())%>%
      filter(IsSimcypOutput==TRUE)
    savePath <- file.path(input$dirInput, "file_list.xlsx")
    write.xlsx(file_list, savePath)
    showNotification("File list saved!", type = "message")
  })

  observeEvent(input$renameBtn, {
    req(input$dirInput)
    files <- list.files(path = input$dirInput, pattern = "\\.xlsx$", full.names = TRUE)

    for(file in files) {
      # Read the "Summary" sheet to check for Simcyp output
      # Ensure the summary sheet and the specific cell for workspace name exist
      if("Summary" %in% excel_sheets(file)) {
        summary <- read.xlsx(file, sheet = "Summary")
        # Assuming the workspace name is in a specific cell, e.g., B2
        # Update this according to your file structure
        workspaceName <-summary%>%filter(Simcyp.Population.Based.Simulator=="Workspace")%>%select(X2)

        if(!is.na(workspaceName) && workspaceName != "") {
          newFileName <- paste0(input$dirInput, "/", workspaceName, ".xlsx")
          file.rename(file, newFileName)
        }
      }
    }

    showNotification("Files renamed successfully", type = "message")
  })


}

shinyApp(ui, server)
