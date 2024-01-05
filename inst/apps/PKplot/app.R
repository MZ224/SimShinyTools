# Import necessary libraries
library(shiny)
library(openxlsx)
library(ggplot2)
library(xml2)
library(dplyr)
library(purrr)
library(tidyr)
library(shinyjs)
library(shinyBS)
library(shinydashboard)
options(scipen=999)

# Set the graphics engine to AGG
options("device"="agg")

# Define UI
ui <- fluidPage(
  shinyjs::useShinyjs(),
  titlePanel("Excel and XML Data Visualisation"),
  
  sidebarLayout(
    sidebarPanel(
      # Action Buttons
      wellPanel(
        actionButton("about_button", "About"),
        actionButton("plot_button", "Plot"),
        actionButton("save_plot", "Save Plot"),
      ),
      
      bsCollapse(
        bsCollapsePanel(
          "Export Settings",
          textInput("jpeg_name", "File Name:", paste0("saved_plot_", Sys.Date())),
          numericInput("jpeg_width", "JPEG Width:", 1500, min = 1),
          numericInput("jpeg_height", "JPEG Height:", 1100, min = 1),
          numericInput("jpeg_dpi", "JPEG DPI:", 300, min = 1)
        )
      ),
      bsCollapse(
        id = "axisSettings",
        bsCollapsePanel(
          title = "Axis Settings",
          textInput("time_unit", "Time Unit:", "hour"),
          textInput("conc_unit", "Concentration Unit:", "ng/mL"),
          checkboxInput("user_x_axis", "Use user specified X-axis", FALSE),
          numericInput("x_break","X-axis Break:",value = 12, min = 1E-6),
          numericInput("x_min", "X-axis Min:", value = 0, min = 0),
          numericInput("x_max", "X-axis Max:", value = 72),
          checkboxInput("user_y_axis", "Use user specified Y-axis", FALSE),
          numericInput("y_min", "Y-axis Min:", value = 0, min = 0),
          numericInput("y_max", "Y-axis Max:", value = 100, min = 0)
        )
      ),
      
      # Directory and File Selection for Excel
      wellPanel(
      h4("Excel File Selection"),
      textInput("dir", "Directory:", value = getwd()),
      actionButton("update", "Update Excel File List"),
      selectInput("xlsx_files", "Choose an .xlsx file:", choices = c("Not selected")),
      checkboxInput("ddi_profile", "Profile with DDI", FALSE),
      ),
      
      # Directory and File Selection for XML
      wellPanel(h4("XML File Selection"),
      checkboxInput("xml_overlay", "Enable XML Overlay", TRUE),
      textInput("dir_xml", "Directory:", value = getwd()),
      actionButton("update_xml", "Update XML File List"),
      selectInput("xml_files", "Choose an .xml file:", choices = c("Not selected")),
      ),

      
      # Plot Settings
      wellPanel(h4("Plot Settings"),
      sliderInput("line_size", "Line Size:", min = 0.5, max = 3, value = 1, step = 0.1),
      textInput("line_color", "Mean Line Color (1st DV):", value = "royalblue"),
      textInput("line_color_ddi", "Mean Line Color (2nd DV):", value = "#EF8A62"),
      textInput("dashed_color", "Dashed Line Color (1st DV):", value = "#67A9CF"),
      textInput("dashed_color_ddi", "Dashed Line Color (2nd DV):", value = "#EF8A62"),
      textInput("shade_color", "Shade Color (1st DV):", value = "#67A9CF"),
      textInput("shade_color_ddi", "Shade Color (2nd DV):", value = "#EF8A62"),
      radioButtons("range_display", "Range Display:", choices = c("None", "Ribbon", "Dashed Lines"), selected = "Ribbon"),
      checkboxInput("log_scale", "Log Scale for Y-axis", FALSE),
      numericInput("font_size", "Font Size:", 11,min = 0.1),
      ),
      
      # Data Point Settings
      wellPanel(h4("Data Point Settings"),
      selectInput("choices", 
                  "Select DVs", 
                  choices = NULL,
                  multiple = TRUE),
      textOutput("selected"),
      numericInput("point_size", "Point Size:", 2),
      textInput("point_color", "Point Color (1st DV):", value = "royalblue"),
      textInput("point_color_ddi", "Point Color (2nd DV):", value = "#EF8A62"),
      sliderInput("point_alpha", "Point Transparency:", min = 0, max = 1, value = 1, step = 0.1),
      numericInput("data_scalar", "Data scalar:", value = 1, min = 0),
      ),
      
    ),
    
    mainPanel(
      plotOutput("plot"),
      uiOutput("warning_message")
    )
  )
)

# Define server logic
server <- function(input, output, session) {
  
  observeEvent(input$about_button, {
    showModal(modalDialog(
      title = "About This App",
      "This Shiny App was developed by Mian Zhang.",
      "Version: 1.0.0",
      "For more information, contact me at: zhangmian.cpu@gmail.com",
      easyClose = TRUE, 
      footer = tagList(
        tags$button("Close", type = "button", class = "btn btn-default", `data-dismiss` = "modal")
      )
    ))
    })
  
  observe({
    if (!input$xml_overlay) {
      updateTextInput(session, "dir_xml", value = getwd())
      shinyjs::disable("update_xml")
      shinyjs::disable("dir_xml")
      shinyjs::disable("xml_files")
      shinyjs::disable("choices")
      shinyjs::disable("point_size")
      shinyjs::disable("point_color")
      shinyjs::disable("point_color_ddi")
      shinyjs::disable("point_alpha")
      
    } else {
      shinyjs::enable("update_xml")
      shinyjs::enable("dir_xml")
      shinyjs::enable("xml_files")
      shinyjs::enable("choices")
      shinyjs::enable("point_size")
      shinyjs::enable("point_color")
      shinyjs::enable("point_color_ddi")
      shinyjs::enable("point_alpha")
    }
  })
  
  observe({
    if (!input$user_x_axis) {
      shinyjs::disable("x_min")
      shinyjs::disable("x_max")
      shinyjs::disable("x_break")
    } else {
      shinyjs::enable("x_min")
      shinyjs::enable("x_max")
      shinyjs::enable("x_break")
    }
  }) 
  
  observe({
    if (!input$user_y_axis) {
      shinyjs::disable("y_min")
      shinyjs::disable("y_max")
      
    } else {
      shinyjs::enable("y_min")
      shinyjs::enable("y_max")
    }
  }) 
  
  observe({
    warning_xlsx <- input$xlsx_files == "Not selected"
    warning_xml <- input$xml_overlay && input$xml_files == "Not selected"
    
    # Check if the selected Excel file has a "Summary" sheet
    valid_sheet <- TRUE  # assume the sheet is valid initially
    if (input$xlsx_files != "Not selected") {
      sheets <- openxlsx::getSheetNames(input$xlsx_files)
      valid_sheet <- ("Summary" %in% sheets)  # becomes FALSE if "Summary" isn't one of the sheets
    }
    
    if (warning_xlsx || warning_xml || !valid_sheet) {
      shinyjs::disable("plot_button")
      shinyjs::disable("save_plot")
    } else {
      shinyjs::enable("plot_button")
      shinyjs::enable("save_plot")
    }
  })
  
  
  observeEvent(input$update, {
    xlsx_files_full <- list.files(path = input$dir, pattern = "\\.xlsx$", full.names = TRUE)
    xlsx_files_names <- basename(xlsx_files_full)  # Extract just the filenames
    choices <- setNames(xlsx_files_full, xlsx_files_names)  # Named list where names are filenames and values are full paths
    updateSelectInput(session, "xlsx_files", choices = choices)
  })
  
  observeEvent(input$update_xml, {
    xml_files_full <- list.files(path = input$dir_xml, pattern = "\\.xml$", full.names = TRUE)
    xml_files_names <- basename(xml_files_full)  # Extract just the filenames
    choices <- setNames(xml_files_full, xml_files_names)  # Named list as before
    updateSelectInput(session, "xml_files", choices = choices)
  })
  
  observeEvent(input$log_scale, {
    if(input$log_scale) {
      # If checkbox is checked
      updateNumericInput(session, "y_min", value = 1, min = 1E-10)
    } else {
      # If checkbox is not checked (reset to original values)
      updateNumericInput(session, "y_min", value = 0, min = 0)
    }
  })
  
  
  
  inhibitor <- reactive({
    if (!(input$xlsx_files == "Not selected")) {
      
      # Check if "Summary" sheet exists in the Excel file
      sheets <- openxlsx::getSheetNames(input$xlsx_files)
      
      if ("Summary" %in% sheets) {
        # If the "Summary" sheet exists, then read it
        summary_sheet <- read.xlsx(input$xlsx_files, sheet = "Summary", rowNames = FALSE, colNames = FALSE)
        
        if ("Inhibitor 1" %in% summary_sheet[, 7]) {
          return("Yes")
        } else {
          return("No")
        }
      } else {
        # If the "Summary" sheet doesn't exist, return a message
        return("No 'Summary' sheet found in the selected file")
      }
    } else {
      return("No file has been selected")
    }
  })
  
  observe({
    if (inhibitor() == "Yes") {
      shinyjs::enable("ddi_profile")
    } else {
      runjs("$('#ddi_profile').prop('checked', false);")  
      shinyjs::disable("ddi_profile")
    }
  })
  
  
  output$warning_message <- renderUI({
    warning <- NULL
    if (input$xlsx_files == "Not selected") {
      warning <- paste(warning, "Please select an .xlsx file.\n")
    } else {
      # Check if the selected Excel file has a "Summary" sheet
      sheets <- openxlsx::getSheetNames(input$xlsx_files)
      if (!("Summary" %in% sheets)) {
        warning <- paste(warning, "The selected .xlsx file does not have a 'Summary' sheet. It might not be a valid output file.\n")
      }
    }
    if (input$xml_overlay && input$xml_files == "Not selected") {
      warning <- paste(warning,"Please select an .xml file.")
    }
    if (!is.null(warning)) {
      tags$div(class = "alert alert-warning", warning)
    }
  })
  


  output$inhibitor <- renderText({
    inhibitor()
  })

  event_data <- reactive({
    if (input$xml_overlay && !(input$xml_files == "Not selected")) {
      xml_data <- read_xml(input$xml_files)
      event_data <- xml_data %>%
        xml_find_all("//subject") %>%
        map_df(~{
          subject_id <- xml_attr(.x, "ID")
          events <- .x %>% xml_find_all("event")
          map_df(events, ~{
            data.frame(
              SubjectID = subject_id,
              FormulationType = xml_attr(.x, "FormulationType"),
              nDoseUnits = xml_attr(.x, "nDoseUnits"),
              nAdminRoute = xml_attr(.x, "nAdminRoute"),
              Compound = xml_attr(.x, "Compound"),
              PeriodID = xml_attr(.x, "PeriodID"),
              dDuration = xml_attr(.x, "dDuration"),
              dWeighting = xml_attr(.x, "dWeighting"),
              dDose = xml_attr(.x, "dDose"),
              UserDVID = xml_attr(.x, "UserDVID"),
              dDependantVariable = xml_attr(.x, "dDependantVariable"),
              dTime = xml_attr(.x, "dTime"),
              stringsAsFactors = FALSE
            )
          })
        })
      event_data <- event_data %>% filter(dDependantVariable != "-1")
      return(event_data)
    }
  })
  
  observe({
    if (input$xml_overlay && !(input$xml_files == "Not selected")) {
      df <- event_data()
      DVs <- unique(df$UserDVID)
      updateSelectInput(session, "choices", choices = DVs, selected = DVs[1])
    } 
  })

    plot_data <- eventReactive(input$plot_button, 
                               {
        df <- read.xlsx(input$xlsx_files, sheet = "Conc Profiles CSys(CPlasma)", rowNames = FALSE, colNames = FALSE,skipEmptyRows = FALSE,skipEmptyCols = FALSE)
        ##df <- read.xlsx("repo-160mg-study10-rif-3a4-37-pgp.xlsx", sheet = "Conc Profiles CSys(CPlasma)", rowNames = FALSE, colNames = FALSE,skipEmptyRows = FALSE,skipEmptyCols = FALSE)
        selected_DV <- input$choices
        
        if (inhibitor() == "Yes") {
          time <- df[31, -c(1:3)]
          
          mean_no_interaction <- df[32, -c(1:3)]
          upper_no_interaction <- df[34, -c(1:3)]
          lower_no_interaction <- df[33, -c(1:3)]
          
          mean_interaction <- df[35, -c(1:3)]
          upper_interaction <- df[37, -c(1:3)]
          lower_interaction <- df[36, -c(1:3)]
          
          df_no_interaction <- data.frame(
            Time = as.numeric(time),
            Mean = as.numeric(mean_no_interaction),
            Upper = as.numeric(upper_no_interaction),
            Lower = as.numeric(lower_no_interaction)
          )
          
          df_interaction <- data.frame(
            Time = as.numeric(time),
            Mean = as.numeric(mean_interaction),
            Upper = as.numeric(upper_interaction),
            Lower = as.numeric(lower_interaction)
          )
          
          p <- ggplot() +
            theme(panel.background = element_rect(fill = "white", colour = "#666666", linewidth = 0.5, linetype = "solid"),
                  panel.grid.major = element_line(linewidth = 0.1, linetype = 'dashed', colour = "lightgrey"),
                  panel.grid.minor = element_line(linewidth = 0.1, linetype = 'dashed', colour = "lightgrey"),
                  axis.text = element_text(size = input$font_size),
                  axis.title = element_text(size = input$font_size + 2),
                  legend.position = "none") + 
            labs(x = paste("Time (", input$time_unit, ")", sep = ""),
                 y = paste("Concentration (", input$conc_unit, ")", sep = ""))
          
          # No Interaction profile
          p <- p + geom_line(data = df_no_interaction, aes(x = Time, y = Mean), linewidth = input$line_size, color = input$line_color)
          
          if (input$ddi_profile) 
            {
            # Interaction profile
            p <- p + geom_line(data = df_interaction, aes(x = Time, y = Mean), linewidth = input$line_size, color = input$line_color_ddi)
            
            }
          

          if (input$range_display == "Ribbon") {
            p <- p + geom_ribbon(data = df_no_interaction,aes(x = Time, ymin = Lower, ymax = Upper), alpha = 0.2, fill = input$shade_color)
          } else if (input$range_display == "Dashed Lines") {
            p <- p + geom_line(data = df_no_interaction, aes(x = Time, y = Lower), linewidth = input$line_size, linetype = "dashed", color = input$dashed_color_ddi)
            p <- p + geom_line(data = df_no_interaction, aes(x = Time, y = Upper), linewidth = input$line_size, linetype = "dashed", color = input$dashed_color_ddi)
          }
          
          if (input$ddi_profile) 
          {
            if (input$range_display == "Ribbon") {
              p <- p + geom_ribbon(data = df_no_interaction,aes(x = Time, ymin = Lower, ymax = Upper), alpha = 0.2, fill = input$shade_color)
              p <- p + geom_ribbon(data = df_interaction,aes(x=Time, ymin = Lower, ymax = Upper), alpha = 0.2, fill = input$shade_color_ddi)
            } else if (input$range_display == "Dashed Lines") {
              p <- p + geom_line(data = df_no_interaction, aes(x = Time, y = Lower), linewidth = input$line_size, linetype = "dashed", color = input$dashed_color)
              p <- p + geom_line(data = df_no_interaction, aes(x = Time, y = Upper), linewidth = input$line_size, linetype = "dashed", color = input$dashed_color)
              p <- p + geom_line(data = df_interaction, aes(x = Time, y = Lower), linewidth = input$line_size, linetype = "dashed", color = input$dashed_color_ddi)
              p <- p + geom_line(data = df_interaction, aes(x = Time, y = Upper), linewidth = input$line_size, linetype = "dashed", color = input$dashed_color_ddi)
            }
          }
          
          if (input$xml_overlay && !is.null(event_data()) && !is.null(selected_DV)) {
            
            p <- p + geom_point(data = event_data()%>%filter(UserDVID==selected_DV), aes(x = as.numeric(dTime), y = as.numeric(dDependantVariable)*input$data_scalar,colour=UserDVID),
                                size = input$point_size, alpha = input$point_alpha)+
              scale_colour_manual(values = c(input$point_color,input$point_color_ddi))
          }
          
          
          if (input$log_scale) {
            df_no_interaction <- df_no_interaction %>%
              filter(Time > 0, Mean > 0, Upper > 0, Lower > 0)
            df_interaction <- df_no_interaction %>%
              filter(Time > 0, Mean > 0, Upper > 0, Lower > 0)
            p <- p + scale_y_log10()
          }
          
        } 
        
        if (!input$ddi_profile & inhibitor() == "No"){
        
          df <- read.xlsx(input$xlsx_files, sheet = "Conc Profiles CSys(CPlasma)", rowNames = FALSE, colNames = FALSE,skipEmptyRows = FALSE,skipEmptyCols = FALSE)
          
          ##df <- read.xlsx("repotrectinib-rel-ba-capsule-160mg-sd.xlsx", sheet = "Conc Profiles CSys(CPlasma)", rowNames = FALSE, colNames = FALSE,skipEmptyRows = FALSE,skipEmptyCols = FALSE)
          
          time <- df[29, -c(1:3)]
          mean_concentration <- df[30, -c(1:3)]
          upper_range <- df[32, -c(1:3)]
          lower_range <- df[31, -c(1:3)]
          
          df_plot <- data.frame(
            Time = as.numeric(time),
            Mean = as.numeric(mean_concentration),
            Upper = as.numeric(upper_range),
            Lower = as.numeric(lower_range)
          )
          
          p <- ggplot(df_plot, aes(x = Time)) +
            theme(panel.background = element_rect(fill = "white", colour = "#666666", linewidth = 0.5, linetype = "solid"),
                  panel.grid.major = element_line(linewidth = 0.1, linetype = 'dashed', colour = "lightgrey"),
                  panel.grid.minor = element_line(linewidth = 0.1, linetype = 'dashed', colour = "lightgrey"),
                  axis.text = element_text(size = input$font_size),
                  axis.title = element_text(size = input$font_size + 2),
                  legend.position = "none") + 
            labs(x = paste("Time (", input$time_unit, ")", sep = ""),
                 y = paste("Concentration (", input$conc_unit, ")", sep = "")) +
            geom_line(aes(y = Mean), size = input$line_size, color = input$line_color)
          
          if (input$range_display == "Ribbon") {
            p <- p + geom_ribbon(aes(ymin = Lower, ymax = Upper), alpha = 0.2, fill = input$shade_color)
          } else if (input$range_display == "Dashed Lines") {
            p <- p + geom_line(aes(y = Upper), linewidth = input$line_size, linetype = "dashed", color = input$dashed_color)
            p <- p + geom_line(aes(y = Lower), linewidth = input$line_size, linetype = "dashed", color = input$dashed_color)
          }
          
          if (input$xml_overlay && !is.null(event_data())) {
            
            p <- p + geom_point(data = event_data(), aes(x = as.numeric(dTime), y = as.numeric(dDependantVariable)*input$data_scalar),
                                size = input$point_size, color = input$point_color, alpha = input$point_alpha)
          }
          
          if (input$log_scale) {
            df_plot <- df_plot %>%
              filter(Time > 0, Mean > 0, Upper > 0, Lower > 0)
            p <- p + scale_y_log10()
          }
          
        }
        
          if (input$user_x_axis&&(!input$user_y_axis)) {
           p <- p + coord_cartesian(xlim = c(input$x_min, input$x_max))+
             scale_x_continuous(breaks = seq(input$x_min, input$x_max, by = input$x_break)) 
          }
        
          if (input$user_y_axis&&input$user_x_axis) {
          p <- p + coord_cartesian(xlim = c(input$x_min, input$x_max),ylim = c(input$y_min, input$y_max))+
            scale_x_continuous(breaks = seq(input$x_min, input$x_max, by = input$x_break)) 
          }
        
          if (input$user_y_axis&&(!input$user_x_axis)) {
          p <- p + coord_cartesian(ylim = c(input$y_min, input$y_max))
          }
        
        
        
        return(p)
        
        })
    
      
  output$plot <- renderPlot({
    plot_data()
  })
  
  observeEvent(input$save_plot, {
    # Extract directory from the xlsx file path
    save_dir <- dirname(input$xlsx_files)
    
    # Combine directory with desired filename
    full_path <- file.path(save_dir, paste0(input$jpeg_name, ".jpeg"))
    
    ggsave(filename = full_path, 
           plot = plot_data(), 
           width = input$jpeg_width / input$jpeg_dpi, 
           height = input$jpeg_height / input$jpeg_dpi, 
           dpi = input$jpeg_dpi)
  })
  
}

# Run the application 
shinyApp(ui = ui, server = server)
