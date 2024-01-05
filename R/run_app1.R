#' Launch My Shiny App
#'
#' This function checks for required packages and runs the Shiny app.
#'
#' @export
PKplot <- function() {
  # List of required packages
  required_packages <- c("shiny", "openxlsx", "ggplot2","xml2","dplyr","purrr","tidyr","shinyjs","shinyBS","shinydashboard")

  # Check which packages are not installed
  not_installed <- required_packages[!required_packages %in% installed.packages()[,"Package"]]

  # Install missing packages
  if(length(not_installed)) {
    install.packages(not_installed, dependencies = TRUE)
  }

  shiny::runApp(system.file("apps/PKplot", package = "SimShinyTools"))
}

