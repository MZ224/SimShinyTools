#' Launch My Shiny App
#'
#' This function checks for required packages and runs the Shiny app.
#'
#' @export
Renamer <- function() {
  # List of required packages
  required_packages <- c("shiny", "readxl", "DT","openxlsx","dplyr","shinythemes")

  # Check which packages are not installed
  not_installed <- required_packages[!required_packages %in% installed.packages()[,"Package"]]

  # Install missing packages
  if(length(not_installed)) {
    install.packages(not_installed, dependencies = TRUE)
  }

  shiny::runApp(system.file("apps/Renamer", package = "SimShinyTools"))
}
