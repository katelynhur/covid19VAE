################################################################################
# This script gathers all unique AE terms from multiple Excel files and sheets,
# and flags those that are present in a predefined administrative list or match
# specific patterns (e.g., ending with "test", "normal", or "negative").
#
################################################################################

################################################################################
# Set the current working directory
################################################################################
library(rstudioapi) # make sure you have it installed
base_dir <- dirname(getActiveDocumentContext()$path)
setwd(base_dir)

################################################################################
## Load required packages
################################################################################
if (!requireNamespace("pacman", quietly = TRUE)) install.packages("pacman")
pacman::p_load(readxl, dplyr, writexl)

################################################################################
# Main script
################################################################################
# Set the directory containing the Excel files
input_directory <- "./excel_output"

# Get a list of all Excel files in the directory
excel_files <- list.files(path = input_directory, pattern = "\\.xlsx$", full.names = TRUE)

# Read the "Admintrative_terms_to-be-removed.xlsx" file
# This file was manually created by manually curating the AE terms
admin_file <- "./Unique_AE_Terms/Admintrative_terms_to-be-removed.xlsx"
admin_sheets <- excel_sheets(admin_file)

# Initialize an empty vector to store administrative terms
admin_terms <- c()

# Loop through each sheet in the administrative file
for (sheet in admin_sheets) {
  admin_data <- read_excel(admin_file, sheet = sheet)
  
  # Check if the first column exists and collect unique terms
  if (ncol(admin_data) > 0) {
    admin_terms <- unique(c(admin_terms, admin_data[[1]]))
  }
}

# Initialize an empty vector to store all unique AE terms
all_ae_terms <- c()

# Loop through each Excel file
for (file in excel_files) {
  # Get all sheet names
  sheets <- excel_sheets(file)
  
  # Loop through each sheet in the current file
  for (sheet in sheets) {
    # Read the sheet and select the first column (AE)
    sheet_data <- read_excel(file, sheet = sheet)
    
    # Check if the first column exists and collect unique terms
    if (ncol(sheet_data) > 0) {
      ae_terms <- unique(sheet_data[[1]])
      all_ae_terms <- unique(c(all_ae_terms, ae_terms))
    }
  }
}

# Create a data frame from the unique AE terms
all_ae_terms_df <- data.frame(AE_Terms = all_ae_terms, stringsAsFactors = FALSE)

# Add a column indicating if the term is in the administrative list or matches specific patterns
all_ae_terms_df$Flag <- ifelse(
  all_ae_terms_df$AE_Terms %in% admin_terms | 
    grepl("test$", all_ae_terms_df$AE_Terms, ignore.case = TRUE) | 
    grepl("normal$", all_ae_terms_df$AE_Terms, ignore.case = TRUE) | 
    grepl("negative$", all_ae_terms_df$AE_Terms, ignore.case = TRUE),1, 0
)

# Save the list of unique AE terms as an Excel file
write_xlsx(all_ae_terms_df, "Unique_AE_Terms/Unique_AE_Terms_List.xlsx")

# Print message indicating completion
cat("AE terms from all Excel files and sheets have been collected and saved to 'Unique_AE_Terms/Unique_AE_Terms_List.xlsx'.\n")

# It is highly recommended to go through this resulting Excel file
# for further manual curation as needed. 
