################################################################################
# This script trims AE terms from all sheets of all Excel files in ./excel_output
# using two curated lists (tab-delimited) under ./Unique_AE_Terms.
# It (1) checks for conflicts between the curated lists,
# (2) filters rows where the FIRST column matches any flagged term (case-insensitive),
# (3) writes trimmed Excel files to ./trimmed_excel_output,
# and (4) saves a summary report as an Excel file.
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
pacman::p_load(readxl, dplyr, writexl, readr)

################################################################################
# Main script
################################################################################

# ------------------------------------------------------------------------------
# Paths and I/O setup
# ------------------------------------------------------------------------------
input_directory  <- "./excel_output"
output_directory <- "./trimmed_excel_output"
dir.create(output_directory, showWarnings = FALSE)

# Curated term files (tab-delimited): AE_Term \t Flag
file_curated_1 <- "Unique_AE_Terms/Unique_AE_Terms_List_1208.txt"
file_curated_2 <- "Unique_AE_Terms/Supplemental File S2_AEs_From-Frontiers_Revised_1208.txt"

# ------------------------------------------------------------------------------
# 1) Load curated AE terms (two files) and normalize for case-insensitive matching
# ------------------------------------------------------------------------------
ae_terms_1 <- read_delim(file_curated_1, delim = "\t",
                         col_names = c("AE_Term", "Flag"), show_col_types = FALSE)
ae_terms_2 <- read_delim(file_curated_2, delim = "\t",
                         col_names = c("AE_Term", "Flag"), show_col_types = FALSE)

ae_terms_1 <- ae_terms_1 %>%
  mutate(AE_Term_lower = tolower(AE_Term))

ae_terms_2 <- ae_terms_2 %>%
  mutate(AE_Term_lower = tolower(AE_Term))

# ------------------------------------------------------------------------------
# 2) Conflict check between curated files (same term, different Flag)
# ------------------------------------------------------------------------------
conflicts <- merge(ae_terms_1, ae_terms_2,
                   by = "AE_Term_lower", suffixes = c("_1", "_2")) %>%
  filter(Flag_1 != Flag_2) %>%
  select(AE_Term_1, Flag_1, AE_Term_2, Flag_2)

if (nrow(conflicts) > 0) {
  cat("Conflicts found between curated files (same term with different Flag):\n")
  print(conflicts)
} else {
  cat("No conflicts found between the two curated files.\n")
}

# Terms to trim: any term flagged (Flag == 1) in either file
terms_to_trim <- unique(c(
  ae_terms_1$AE_Term_lower[ae_terms_1$Flag == 1],
  ae_terms_2$AE_Term_lower[ae_terms_2$Flag == 1]
))

# ------------------------------------------------------------------------------
# 3) Process all Excel files and sheets under ./excel_output
# ------------------------------------------------------------------------------
excel_files <- list.files(path = input_directory, pattern = "\\.xlsx$", full.names = TRUE)
summary_stats <- dplyr::tibble()

for (file in excel_files) {
  sheets <- excel_sheets(file)
  processed_sheets <- list()
  
  for (sheet in sheets) {
    sheet_data <- read_excel(file, sheet = sheet)
    
    if (ncol(sheet_data) > 0) {
      original_count <- nrow(sheet_data)
      
      # Use FIRST column as AE term column (case-insensitive filter)
      data_filtered <- sheet_data %>%
        filter(!tolower(dplyr::pull(., 1)) %in% terms_to_trim)
      
      processed_sheets[[sheet]] <- data_filtered
      
      # Update summary
      summary_stats <- bind_rows(
        summary_stats,
        tibble(
          File            = basename(file),
          Sheet           = sheet,
          Total_Terms     = original_count,
          Trimmed_Terms   = original_count - nrow(data_filtered),
          Remaining_Terms = nrow(data_filtered)
        )
      )
    }
  }
  
  # Save processed workbook with same name into the output directory
  write_xlsx(processed_sheets, file.path(output_directory, basename(file)))
}

# ------------------------------------------------------------------------------
# 4) Summary report
# ------------------------------------------------------------------------------
cat("\nSummary Report:\n")
cat("Total rows in curated files (not unique): ",
    nrow(ae_terms_1) + nrow(ae_terms_2), "\n", sep = "")
cat("Unique terms marked for trimming: ", length(terms_to_trim), "\n\n", sep = "")

summary_output <- summary_stats %>%
  select(File, Sheet, Total_Terms, Trimmed_Terms, Remaining_Terms)

print(summary_output)

# Save summary as an Excel file
summary_path <- "Unique_AE_Terms/Unique_AE-term-trimming_summary_report_1208.xlsx"
write_xlsx(summary_output, path = summary_path)

cat("Trimmed files saved to '", output_directory, "'.\n", sep = "")
cat("Summary report saved to '", summary_path, "'.\n", sep = "")
