######################################################
## Set the current working directory
######################################################
library(rstudioapi) # make sure you have it installed
current_path <- getActiveDocumentContext()$path 
setwd(dirname(current_path ))
base_dir = dirname(current_path)

# Load necessary libraries
library(readxl)
library(dplyr)
library(writexl)
library(readr)  # Added for reading text files

# Create output directory if it doesn't exist
dir.create("trimmed_excel_output", showWarnings = FALSE)

# 1. Load AE terms from both files
# Unique_AE_Terms_List_1208.txt file contains 
# curated information for trimming
# Update as needed.
ae_terms_1 <- read_delim("Unique_AE_Terms/Unique_AE_Terms_List_1208.txt", 
                        delim = "\t", 
                        col_names = c("AE_Term", "Flag"))

# Supplemental File S2_AEs_From-Frontiers_Revised_1208.txt
# This file should contain the list of trimmed AE terms from the previous publication
# However, due to the chagnes in the trimmimg strategy, 
# manual review and editing is needed, with the resulting version is to be saved to this file
ae_terms_2 <- read_delim("Unique_AE_Terms/Supplemental File S2_AEs_From-Frontiers_Revised_1208.txt", 
                        delim = "\t", 
                        col_names = c("AE_Term", "Flag"))

# Convert terms to lowercase for case-insensitive comparison
ae_terms_1$AE_Term_lower <- tolower(ae_terms_1$AE_Term)
ae_terms_2$AE_Term_lower <- tolower(ae_terms_2$AE_Term)

# 2. Check for conflicts
conflicts <- merge(ae_terms_1, ae_terms_2, 
                  by = "AE_Term_lower", 
                  suffixes = c("_1", "_2")) %>%
  filter(Flag_1 != Flag_2) %>%
  select(AE_Term_1, Flag_1, AE_Term_2, Flag_2)

if (nrow(conflicts) > 0) {
  cat("Conflicts found:\n")
  print(conflicts)
} else {
  cat("No conflicts found between the two files.\n")
}

# Combine terms to be trimmed (Flag = 1)
terms_to_trim <- unique(c(
  ae_terms_1$AE_Term_lower[ae_terms_1$Flag == 1],
  ae_terms_2$AE_Term_lower[ae_terms_2$Flag == 1]
))


# 3. Process Excel files
excel_files <- list.files(path = "excel_output", 
                         pattern = "\\.xlsx$", 
                         full.names = TRUE)

# Initialize summary statistics
summary_stats <- data.frame()

for (file in excel_files) {
  sheets <- excel_sheets(file)
  
  # Create a list to store processed sheets
  processed_sheets <- list()
  
  for (sheet in sheets) {
    data <- read_excel(file, sheet = sheet)
    
    if (ncol(data) > 0) {
      original_count <- nrow(data)
      
      # Filter out terms to be trimmed (case-insensitive)
      data_filtered <- data %>%
        filter(!tolower(pull(.[,1])) %in% terms_to_trim)
      
      processed_sheets[[sheet]] <- data_filtered
      
      # Update summary statistics
      summary_stats <- rbind(summary_stats, data.frame(
        File = basename(file),
        Sheet = sheet,
        Total_Terms = original_count,
        Trimmed_Terms = original_count - nrow(data_filtered),
        Remaining_Terms = nrow(data_filtered)
      ))
    }
  }
  
  # Save processed file
  write_xlsx(processed_sheets, 
             paste0("trimmed_excel_output/", basename(file)))
}

# 4. Report summary
cat("\nSummary Report:\n")
cat("Total unique AE terms loaded:", 
    nrow(ae_terms_1) + nrow(ae_terms_2), "\n")
cat("Terms marked for trimming:", 
    length(terms_to_trim), "\n\n")

# Display detailed summary
cat("File-wise summary:\n")
summary_output <- summary_stats %>%
      select(File, Sheet, Total_Terms, Trimmed_Terms, Remaining_Terms)
print(summary_output )

# # Save summary to a text file
# write.table(summary_output, 
#             file = "AE-term-trimming_summary_report.txt", 
#             sep = "\t", 
#             row.names = FALSE, 
#             quote = FALSE)

# cat("Summary report saved to 'Unique_AE_Terms/AE-term-trimming_summary_report_1208.txt'.\n")

write_xlsx(summary_output, 
            path = "Unique_AE_Terms/Unique_AE-term-trimming_summary_report_1208.xlsx")  # Save as an Excel file

cat("Summary report saved to 'Unique_AE-term-trimming_summary_report_1208.xlsx'.\n")