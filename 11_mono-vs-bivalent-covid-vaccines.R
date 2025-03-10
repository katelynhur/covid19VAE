################################################################################
#####################  Packages  ###############################################
################################################################################

# Install pacman for package management if not already installed
if (!requireNamespace("pacman", quietly = TRUE)) {
  install.packages("pacman")
}
library(pacman)

# Required CRAN packages
cran_packages <- c(
  "readxl", "writexl", "dplyr", "ggplot2", "ggVennDiagram", "tibble"
)

# Required Bioconductor packages
bioc_packages <- c("VennDetail")

# Automatically install and load CRAN and Bioconductor packages
p_load(char = cran_packages, install = TRUE)
p_load(char = bioc_packages, install = TRUE, repos = BiocManager::repositories())

################################################################################
#####################  Set Current Working Directory ###########################
################################################################################

current_path <- rstudioapi::getActiveDocumentContext()$path 
setwd(dirname(current_path))
print(current_path)
base_dir <- dirname(current_path)

################################################################################
#####################  Define Output Directory #################################
################################################################################

output_dir <- file.path(base_dir, "VennOutput")
dir.create(output_dir, showWarnings = FALSE, recursive = TRUE)

################################################################################
#####################  Function to Generate Venn Diagrams ######################
################################################################################

process_venn_diagrams <- function(file_path, prefix) {
  # Read data from Excel
  pfizer_data <- read_excel(file_path, sheet = "PFIZER-BIONTECH")
  pfizer_bivalent_data <- read_excel(file_path, sheet = "PFIZER-BIONTECH BIVALENT")
  moderna_data <- read_excel(file_path, sheet = "MODERNA")
  moderna_bivalent_data <- read_excel(file_path, sheet = "MODERNA BIVALENT")
  
  # Extract adverse events (AEs) columns
  pfizer_ae <- pfizer_data[[1]]
  pfizer_bivalent_ae <- pfizer_bivalent_data[[1]]
  moderna_ae <- moderna_data[[1]]
  moderna_bivalent_ae <- moderna_bivalent_data[[1]]
  
  # Prepare Venn lists
  pfizer_list <- list("\n\n\nPfizer\nMono" = pfizer_ae, "\n\n\nPfizer\nBi" = pfizer_bivalent_ae)
  moderna_list <- list("\n\n\nModerna\nMono" = moderna_ae, "\n\n\nModerna\nBi" = moderna_bivalent_ae)
  
  custom_colors <- c(
    PFIZER = "#023981",          # Blue
    PFIZER_BIVALENT = "#92C1FE", # Lighter Blue
    MODERNA = "#d62728",         # Red
    MODERNA_BIVALENT = "#ff9896",# Lighter Red
    JANSSEN = "#2ca02c",         # Green
    NOVAVAX = "#ffd700"          # Yellow
  )
  
  # Generate and save Venn diagrams using VennDetail
  pfizer_venn <- VennDetail::venndetail(pfizer_list)
  pdf(file.path(output_dir, paste0(prefix, "_PFIZER_VennDetail_Mono-vs-Bivalent.pdf")), width=10, height=8)
  plot(pfizer_venn, mycol=custom_colors[c("PFIZER","PFIZER_BIVALENT")],
       cat.cex = 3, alpha = 0.5, cex = 3.8, cat.fontface = "bold", margin = 0.1)
  dev.off()
  
  moderna_venn <- VennDetail::venndetail(moderna_list)
  pdf(file.path(output_dir, paste0(prefix, "_MODERNA_VennDetail_Mono-vs-Bivalent.pdf")), width=10, height=8)
  plot(moderna_venn, mycol=custom_colors[c("MODERNA","MODERNA_BIVALENT")],
       cat.cex = 3, alpha = 0.5, cex = 3.8, cat.fontface = "bold", margin = 0.1)
  dev.off()
  
  # Save intersection results
  pfizer_venn_detail <- result(pfizer_venn)
  pfizer_combined <- pfizer_venn_detail %>%
    left_join(pfizer_data, by = c("Detail" = "AE")) %>%
    left_join(pfizer_bivalent_data, by = c("Detail" = "AE"))
  write_xlsx(pfizer_combined, file.path(output_dir, paste0(prefix, "_PFIZER_Combined_Output_Mono-vs-Bivalent.xlsx")))
  
  moderna_venn_detail <- result(moderna_venn)
  moderna_combined <- moderna_venn_detail %>%
    left_join(moderna_data, by = c("Detail" = "AE")) %>%
    left_join(moderna_bivalent_data, by = c("Detail" = "AE"))
  write_xlsx(moderna_combined, file.path(output_dir, paste0(prefix, "_MODERNA_Combined_Output_Mono-vs-Bivalent.xlsx")))
}

################################################################################
#####################  Process All Files #######################################
################################################################################

# # Define file paths
# files <- list(
#   OUTPUT9 = file.path(base_dir, "excel_output", "OUTPUT9.xlsx")
# )

# Define file paths for Original and Trimmed
files <- list(
  NonTrimmed = file.path(base_dir, "excel_output", "OUTPUT9.xlsx"),
  Trimmed = file.path(base_dir, "trimmed_excel_output", "OUTPUT9.xlsx")
)

# Process each file
for (file_name in names(files)) {
  process_venn_diagrams(files[[file_name]], file_name)
}


