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
  "readxl", "writexl", "tidyverse", "ggplot2", "ComplexUpset", 
  "ggVennDiagram", "tibble"  # Added ggVennDiagram
)

# Required Bioconductor packages
bioc_packages <- c(
  "VennDetail"
)

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

process_venn_diagrams <- function(file_path, prefix, vaccines) {
  # Read data from Excel
  pfizer_data <- read_excel(file_path, sheet = "PFIZER-BIONTECH")
  moderna_data <- read_excel(file_path, sheet = "MODERNA")
  janssen_data <- read_excel(file_path, sheet = "JANSSEN")
  
  # Extract adverse events (AEs) columns
  pfizer_ae <- pfizer_data[[1]]
  moderna_ae <- moderna_data[[1]]
  janssen_ae <- janssen_data[[1]]
  
  # Prepare Venn list
  venn_list <- list(
    PFIZER = pfizer_ae,
    MODERNA = moderna_ae,
    JANSSEN = janssen_ae
  )
  
  # Generate Venn diagram using ggVennDiagram
  venn_plot <- ggVennDiagram(
    venn_list,
    label = "count",
    label_alpha = 0,
    set_color = c(
      PFIZER = "#1f77b4",  # Blue
      MODERNA = "#d62728", # Red
      JANSSEN = "#2ca02c"  # Green
    )
  ) +
    scale_fill_gradient(low = "white", high = "red") +
    theme(legend.position = "right")
  
  # Save the Venn plot
  ggsave(
    filename = file.path(output_dir, paste0(prefix, "_VennDiagram_ggVennDiagram.pdf")),
    plot = venn_plot,
    width = 10, height = 8
  )
  
  # Create Venn diagram using VennDetail
  six_way_venn <- VennDetail::venndetail(venn_list)
  
  pdf(file.path(output_dir, paste0(prefix, "_VennDiagram_VennDetail.pdf")), width = 12, height = 8)
  plot(six_way_venn)
  dev.off()
  
  # Save the intersection results
  venn_results <- result(six_way_venn)
  write_xlsx(venn_results, file.path(output_dir, paste0(prefix, "_VennResults.xlsx")))
}

################################################################################
#####################  Process All Files #######################################
################################################################################

# Define file paths
files <- list(
  Trimmed_OUTPUT6 = file.path(base_dir, "trimmed_excel_output", "OUTPUT6.xlsx"),
  Trimmed_OUTPUT9 = file.path(base_dir, "trimmed_excel_output", "OUTPUT9.xlsx"),
  NonTrimmed_OUTPUT6 = file.path(base_dir, "excel_output", "OUTPUT6.xlsx"),
  NonTrimmed_OUTPUT9 = file.path(base_dir, "excel_output", "OUTPUT9.xlsx")
)

# Process each file
for (file_name in names(files)) {
  process_venn_diagrams(files[[file_name]], file_name, c("PFIZER", "MODERNA", "JANSSEN"))
}
