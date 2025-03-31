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
  "ggVennDiagram", "tibble"
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
    Pfizer = pfizer_ae,
    Moderna = moderna_ae,
    Janssen = janssen_ae
  )
  
  # Define custom colors for each set
  custom_colors <- c(
    Pfizer = "#023981",  # Blue
    Moderna = "#d62728", # Red
    Janssen = "#2ca02c"  # Green
  )
  
  # Generate Venn diagram using ggVennDiagram
  venn_plot <- ggVennDiagram(
    venn_list,
    label = "count",
    label_alpha = 0,
    set_color = custom_colors
  ) +
    scale_fill_gradient(low = "white", high = "red") +
    theme(legend.position = "right")
    # theme(legend.position = "right",
    #       text = element_text(size = 16),        # Overall text size
    #       plot.title = element_text(size = 18, face = "bold"),  # Title font size
    #       axis.text = element_text(size = 14),   # Axis text (if applicable)
    #       legend.text = element_text(size = 14), # Legend text size
    #       strip.text = element_text(size = 16)   # Set labels size
    # )

  # Save the Venn plot
  ggsave(
    filename = file.path(output_dir, paste0(prefix, "_VennDiagram_ggVennDiagram.pdf")),
    plot = venn_plot,
    width = 10, height = 8
  )
  
  # Create Venn diagram using VennDetail
  three_way_venn <- VennDetail::venndetail(venn_list)
  
  pdf(file.path(output_dir, paste0(prefix, "_VennDiagram_VennDetail.pdf")), width = 10, height = 8)
  plot(
    three_way_venn,
    mycol = custom_colors,  # Set fill colors
    #col = "black",#custom_colors,    # Set border colors
    col = custom_colors,    # Set border colors
    cat.cex = 3,          # Category label size
    # cat.pos = c(-30, 30, 180),  # Adjust angles of category labels
    # cat.dist = c(0.1, 0.1, 0.1),  # Distance of category labels from circles
    alpha = 0.5,            # Transparency level
    cex = 3.8,                # Intersection label size
    cat.fontface = "bold",  # Category label font face
    margin = 0.1           # Margin around the diagram
  )
  dev.off()
  
  # Save the intersection results
  venn_results <- result(three_way_venn)
  combined_results <- venn_results %>%
    left_join(pfizer_data, by = c("Detail" = "AE")) %>%
    left_join(moderna_data, by = c("Detail" = "AE")) %>%
    left_join(janssen_data, by = c("Detail" = "AE"))
  write_xlsx(combined_results, file.path(output_dir, paste0(prefix, "_VennResults.xlsx")))
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
